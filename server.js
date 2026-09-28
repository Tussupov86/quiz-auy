// АЮ READY КВИЗ — сервер счёта с живой синхронизацией (без зависимостей)
const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const PORT = process.env.PORT || 3000;
const DATA_FILE = path.join(__dirname, 'data.json');
const ROUNDS = 7;
const BOOT_ID = crypto.randomBytes(6).toString('hex');

let state = { rev: 0, fresh: true, teams: [], updatedAt: 0 };
try {
  const saved = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
  if (saved && Array.isArray(saved.teams)) state = { ...saved, fresh: false };
} catch (e) { /* файла нет — начинаем с пустого счёта */ }

function cleanTeams(teams) {
  if (!Array.isArray(teams)) return null;
  return teams.slice(0, 60).map(t => ({
    id: String(t && t.id || crypto.randomBytes(4).toString('hex')).slice(0, 32),
    name: String(t && t.name || '').slice(0, 80),
    scores: Array.from({ length: ROUNDS }, (_, r) =>
      String((t && Array.isArray(t.scores) && t.scores[r] != null) ? t.scores[r] : '')
        .replace(/[^0-9.,\-]/g, '').slice(0, 8)),
  }));
}

let saveTimer = null;
function persist() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    fs.writeFile(DATA_FILE, JSON.stringify(state), () => {});
  }, 300);
}

const clients = new Set();
function publicState() { return { ...state, boot: BOOT_ID }; }
function broadcast(by) {
  const msg = `data: ${JSON.stringify({ ...publicState(), by })}\n\n`;
  for (const res of clients) { try { res.write(msg); } catch (e) {} }
}
setInterval(() => { for (const res of clients) { try { res.write(': ping\n\n'); } catch (e) {} } }, 25000);

function send(res, code, body, type = 'application/json; charset=utf-8') {
  res.writeHead(code, { 'Content-Type': type, 'Cache-Control': 'no-store' });
  res.end(typeof body === 'string' ? body : JSON.stringify(body));
}

const INDEX = fs.existsSync(path.join(__dirname, 'public', 'index.html'))
  ? path.join(__dirname, 'public', 'index.html')
  : path.join(__dirname, 'index.html');
const { buildPptx } = require('./pptx');

const server = http.createServer((req, res) => {
  const url = new URL(req.url, 'http://x');

  if (req.method === 'GET' && (url.pathname === '/' || url.pathname === '/index.html')) {
    return fs.readFile(INDEX, (err, buf) => err ? send(res, 500, 'error', 'text/plain') : send(res, 200, buf.toString('utf8'), 'text/html; charset=utf-8'));
  }

  if (req.method === 'GET' && url.pathname === '/api/state') return send(res, 200, publicState());

  if (req.method === 'GET' && url.pathname === '/api/pptx') {
    try {
      const buf = buildPptx(state.teams);
      const d = new Date(Date.now() + 6 * 3600 * 1000); // время Бишкека
      const stamp = String(d.getUTCDate()).padStart(2, '0') + '-' + String(d.getUTCMonth() + 1).padStart(2, '0');
      const name = `Итоги_квиза_${stamp}.pptx`;
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': `attachment; filename="quiz_${stamp}.pptx"; filename*=UTF-8''${encodeURIComponent(name)}`,
        'Content-Length': buf.length,
        'Cache-Control': 'no-store',
      });
      return res.end(buf);
    } catch (e) {
      console.error(e);
      return send(res, 500, 'Не удалось собрать файл', 'text/plain; charset=utf-8');
    }
  }

  if (req.method === 'GET' && url.pathname === '/api/events') {
    res.writeHead(200, {
      'Content-Type': 'text/event-stream; charset=utf-8',
      'Cache-Control': 'no-store',
      'Connection': 'keep-alive',
      'X-Accel-Buffering': 'no',
    });
    res.write(`retry: 2000\ndata: ${JSON.stringify(publicState())}\n\n`);
    clients.add(res);
    req.on('close', () => clients.delete(res));
    return;
  }

  if (req.method === 'POST' && url.pathname === '/api/state') {
    let body = '';
    req.on('data', c => { body += c; if (body.length > 200000) req.destroy(); });
    req.on('end', () => {
      let data;
      try { data = JSON.parse(body); } catch (e) { return send(res, 400, { error: 'bad_json' }); }
      const teams = cleanTeams(data.teams);
      if (!teams) return send(res, 400, { error: 'bad_teams' });
      state = { rev: state.rev + 1, fresh: false, teams, updatedAt: Date.now() };
      persist();
      broadcast(String(data.by || '').slice(0, 40));
      send(res, 200, { ok: true, rev: state.rev });
    });
    return;
  }

  send(res, 404, 'Not found', 'text/plain; charset=utf-8');
});

server.listen(PORT, () => console.log(`Quiz server on :${PORT}`));
