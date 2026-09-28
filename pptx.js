// Сборка PowerPoint с итоговой таблицей по шаблону «Итоги квиза» (без зависимостей)
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const ROUNDS = 7;
const EMU = 914400;
const TPL = JSON.parse(fs.readFileSync(path.join(__dirname, 'pptx-template.json'), 'utf8'));
const HEAD = Buffer.from(TPL.__slideHead, 'base64').toString('utf8');

// ---------- счёт и места (те же правила, что на сайте) ----------
const num = v => { const n = parseFloat(String(v).replace(',', '.')); return isFinite(n) ? n : 0; };
const total = t => t.scores.reduce((s, v) => s + num(v), 0);
const fmt = n => (Math.round(n * 100) / 100).toString().replace('.', ',');
function compare(a, b) {
  const d = total(b) - total(a);
  if (Math.abs(d) > 1e-9) return d;
  for (let r = ROUNDS - 1; r >= 0; r--) {
    const x = num(b.scores[r]) - num(a.scores[r]);
    if (Math.abs(x) > 1e-9) return x;
  }
  return 0;
}
function ranked(teams) {
  const list = teams.map((t, i) => ({ t, label: (t.name || '').trim() || 'Команда ' + (i + 1) })).sort((a, b) => compare(a.t, b.t));
  let place = 0;
  list.forEach((x, i) => { if (i === 0 || compare(list[i - 1].t, x.t) !== 0) place = i + 1; x.place = place; });
  return list;
}

// ---------- XML-кирпичики ----------
const E = inches => Math.round(inches * EMU);
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const STYLE = '<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>';
const xfrm = (x, y, w, h) => `<a:xfrm><a:off x="${E(x)}" y="${E(y)}"/><a:ext cx="${E(w)}" cy="${E(h)}"/></a:xfrm>`;
const run = (text, sz, bold, color, face) =>
  `<a:r><a:rPr lang="ru-RU" sz="${Math.round(sz * 100)}" b="${bold ? 1 : 0}"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:latin typeface="${face}"/><a:cs typeface="${face}"/></a:rPr><a:t>${esc(text)}</a:t></a:r>`;

function textBox(id, x, y, w, h, text, sz, bold, color, face, algn) {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="TextBox ${id - 1}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(x, y, w, h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="none" lIns="0" rIns="0" tIns="0" bIns="0" anchor="ctr"><a:noAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="${algn}"/>${run(text, sz, bold, color, face)}</a:p></p:txBody></p:sp>`;
}
function shape(id, name, prst, adj, x, y, w, h, fillXml, lnXml, text) {
  const av = adj ? `<a:avLst><a:gd name="adj" fmla="val ${adj}"/></a:avLst>` : '<a:avLst/>';
  const body = text
    ? `<p:txBody><a:bodyPr rtlCol="0" wrap="none" anchor="ctr" lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/><a:p><a:pPr algn="ctr"/>${text}</a:p></p:txBody>`
    : '<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p/></p:txBody>';
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name} ${id - 1}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(x, y, w, h)}<a:prstGeom prst="${prst}">${av}</a:prstGeom>${fillXml}${lnXml}<a:effectLst/></p:spPr>${STYLE}${body}</p:sp>`;
}
const solid = c => `<a:solidFill><a:srgbClr val="${c}"/></a:solidFill>`;
const line = c => `<a:ln w="12700">${solid(c)}</a:ln>`;
const noLine = '<a:ln><a:noFill/></a:ln>';

const CIRCLE = { 1: ['F2C037', '2A1E00'], 2: ['B4BDD4', '141A2E'], 3: ['D98A55', 'FFFFFF'] };
const COLS_X = [5.55, 6.33, 7.11, 7.89, 8.67, 9.45, 10.23];

// одна строка таблицы = группа из 11 фигур, как в шаблоне
function rowGroup(gid, x, y, pitch, f) {
  const place = x.place, top3 = place <= 3;
  const fs = Math.min(1, f * 1.25); // шрифт уменьшается мягче, чем строки
  const h = 0.48 * f, X0 = 0.6, W = 12.13;
  let id = gid + 1, parts = '';
  const bg = place === 1
    ? '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="3B3317"/></a:gs><a:gs pos="70000"><a:srgbClr val="161B31"/></a:gs></a:gsLst><a:lin scaled="0" ang="0"/></a:gradFill>'
    : solid('161B31');
  parts += shape(id++, 'Rounded Rectangle', 'roundRect', 22000, X0, y, W, h, bg, line(place === 1 ? 'F2C037' : '2A3152'));
  const [cf, ct] = CIRCLE[place] || ['1E2440', '9AA3C2'];
  const d = 0.38 * f;
  const digits = String(place).length;
  const circleSz = Math.min(13 * fs, (d * 72 * 0.9) / (digits * 0.62 + 0.2));
  parts += shape(id++, 'Oval', 'ellipse', null, 0.83, y + (h - d) / 2, d, d, solid(cf), noLine, run(String(place), Math.max(5, circleSz), false, ct, 'Russo One'));
  // длинные названия уменьшаем, чтобы влезли в одну строку
  const nameSz = Math.max(6, Math.min(17 * fs, (3.9 * 72) / (x.label.length * 0.66)));
  parts += textBox(id++, 1.5, y, 3.9, h, x.label, nameSz, true, 'EEF1FB', 'Golos Text', 'l');
  const scColor = top3 ? 'C9CFE3' : '9AA3C2';
  x.t.scores.forEach((v, r) => {
    parts += textBox(id++, COLS_X[r], y, 0.78, h, v === '' ? '–' : fmt(num(v)), Math.max(7, 15 * fs), false, scColor, 'Golos Text', 'ctr');
  });
  const ph = 0.34 * f;
  const [pf, pt] = place === 1 ? ['F2C037', '2A1E00'] : ['6F8BFF', '0E1222'];
  parts += shape(id++, 'Rounded Rectangle', 'roundRect', 30000, 11.32, y + (h - ph) / 2, 1.14, ph, solid(pf), noLine, run(fmt(total(x.t)), Math.max(7, Math.min(18 * fs, ph * 72 * 0.8)), false, pt, 'Russo One'));
  const off = `<a:off x="${E(X0)}" y="${E(y)}"/><a:ext cx="${E(W)}" cy="${E(h)}"/>`;
  return `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${gid}" name="Место ${x.rowIndex + 1}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm>${off}<a:chOff x="${E(X0)}" y="${E(y)}"/><a:chExt cx="${E(W)}" cy="${E(h)}"/></a:xfrm></p:grpSpPr>${parts}</p:grpSp>`;
}

// анимация как в шаблоне: по щелчку строки появляются снизу вверх, первое место — последним
function timing(groupIds) {
  let c = 2, pars = '';
  groupIds.forEach((spid, k) => {
    const dur = k === groupIds.length - 1 ? 1600 : 1200;
    const a = ++c, b = ++c, e1 = ++c, s1 = ++c, t1 = ++c, x1 = ++c, y1 = ++c;
    pars += `<p:par><p:cTn id="${a}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${b}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${e1}" presetID="42" presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>`
      + `<p:set><p:cBhvr><p:cTn id="${s1}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>`
      + `<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="${t1}" dur="${dur}"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:animEffect>`
      + `<p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="${x1}" dur="${dur}" decel="100000" fill="hold"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav></p:tavLst></p:anim>`
      + `<p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="${y1}" dur="${dur}" decel="100000" fill="hold"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="#ppt_y+.1"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="#ppt_y"/></p:val></p:tav></p:tavLst></p:anim>`
      + `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`;
  });
  if (!groupIds.length) return '';
  return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${pars}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
}

function slideXml(teams) {
  // пустые карточки (без названия и без баллов) на слайд не попадают
  const list = ranked(teams.filter(t => (t.name || '').trim() || t.scores.some(v => v !== '')));
  const n = list.length;
  const TOP = 1.85, BOTTOM = 7.3;
  const pitch = n ? Math.min(0.54, (BOTTOM - TOP) / n) : 0.54;
  const f = pitch / 0.54;
  let rows = '', ids = [];
  list.forEach((x, i) => {
    x.rowIndex = i;
    const gid = 18 + i * 12;
    ids.push(gid);
    rows += rowGroup(gid, x, TOP + i * pitch, pitch, f);
  });
  return HEAD + rows + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' + timing(ids.reverse()) + '</p:sld>';
}

// ---------- минимальный ZIP ----------
const CRC_TABLE = (() => { const t = new Int32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c; } return t; })();
function crc32(buf) { let c = -1; for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xFF] ^ (c >>> 8); return (c ^ -1) >>> 0; }
function zip(files) {
  const locals = [], centrals = [];
  let offset = 0;
  for (const [name, data] of files) {
    const nameBuf = Buffer.from(name, 'utf8');
    const comp = zlib.deflateRawSync(data);
    const crc = crc32(data);
    const lh = Buffer.alloc(30);
    lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(20, 4); lh.writeUInt16LE(0x0800, 6); lh.writeUInt16LE(8, 8);
    lh.writeUInt16LE(0, 10); lh.writeUInt16LE(0x21, 12); lh.writeUInt32LE(crc, 14);
    lh.writeUInt32LE(comp.length, 18); lh.writeUInt32LE(data.length, 22); lh.writeUInt16LE(nameBuf.length, 26); lh.writeUInt16LE(0, 28);
    locals.push(lh, nameBuf, comp);
    const ch = Buffer.alloc(46);
    ch.writeUInt32LE(0x02014b50, 0); ch.writeUInt16LE(20, 4); ch.writeUInt16LE(20, 6); ch.writeUInt16LE(0x0800, 8); ch.writeUInt16LE(8, 10);
    ch.writeUInt16LE(0, 12); ch.writeUInt16LE(0x21, 14); ch.writeUInt32LE(crc, 16); ch.writeUInt32LE(comp.length, 20); ch.writeUInt32LE(data.length, 24);
    ch.writeUInt16LE(nameBuf.length, 28); ch.writeUInt32LE(offset, 42);
    centrals.push(ch, nameBuf);
    offset += 30 + nameBuf.length + comp.length;
  }
  const cd = Buffer.concat(centrals);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(files.length, 8); end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(cd.length, 12); end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, cd, end]);
}

function buildPptx(teams) {
  const files = [];
  for (const [name, b64] of Object.entries(TPL)) {
    if (name === '__slideHead') continue;
    files.push([name, Buffer.from(b64, 'base64')]);
    if (name === 'ppt/slides/_rels/slide1.xml.rels') files.push(['ppt/slides/slide1.xml', Buffer.from(slideXml(teams), 'utf8')]);
  }
  if (!files.some(f => f[0] === 'ppt/slides/slide1.xml')) files.push(['ppt/slides/slide1.xml', Buffer.from(slideXml(teams), 'utf8')]);
  return zip(files);
}

module.exports = { buildPptx };
