/* XLSnip v2026.10.07c – complément Excel d'extraction de données depuis des documents (PDF / images) */
const NS_DOC = 'urn:xlsnip:doc', NS_SNIP = 'urn:xlsnip:snips', NS_ORG = 'urn:xlsnip:organization';
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

const S = { docs: [], cur: null, pdf: null, page: 1, zoom: 1.3, mode: 'text', snips: [], focus: null, task: null, ocr: null, busy: false, rf: false, again: false, lastRf: 0, folders: [], docFolder: (() => { try { return localStorage.getItem('xlsnip.docFolder') || 'all'; } catch (e) { return 'all'; } })(), orgSel: 'root', orgDocSel: null, batch:{active:false, folderId:null, orientation:'vertical', docs:[], startRow:0, startCol:0, sheetName:'', page:1}, search: { q: '', matches: [], index: -1 } };
const $ = id => document.getElementById(id);
const say = t => { $('st').textContent = t; };
const esc = s => s.replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const b64 = u => { let s = ''; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); };
const unb64 = s => { const b = atob(s), u = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); return u; };
const jb64 = o => btoa(unescape(encodeURIComponent(JSON.stringify(o))));
const bj64 = s => JSON.parse(decodeURIComponent(escape(atob(s))));
const r1 = x => Math.round(x * 10) / 10;
const escT = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const fmtSize = n => n >= 1048576 ? (n / 1048576).toFixed(1) + ' Mo' : Math.max(1, Math.round(n / 1024)) + ' Ko';

/* ---------- Persistance dans le classeur (parties XML personnalisées) ---------- */
async function saveDoc(d) {
  await Excel.run(async c => {
    const xml = `<d xmlns="${NS_DOC}" id="${d.id}" name="${esc(d.name)}" folder="${esc(d.folder || 'root')}" wf="json"><f>${b64(d.data)}</f><w>${escT(JSON.stringify(d.words))}</w></d>`;
    const p = c.workbook.customXmlParts.add(xml); p.load('id'); await c.sync(); d.pid = p.id;
  });
}
async function saveSnips() {
  await Excel.run(async c => {
    const sp = c.workbook.customXmlParts.getByNamespace(NS_SNIP); sp.load('items'); await c.sync();
    sp.items.forEach(p => p.delete());
    c.workbook.customXmlParts.add(`<s xmlns="${NS_SNIP}">${jb64(S.snips)}</s>`); await c.sync();
  });
}
async function loadAll() {
  await Excel.run(async c => {
    const parts = c.workbook.customXmlParts.getByNamespace(NS_DOC); parts.load('items'); await c.sync();
    const xs = parts.items.map(p => p.getXml());
    const sp = c.workbook.customXmlParts.getByNamespace(NS_SNIP); sp.load('items'); await c.sync();
    const ss = sp.items.map(p => p.getXml()); await c.sync();
    const op = c.workbook.customXmlParts.getByNamespace(NS_ORG); op.load('items'); await c.sync();
    const os = op.items.map(p => p.getXml()); await c.sync();
    S.folders = os.length ? bj64(new DOMParser().parseFromString(os[0].value, 'text/xml').documentElement.textContent) : [{id:'root',name:'Documents',parent:null}];
    if (!S.folders.length) S.folders = [{id:'root',name:'Documents',parent:null}];
    S.docs = xs.map((x, i) => {
      const e = new DOMParser().parseFromString(x.value, 'text/xml').documentElement;
      return { pid: parts.items[i].id, id: e.getAttribute('id'), name: e.getAttribute('name'), folder: e.getAttribute('folder') || 'root',
        data: unb64(e.getElementsByTagName('f')[0].textContent), words: e.getAttribute('wf') === 'json' ? JSON.parse(e.getElementsByTagName('w')[0].textContent) : bj64(e.getElementsByTagName('w')[0].textContent) };
    });
    S.snips = ss.length ? bj64(new DOMParser().parseFromString(ss[0].value, 'text/xml').documentElement.textContent) : [];
  });
}

/* ---------- Suivi des snips : plages nommées masquées (suivent insertions/suppressions de lignes et colonnes) ---------- */
const NM = 'XLSnip_';
const newId = () => Date.now().toString(36) + Math.floor(Math.random() * 1296).toString(36);
let rfTimer = null;
function schedRefresh() { clearTimeout(rfTimer); rfTimer = setTimeout(refreshSnips, 400); }
async function refreshSnips() {
  if (S.busy || S.rf) { S.again = true; return; }
  S.rf = true; S.lastRf = Date.now();
  let changed = false;
  try {
    await Excel.run(async c => {
      const names = c.workbook.names; names.load('items/name'); await c.sync();
      const have = new Set(names.items.map(n => n.name));
      // Migration : snips créés avant l'ajout du suivi (sans plage nommée)
      for (const s of S.snips.filter(s => !s.n)) {
        const w = c.workbook.worksheets.getItemOrNullObject(s.sh); await c.sync();
        if (w.isNullObject) continue;
        s.n = NM + newId(); names.add(s.n, w.getRange(s.a)).visible = false; have.add(s.n); changed = true;
      }
      await c.sync();
      const rs = S.snips.map(s => {
        if (!s.n || !have.has(s.n)) return null;
        const rg = names.getItem(s.n).getRangeOrNullObject(); rg.load('address,values'); return rg;
      });
      await c.sync();
      const out = [];
      S.snips.forEach((s, i) => {
        const rg = rs[i]; if (!rg || rg.isNullObject) return;          // lignes/colonnes/feuille supprimées
        const k = rg.address.lastIndexOf('!');
        const sh = rg.address.slice(0, k).replace(/^'|'$/g, '').replace(/''/g, "'"), a = rg.address.slice(k + 1).replace(/\$/g, '');
        if (rg.values.every(row => row.every(v => v === '' || v === null))) { resetFmt(rg, s); return; }   // contenu effacé : on retire aussi le formatage
        if (s.sh !== sh || s.a !== a) { s.sh = sh; s.a = a; changed = true; }
        out.push(s);
      });
      if (out.length !== S.snips.length) changed = true;
      const keep = new Set(out.map(s => s.n));
      names.items.forEach(n => { if (n.name.startsWith(NM) && !keep.has(n.name)) n.delete(); });
      S.snips = out; await c.sync();
    });
    if (changed) { await saveSnips(); if (S.cur && S.pdf) { if (S.focusN && !S.snips.some(s => s.n === S.focusN)) S.focusN = null; drawMarks(); } }
  } catch (e) { /* on réessaiera au prochain événement */ }
  finally { S.rf = false; if (S.again) { S.again = false; schedRefresh(); } }
}

/* ---------- Analyse des documents (texte natif + OCR si besoin) ---------- */
async function imgToPdf(f) {
  const url = URL.createObjectURL(f);
  const im = await new Promise((ok, ko) => { const i = new Image(); i.onload = () => ok(i); i.onerror = ko; i.src = url; });
  const cv = document.createElement('canvas'); cv.width = im.naturalWidth; cv.height = im.naturalHeight;
  cv.getContext('2d').drawImage(im, 0, 0);
  const p = new window.jspdf.jsPDF({ unit: 'px', format: [cv.width, cv.height], orientation: cv.width > cv.height ? 'l' : 'p', hotfixes: ['px_scaling'] });
  p.addImage(cv.toDataURL('image/jpeg', 0.9), 'JPEG', 0, 0, cv.width, cv.height);
  URL.revokeObjectURL(url);
  return new Uint8Array(p.output('arraybuffer'));
}
async function ocrPage(pg) {
  const vp = pg.getViewport({ scale: 2 });
  const c = document.createElement('canvas'); c.width = vp.width; c.height = vp.height;
  await pg.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
  S.ocr = S.ocr || await Tesseract.createWorker('fra+eng');
  const { data } = await S.ocr.recognize(c);
  return data.words.filter(w => w.text.trim() && w.confidence > 30)
    .map(w => [w.text, r1(w.bbox.x0 / 2), r1(w.bbox.y0 / 2), r1(w.bbox.x1 / 2), r1(w.bbox.y1 / 2)]);
}
async function analyse(pdf, name) {
  const all = [];
  for (let n = 1; n <= pdf.numPages; n++) {
    say(`Analyse de « ${name} » : page ${n}/${pdf.numPages}…`);
    const pg = await pdf.getPage(n), vp = pg.getViewport({ scale: 1 }), tc = await pg.getTextContent(), ws = [];
    tc.items.forEach(it => {
      if (!it.str || !it.str.trim()) return;
      const t = pdfjsLib.Util.transform(vp.transform, it.transform), h = Math.hypot(t[2], t[3]), w = it.width * vp.scale, len = it.str.length;
      for (const m of it.str.matchAll(/\S+/g)) {
        const x0 = t[4] + w * m.index / len, x1 = t[4] + w * (m.index + m[0].length) / len;
        ws.push([m[0], r1(x0), r1(t[5] - h), r1(x1), r1(t[5])]);
      }
    });
    if (ws.length < 3) { say(`OCR de la page ${n}/${pdf.numPages} (peut prendre un moment)…`); all.push(await ocrPage(pg)); }
    else all.push(ws);
  }
  return all;
}
async function importFile(f, options = {}) {
  let buf;
  if (/^image\//.test(f.type)) buf = await imgToPdf(f);
  else if (/\.pdf$/i.test(f.name) || f.type === 'application/pdf') buf = new Uint8Array(await f.arrayBuffer());
  else { say(`« ${f.name} » : format non pris en charge. Exportez-le d'abord en PDF.`); return; }
  const pdf = await pdfjsLib.getDocument({ data: buf.slice() }).promise;
  const d = { id: 'd' + Date.now() + Math.floor(Math.random() * 1000), name: f.name, folder: options.folder || (S.docFolder !== 'all' ? S.docFolder : 'root'), data: buf, words: await analyse(pdf, f.name) };
  say('Enregistrement dans le classeur…');
  await saveDoc(d); S.docs.push(d); d.pdfObj = pdf; S.search.matches = buildSearchMatches(S.search.q);
  fillDocs(); if(options.open !== false) await openDoc(d.id); else renderOrgTree();
  say(`« ${f.name} » importé (${pdf.numPages} page(s)).`);
}

/* ---------- Visionneuse ---------- */
/* Sélecteur de dossier : filtre la liste des documents */
const saveDocFolder = () => { try { localStorage.setItem('xlsnip.docFolder', S.docFolder); } catch (e) {} };
function folderDescendants(id) {   // le dossier + tous ses sous-dossiers (récursif)
  const out = new Set([id]); let added = true;
  while (added) { added = false; S.folders.forEach(f => { if (out.has(f.parent) && !out.has(f.id)) { out.add(f.id); added = true; } }); }
  return out;
}
function folderOptionList() {
  const out = [], seen = new Set();
  const walk = (id, depth) => {
    const f = S.folders.find(x => x.id === id); if (!f || seen.has(id)) return; seen.add(id);
    const ids = folderDescendants(id), n = S.docs.filter(d => ids.has(d.folder || 'root')).length;
    out.push({ id, label: '\u00a0\u00a0'.repeat(depth) + (depth ? '\u21b3 ' : '') + '\ud83d\udcc1 ' + f.name + ' (' + n + ')' });
    S.folders.filter(c => c.parent === id).forEach(c => walk(c.id, depth + 1));
  };
  walk('root', 0); S.folders.forEach(f => walk(f.id, 1));   // dossiers orphelins éventuels
  return out;
}
const docsInView = () => { if (S.docFolder === 'all') return S.docs; const ids = folderDescendants(S.docFolder); return S.docs.filter(d => ids.has(d.folder || 'root')); };
function fillFolderSelect() {
  const sel = $('docFolder'); if (!sel) return;
  const opts = folderOptionList();
  if (S.docFolder !== 'all' && !opts.some(o => o.id === S.docFolder)) S.docFolder = 'all';
  sel.innerHTML = `<option value="all">Tous les dossiers (${S.docs.length})</option>` + opts.map(o => `<option value="${esc(o.id)}">${esc(o.label)}</option>`).join('');
  sel.value = S.docFolder;
}
function fillDocs() {
  fillFolderSelect();
  const list = docsInView();
  const base = S.docFolder === 'all' ? 'root' : S.docFolder;   // les documents venant d'un sous-dossier sont préfixés du nom de ce sous-dossier
  const pre = d => { const fid = d.folder || 'root'; if (fid === base) return ''; const f = S.folders.find(x => x.id === fid); return f ? f.name + ' / ' : ''; };
  $('docs').innerHTML = list.length ? list.map(d => `<option value="${d.id}">${esc(pre(d) + d.name)} (${fmtSize(d.data.length)})</option>`).join('') : '<option value="">(aucun document)</option>';
  if (S.cur && list.includes(S.cur)) $('docs').value = S.cur.id;
}
function syncDocSelect(d) {   // un document ouvert hors du dossier affiché : on bascule sur son dossier
  if (S.docFolder !== 'all' && !docsInView().includes(d)) { S.docFolder = d.folder || 'root'; saveDocFolder(); fillDocs(); }
  else $('docs').value = d.id;
}
async function openDoc(id, page) {
  const d = S.docs.find(x => x.id === id); if (!d) return;
  S.cur = d; d.pdfObj = d.pdfObj || await pdfjsLib.getDocument({ data: d.data.slice() }).promise;
  S.pdf = d.pdfObj; syncDocSelect(d); $('empty').hidden = true; $('wrap').hidden = false;
  if (!page) { const p = await S.pdf.getPage(1); S.zoom = Math.max(0.5, ($('view').clientWidth - 24) / p.getViewport({ scale: 1 }).width); }
  await showPage(page || 1);
}
async function showPage(n) {
  n = Math.min(Math.max(1, n), S.pdf.numPages); S.page = n;
  $('pg').value = n; $('pg').max = S.pdf.numPages; $('pn').textContent = '/ ' + S.pdf.numPages;
  if (S.task) { try { S.task.cancel(); } catch (e) {} }
  const pg = await S.pdf.getPage(n), vp = pg.getViewport({ scale: S.zoom }), cv = $('cv');
  cv.width = vp.width; cv.height = vp.height;
  S.task = pg.render({ canvasContext: cv.getContext('2d'), viewport: vp });
  try { await S.task.promise; } catch (e) { return; }
  drawMarks();
}
function drawMarks() {
  const ov = $('ov');
  ov.querySelectorAll('.mark,.searchMark').forEach(e => e.remove());
  const add = (r, cls) => {
    const e = document.createElement('div'); e.className = 'mark ' + cls; const z = S.zoom;
    Object.assign(e.style, { left: r[0] * z + 'px', top: r[1] * z + 'px', width: (r[2] - r[0]) * z + 'px', height: (r[3] - r[1]) * z + 'px' });
    ov.appendChild(e);
  };
  S.snips.filter(s => s.d === S.cur.id && s.p === S.page).forEach(s => add(s.r, (s.k || '') + (S.focusN && s.n === S.focusN ? ' f' : '')));

  const q = S.search.q.trim();
  if (!q) return;
  const hits = S.search.matches.filter(m => m.d === S.cur.id && m.p === S.page);
  hits.forEach((m, mi) => m.words.forEach((wi, j) => {
    const w = S.cur.words[S.page - 1]?.[wi]; if (!w) return;
    const e = document.createElement('div');
    e.className = 'searchMark' + (m === S.search.matches[S.search.index] ? ' current' : '');
    const z = S.zoom;
    Object.assign(e.style, { left: w[1] * z + 'px', top: w[2] * z + 'px', width: Math.max(2, (w[3] - w[1]) * z) + 'px', height: Math.max(2, (w[4] - w[2]) * z) + 'px' });
    ov.appendChild(e);
  }));
}

/* ---------- Recherche globale ---------- */
const normSearch = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLocaleLowerCase('fr');
const escRe = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

function buildSearchMatches(q) {
  q = normSearch(q.trim());
  if (!q) return [];
  const out = [];
  const re = new RegExp(escRe(q), 'gi');

  for (const d of S.docs) {
    (d.words || []).forEach((ws, pi) => {
      if (!ws || !ws.length) return;
      // On conserve une table caractère -> mot afin de pouvoir surligner
      // toute une expression, même si elle traverse plusieurs mots.
      let text = '', ranges = [];
      ws.forEach((w, wi) => {
        const token = normSearch(w[0]);
        if (!token) return;
        if (text) text += ' ';
        const a = text.length;
        text += token;
        ranges.push({ a, b: text.length, wi });
      });
      let m;
      while ((m = re.exec(text))) {
        const words = ranges.filter(r => r.b > m.index && r.a < m.index + m[0].length).map(r => r.wi);
        if (words.length) out.push({ d: d.id, p: pi + 1, words });
        if (m[0].length === 0) re.lastIndex++;
      }
    });
  }
  return out;
}

function updateSearchCount() {
  const c = $('searchCount');
  if (!S.search.q.trim()) { c.textContent = ''; return; }
  c.textContent = S.search.matches.length ? `${S.search.index + 1} / ${S.search.matches.length}` : 'Aucun';
}

async function goSearch(delta = 1, reset = false) {
  const q = $('search').value.trim();
  if (q !== S.search.q || reset) {
    S.search.q = q;
    S.search.matches = buildSearchMatches(q);
    S.search.index = S.search.matches.length ? (delta > 0 ? 0 : S.search.matches.length - 1) : -1;
  } else if (S.search.matches.length) {
    S.search.index = (S.search.index + delta + S.search.matches.length) % S.search.matches.length;
  }
  updateSearchCount();
  const m = S.search.matches[S.search.index];
  if (!m) { if (S.cur && S.pdf) drawMarks(); return; }
  if (!S.cur || S.cur.id !== m.d) await openDoc(m.d, m.p);
  else if (S.page !== m.p) await showPage(m.p);
  else drawMarks();
}

$('search').addEventListener('input', () => goSearch(1, true));
$('search').addEventListener('keydown', e => {
  if (e.key === 'Enter') { e.preventDefault(); goSearch(e.shiftKey ? -1 : 1); }
  else if (e.key === 'Escape') {
    $('search').value = ''; goSearch(1, true); $('search').blur();
  }
});
$('searchPrev').onclick = () => goSearch(-1);
$('searchNext').onclick = () => goSearch(1);


/* ---------- Extraction ---------- */
function lines(ws) {
  const hs = ws.map(w => w[4] - w[2]).sort((a, b) => a - b), mh = hs[hs.length >> 1] || 10, rows = [];
  [...ws].sort((a, b) => (a[2] + a[4]) - (b[2] + b[4])).forEach(w => {
    const cy = (w[2] + w[4]) / 2, r = rows.find(r => Math.abs(r.cy - cy) < mh * 0.6);
    if (r) { r.w.push(w); r.cy = (r.cy * (r.w.length - 1) + cy) / r.w.length; } else rows.push({ cy, w: [w] });
  });
  rows.sort((a, b) => a.cy - b.cy);
  return { rows: rows.map(r => r.w.sort((a, b) => a[1] - b[1])), mh };
}
function toGrid(ws) {
  const { rows, mh } = lines(ws), iv = ws.map(w => [w[1], w[3]]).sort((a, b) => a[0] - b[0]), cols = [];
  iv.forEach(v => { const l = cols[cols.length - 1]; if (l && v[0] - l[1] < mh * 0.7) l[1] = Math.max(l[1], v[1]); else cols.push([v[0], v[1]]); });
  return rows.map(r => {
    const row = cols.map(() => []);
    r.forEach(w => { const cx = (w[1] + w[3]) / 2; let i = cols.findIndex(c => cx >= c[0] && cx <= c[1]); if (i < 0) i = 0; row[i].push(w[0]); });
    return row.map(a => a.join(' '));
  });
}
function parseNum(t) {
  let s = t.replace(/[€$£%\s\u00a0]/g, ''); if (!/\d/.test(s)) return null;
  const neg = /^\(.*\)$/.test(s) || /^-/.test(s) || /-$/.test(s); s = s.replace(/[()\-]/g, '');
  const lc = s.lastIndexOf(','), ld = s.lastIndexOf('.');
  if (lc > -1 && ld > -1) s = lc > ld ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
  else if (lc > -1) s = (s.match(/,/g) || []).length > 1 ? s.replace(/,/g, '') : s.replace(',', '.');
  else if ((s.match(/\./g) || []).length > 1) s = s.replace(/\./g, '');
  const n = parseFloat(s); return isNaN(n) ? null : (neg ? -n : n);
}
const looksNum = t => /^[(\-]?[€$£]?\s*\d[\d\s\u00a0.,]*\s*[%€$£]?\)?-?$/.test(t.trim());
const safe = v => (typeof v === 'string' && v[0] === '=') ? "'" + v : v;

async function snip(r, cont, batchTarget = null) {
  if (S.mode === 'img') { await insertImage({ d: S.cur.id, p: S.page, r }, true); return; }   // mode Image : copie de la zone, sans lecture de texte
  await refreshSnips();
  const chk = S.mode === 'ok' || S.mode === 'nok';
  let grid, raw = null, num = null;
  if (chk) grid = [[S.mode === 'ok' ? '\u2713' : '\u2717']];
  else {
    const ws = (S.cur.words[S.page - 1] || []).filter(w => { const cx = (w[1] + w[3]) / 2, cy = (w[2] + w[4]) / 2; return cx >= r[0] && cx <= r[2] && cy >= r[1] && cy <= r[3]; });
    if (!ws.length) { say('Aucun texte détecté dans cette zone.'); return false; }
    const text = lines(ws).rows.map(l => l.map(w => w[0]).join(' ')).join(' ');
    if (S.mode === 'table') grid = toGrid(ws).map(row => row.map(t => (looksNum(t) && parseNum(t) !== null) ? parseNum(t) : safe(t)));
    else if (S.mode === 'num') { const n = parseNum(text); if (n === null) { say(`« ${text} » n'est pas un nombre.`); return; } grid = [[n]]; }
    else if (S.mode === 'sum') {
      const ns = toGrid(ws).flat().filter(t => looksNum(t) && parseNum(t) !== null).map(parseNum);
      if (!ns.length) { say('Aucun nombre dans cette zone.'); return; }
      num = ns.reduce((x, y) => x + y, 0); grid = [[num]];
    } else { raw = text; grid = [[safe(text)]]; }
  }
  S.busy = true; let msg = '';
  try {
    await Excel.run(async c => {
      const sh = batchTarget ? c.workbook.worksheets.getItem(batchTarget.sheetName) : c.workbook.worksheets.getActiveWorksheet(); sh.load('name');
      const act = batchTarget ? sh.getCell(batchTarget.rowIndex, batchTarget.columnIndex) : c.workbook.getActiveCell(); act.load('address,values'); await c.sync();
      const aA = act.address.slice(act.address.lastIndexOf('!') + 1).replace(/\$/g, '');
      // Ctrl : on poursuit le snip précédent si la cellule active est toujours celle-ci (modes Texte et Somme)
      let tg = null;
      if (cont && (S.mode === 'text' || S.mode === 'sum')) {
        const l = S.last && S.snips.find(s => s.n === S.last);
        if (l && l.m === S.mode && l.sh === sh.name && l.a === aA) tg = l;
      }
      if (tg) {
        const old = act.values[0][0];
        if (S.mode === 'sum') { const o = typeof old === 'number' ? old : (parseNum(String(old)) || 0); grid = [[Math.round((o + num) * 1e10) / 1e10]]; }
        else grid = [[safe(old === '' || old === null ? raw : old + ' ' + raw)]];
      }
      const rg = act.getResizedRange(grid.length - 1, grid[0].length - 1);
      rg.load('address'); rg.values = grid; await c.sync();
      const a = rg.address.slice(rg.address.lastIndexOf('!') + 1).replace(/\$/g, '');
      const prev = S.snips.find(s => s.sh === sh.name && s.a === a);
      if (chk) {
        const ok = S.mode === 'ok';
        rg.format.fill.color = ok ? '#C6EFCE' : '#FFC7CE';
        rg.format.font.color = ok ? '#006100' : '#9C0006';
        rg.format.font.bold = true; rg.format.horizontalAlignment = 'Center';
      } else {
        if (tg ? tg.f !== 0 : $('hl').checked) rg.format.fill.color = '#E2F0D9';
        if (prev && prev.k) { rg.format.font.color = '#000000'; rg.format.font.bold = false; rg.format.horizontalAlignment = 'General'; }
      }
      const nm = tg ? tg.n : NM + newId();
      if (!tg) c.workbook.names.add(nm, rg).visible = false;
      await c.sync();
      if (!tg) S.snips = S.snips.filter(s => !(s.sh === sh.name && s.a === a));
      S.snips.push(Object.assign({ sh: sh.name, a, d: S.cur.id, p: S.page, r: r.map(r1), n: nm, m: S.mode,
        f: tg ? tg.f : ((chk || $('hl').checked) ? 1 : 0) }, chk ? { k: S.mode } : {}));
      S.last = nm;
      msg = chk ? (S.mode === 'ok' ? 'Pointé : valide \u2713' : 'Pointé : invalide \u2717')
        : S.mode === 'table' ? `Tableau ${grid.length}×${grid[0].length} inséré.`
        : S.mode === 'sum' ? `Somme = ${grid[0][0]}${tg ? ' (zone ajoutée)' : ''}`
        : tg ? 'Zone ajoutée à la cellule.' : 'Valeur insérée.';
    });
    await saveSnips(); drawMarks(); say(msg); return true;
  } catch (e) { say('Erreur Excel : ' + e.message); return false; }
  finally { S.busy = false; schedRefresh(); }
}

/* ---------- Retrouver la source d'une cellule ---------- */
const colN = s => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
function parseA1(a) {
  const m = a.replace(/\$/g, '').match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/); if (!m) return null;
  return { c1: colN(m[1]), r1: +m[2], c2: colN(m[3] || m[1]), r2: +(m[4] || m[2]) };
}
async function onSel() {
  if (S.snips.length && Date.now() - S.lastRf > 3000) schedRefresh();
  if (!$('follow').checked || !S.snips.length) return;
  try {
    await Excel.run(async c => {
      const r = c.workbook.getSelectedRange(); r.load('address'); await c.sync();
      const i = r.address.lastIndexOf('!'), sh = r.address.slice(0, i).replace(/^'|'$/g, '').replace(/''/g, "'"), a = parseA1(r.address.slice(i + 1));
      if (!a) return;
      const s = S.snips.find(s => { const q = parseA1(s.a); return s.sh === sh && q && a.c1 >= q.c1 && a.c1 <= q.c2 && a.r1 >= q.r1 && a.r1 <= q.r2; });
      if (!s) { if (S.focusN) { S.focusN = null; if (S.cur && S.pdf) drawMarks(); } return; }
      S.focusN = s.n || null;
      await openDoc(s.d, s.p); $('view').scrollTop = Math.max(0, s.r[1] * S.zoom - 40);
      say('Source de la cellule affichée.');
    });
  } catch (e) { /* sélection non analysable : on ignore */ }
}

/* ---------- Interface ---------- */
$('imp').onclick = () => $('file').click();
$('file').onchange = async e => {
  for (const f of e.target.files) { try { await importFile(f); } catch (err) { say(`Échec de l'import de « ${f.name} » : ${err.message}`); } }
  e.target.value = '';
};
$('docs').onchange = e => { if (e.target.value) openDoc(e.target.value); };
// Garantit la présence du menu des dossiers (entre « Importer » et la liste des documents), même si taskpane.html est une ancienne version en cache
function ensureFolderSelect() {
  if ($('docFolder')) return;
  const s = document.createElement('select');
  s.id = 'docFolder'; s.setAttribute('aria-label', 'Dossier'); s.title = "Dossier d'où proviennent les documents de la liste";
  s.style.cssText = 'flex:0 1 36%;min-width:84px;max-width:100%';
  $('docs').parentNode.insertBefore(s, $('docs'));
}
ensureFolderSelect();
$('docFolder').onchange = async e => {
  S.docFolder = e.target.value; saveDocFolder(); fillDocs();
  const list = docsInView();
  if (!list.length) { say('Ce dossier ne contient aucun document.'); return; }
  if (!S.cur || !list.includes(S.cur)) await openDoc(list[0].id);
};
$('del').onclick = async () => {
  if (!S.cur || !(await ask('Retirer ce document ?', `<p>« ${esc(S.cur.name)} » et ses snips seront retirés du classeur.</p>`, 'Retirer'))) return;
  const d = S.cur;
  await Excel.run(async c => { c.workbook.customXmlParts.getItem(d.pid).delete(); await c.sync(); });
  S.docs = S.docs.filter(x => x !== d); S.snips = S.snips.filter(s => s.d !== d.id); S.search.matches = buildSearchMatches(S.search.q); S.search.index = Math.min(S.search.index, S.search.matches.length - 1); await saveSnips();
  S.cur = null; S.pdf = null; fillDocs();
  if (S.docs.length) await openDoc(S.docs[0].id); else { $('wrap').hidden = true; $('empty').hidden = false; $('pn').textContent = '/ –'; $('pg').value = ''; }
};
$('modes').onclick = e => { const m = e.target.dataset.m; if (!m) return; S.mode = m; [...$('modes').children].forEach(b => b.classList.toggle('on', b === e.target)); };
const goPage = () => {
  const v = parseInt($('pg').value, 10);
  if (!S.pdf) return;
  if (isNaN(v)) { $('pg').value = S.page; return; }
  showPage(v);
};
$('pg').addEventListener('change', goPage);
$('pg').addEventListener('keydown', e => { if (e.key === 'Enter') { goPage(); $('pg').blur(); } });
$('pg').addEventListener('focus', () => $('pg').select());
$('prev').onclick = () => S.pdf && showPage(S.page - 1);
$('next').onclick = () => S.pdf && showPage(S.page + 1);
$('zi').onclick = () => { if (S.pdf) { S.zoom *= 1.2; showPage(S.page); } };
$('zo').onclick = () => { if (S.pdf) { S.zoom /= 1.2; showPage(S.page); } };
$('fit').onclick = async () => { if (!S.pdf) return; const p = await S.pdf.getPage(S.page); S.zoom = ($('view').clientWidth - 24) / p.getViewport({ scale: 1 }).width; showPage(S.page); };

let drag = null;
const ov = $('ov');
ov.addEventListener('pointerdown', e => {
  if (!S.pdf || e.button !== 0) return; ov.setPointerCapture(e.pointerId);
  const b = ov.getBoundingClientRect(); drag = { x: e.clientX - b.left, y: e.clientY - b.top, b, ctrl: e.ctrlKey || e.metaKey };
  drag.el = document.createElement('div'); drag.el.className = 'sel'; ov.appendChild(drag.el);
});
ov.addEventListener('pointermove', e => {
  if (!drag) return; const x = e.clientX - drag.b.left, y = e.clientY - drag.b.top; drag.x2 = x; drag.y2 = y;
  Object.assign(drag.el.style, { left: Math.min(drag.x, x) + 'px', top: Math.min(drag.y, y) + 'px', width: Math.abs(x - drag.x) + 'px', height: Math.abs(y - drag.y) + 'px' });
});
ov.addEventListener('pointerup', async e => {
  if (!drag) return; const d = drag; drag = null; d.el.remove(); if (d.x2 == null) return;
  const z = S.zoom, r = [Math.min(d.x, d.x2) / z, Math.min(d.y, d.y2) / z, Math.max(d.x, d.x2) / z, Math.max(d.y, d.y2) / z];
  if (r[2] - r[0] < 3 || r[3] - r[1] < 3) return;
  if(S.batch.active){ await runBatchSnips(r); return; }
  await snip(r, e.ctrlKey || e.metaKey || d.ctrl);
});

/* ---------- Formatage, suppression de snip, image de zone, menu contextuel ---------- */
function resetFmt(rg, s) {
  if (s.f !== 0) rg.format.fill.clear();
  if (s.k) { rg.format.font.color = '#000000'; rg.format.font.bold = false; rg.format.horizontalAlignment = 'General'; }
}
async function deleteSnip(s, clear) {
  S.busy = true;
  try {
    await Excel.run(async c => {
      if (s.n) {
        const it = c.workbook.names.getItemOrNullObject(s.n); await c.sync();
        if (!it.isNullObject) {
          const rg = it.getRangeOrNullObject(); await c.sync();
          if (!rg.isNullObject) { if (clear) rg.clear('Contents'); resetFmt(rg, s); }
          it.delete();
        }
      }
      await c.sync();
    });
    S.snips = S.snips.filter(x => s.n ? x.n !== s.n : x !== s); S.focusN = null;
    await saveSnips(); drawMarks();
    say(clear ? 'Snip supprimé et cellule vidée.' : 'Snip retiré (valeur conservée).');
  } catch (e) { say('Erreur : ' + e.message); }
  finally { S.busy = false; schedRefresh(); }
}
async function cropImage(d, page, r, jpeg) {
  d.pdfObj = d.pdfObj || await pdfjsLib.getDocument({ data: d.data.slice() }).promise;
  const pg = await d.pdfObj.getPage(page), sc = 3, vp = pg.getViewport({ scale: sc });
  const full = document.createElement('canvas'); full.width = vp.width; full.height = vp.height;
  const fx = full.getContext('2d'); fx.fillStyle = '#fff'; fx.fillRect(0, 0, full.width, full.height);
  await pg.render({ canvasContext: fx, viewport: vp }).promise;
  const x = Math.max(0, r[0] * sc), y = Math.max(0, r[1] * sc), w = Math.min(full.width - x, (r[2] - r[0]) * sc), h = Math.min(full.height - y, (r[3] - r[1]) * sc);
  const out = document.createElement('canvas'); out.width = Math.max(1, Math.round(w)); out.height = Math.max(1, Math.round(h));
  out.getContext('2d').drawImage(full, x, y, w, h, 0, 0, out.width, out.height);
  return out.toDataURL(jpeg ? 'image/jpeg' : 'image/png', 0.88).split(',')[1];
}
async function insertImage(s, jpeg) {
  try {
    const d = S.docs.find(x => x.id === s.d); say('Création de l\u2019image\u2026');
    const b = await cropImage(d, s.p, s.r, jpeg);
    let fitted = false;
    await Excel.run(async c => {
      const ws = c.workbook.worksheets.getActiveWorksheet(), sel = c.workbook.getSelectedRange();
      sel.load('left,top,width,height,rowCount,columnCount'); await c.sync();
      const w = s.r[2] - s.r[0], h = s.r[3] - s.r[1];
      let k = Math.min(1, 360 / w);
      if (sel.rowCount * sel.columnCount > 1) { k = Math.min(sel.width / w, sel.height / h); fitted = true; }   // plage sélectionnée : image ajustée à la plage
      const sh = ws.shapes.addImage(b); sh.name = 'XLSnip_img_p' + s.p + '_' + newId(); sh.lockAspectRatio = false;
      sh.left = sel.left; sh.top = sel.top; sh.width = w * k; sh.height = h * k; await c.sync();
    });
    say(fitted ? 'Image insérée et ajustée à la plage sélectionnée.' : 'Image insérée à la cellule sélectionnée.');
  } catch (e) { say('Erreur (image) : ' + e.message); }
}
const cm = document.createElement('div'); cm.id = 'cm'; document.body.appendChild(cm);
const hideCm = () => { cm.style.display = 'none'; };
document.addEventListener('pointerdown', e => { if (!cm.contains(e.target)) hideCm(); });
document.addEventListener('keydown', e => { if (e.key === 'Escape') hideCm(); });
$('view').addEventListener('scroll', hideCm);
ov.addEventListener('contextmenu', e => {
  e.preventDefault(); hideCm(); if (!S.pdf) return;
  const b = ov.getBoundingClientRect(), x = (e.clientX - b.left) / S.zoom, y = (e.clientY - b.top) / S.zoom, area = s => (s.r[2] - s.r[0]) * (s.r[3] - s.r[1]);
  const hits = S.snips.filter(s => s.d === S.cur.id && s.p === S.page && x >= s.r[0] && x <= s.r[2] && y >= s.r[1] && y <= s.r[3]).sort((a, b) => area(a) - area(b));
  if (!hits.length) return;
  const s = hits[0], nz = S.snips.filter(x => s.n && x.n === s.n).length, zt = nz > 1 ? ` (${nz} zones)` : ''; cm.innerHTML = '';
  [[`Insérer l\u2019image de la zone dans Excel`, () => insertImage(s)],
   [`Supprimer le snip${zt} et vider la cellule ${s.a}`, () => deleteSnip(s, true)],
   [`Retirer le snip${zt} (garder la valeur)`, () => deleteSnip(s, false)]].forEach(([t, fn]) => {
    const bt = document.createElement('button'); bt.textContent = t; bt.onclick = () => { hideCm(); fn(); }; cm.appendChild(bt);
  });
  cm.style.display = 'block';
  cm.style.left = Math.max(0, Math.min(e.clientX, innerWidth - cm.offsetWidth - 4)) + 'px';
  cm.style.top = Math.max(0, Math.min(e.clientY, innerHeight - cm.offsetHeight - 4)) + 'px';
});

/* ---------- Boîte de dialogue intégrée + compression ---------- */
function ask(title, body, okLabel, init) {
  return new Promise(res => {
    $('mt').textContent = title; $('mb').innerHTML = body; $('mok').textContent = okLabel || 'Valider'; $('modal').hidden = false; if (init) init();
    const done = v => { $('modal').hidden = true; res(v); };
    $('mok').onclick = () => done(true); $('mno').onclick = () => done(false);
  });
}
document.addEventListener('keydown', e => { if (e.key === 'Escape' && !$('modal').hidden) $('mno').click(); });

async function rasterize(d, q, sc, tag, gray) {
  const pdf = d.pdfObj || await pdfjsLib.getDocument({ data: d.data.slice() }).promise; d.pdfObj = pdf;
  let out = null;
  for (let n = 1; n <= pdf.numPages; n++) {
    say(`${tag} page ${n}/${pdf.numPages}…`);
    const pg = await pdf.getPage(n), v1 = pg.getViewport({ scale: 1 }), v = pg.getViewport({ scale: sc });
    const cv = document.createElement('canvas'); cv.width = Math.ceil(v.width); cv.height = Math.ceil(v.height);
    const cx = cv.getContext('2d'); cx.fillStyle = '#fff'; cx.fillRect(0, 0, cv.width, cv.height);
    await pg.render({ canvasContext: cx, viewport: v }).promise;
    if (gray) {
      const im = cx.getImageData(0, 0, cv.width, cv.height), px = im.data;
      for (let i = 0; i < px.length; i += 4) { const y = (px[i] * 77 + px[i + 1] * 150 + px[i + 2] * 29) >> 8; px[i] = px[i + 1] = px[i + 2] = y; }
      cx.putImageData(im, 0, 0);
    }
    const jpg = cv.toDataURL('image/jpeg', q); cv.width = cv.height = 0;
    const o = v1.width > v1.height ? 'l' : 'p';
    if (!out) out = new window.jspdf.jsPDF({ unit: 'pt', format: [v1.width, v1.height], orientation: o, compress: true });
    else out.addPage([v1.width, v1.height], o);
    out.addImage(jpg, 'JPEG', 0, 0, v1.width, v1.height, undefined, 'FAST');
  }
  return new Uint8Array(out.output('arraybuffer'));
}
const LVL = `<label>Niveau&nbsp;: <select id="cq"><option value="0.7|1.5">Standard</option><option value="0.5|1.2">Forte</option><option value="0.4|1">Maximale</option></select></label>`;
const lvl = () => $('cq').value.split('|').map(Number);
const setBusy = v => ['imp', 'del', 'cmp', 'gray', 'dup'].forEach(id => { $(id).disabled = v; });
const pct = (b, a) => b ? Math.round(100 * (1 - a / b)) : 0;

async function reencode(list, q, sc, gray, verb) {
  let before = 0, after = 0, changed = 0;
  for (const d of list) {
    before += d.data.length;
    const nd = await rasterize(d, q, sc, `${verb} « ${d.name} » :`, gray);
    if (nd.length < d.data.length * 0.95) { d.data = nd; d.pdfObj = null; d.hash = null; changed++; }   // sinon on garde l'original
    after += d.data.length;
    say(`Enregistrement de « ${d.name} »…`);
    const oldPid = d.pid; await saveDoc(d);
    await Excel.run(async c => { c.workbook.customXmlParts.getItem(oldPid).delete(); await c.sync(); });
  }
  fillDocs(); if (S.cur) await openDoc(S.cur.id, S.page);
  return { before, after, changed, total: list.length };
}

$('cmp').onclick = async () => {
  if (!S.docs.length) { say('Aucun document à compresser.'); return; }
  const tot = S.docs.reduce((n, d) => n + d.data.length, 0);
  const ok = await ask('Compresser les documents ?',
    `<p>${S.docs.length} document(s) intégré(s), ${fmtSize(tot)} au total.</p>` +
    `<p>Les pages sont converties en images JPEG allégées. Les snips et le texte détecté sont conservés, mais les documents perdent leur netteté d'origine et leur texte sélectionnable.</p>` +
    `<p><b>Action irréversible</b> : gardez une copie du classeur avant de valider.</p>` + LVL, 'Compresser');
  if (!ok) { say('Compression annulée.'); return; }
  const [q, sc] = lvl(); setBusy(true);
  try {
    const r = await reencode(S.docs, q, sc, false, 'Compression de');
    say(`Compression terminée : ${fmtSize(r.before)} → ${fmtSize(r.after)} (−${pct(r.before, r.after)} %). Enregistrez le classeur pour appliquer le gain.`);
  } catch (e) { say('Erreur de compression : ' + e.message); }
  finally { setBusy(false); fillDocs(); }
};

$('gray').onclick = async () => {
  if (!S.docs.length) { say('Aucun document à convertir.'); return; }
  const rows = S.docs.map(d => `<label><input type="checkbox" class="gk" value="${d.id}">${esc(d.name)} (${fmtSize(d.data.length)})</label>`).join('');
  const ok = await ask('Passer en niveaux de gris',
    `<p>Cochez les documents à convertir.</p><div class="gtools"><button type="button" id="gsa">Tout sélectionner</button><button type="button" id="gsn">Tout désélectionner</button></div>` +
    `<div class="gl">${rows}</div><p>Les pages sont converties en images en niveaux de gris : <b>action irréversible</b> (gardez une copie du classeur). Un document qui ne devient pas plus léger est laissé en couleur.</p>` + LVL, 'Convertir',
    () => { const all = v => document.querySelectorAll('.gk').forEach(x => { x.checked = v; }); $('gsa').onclick = () => all(true); $('gsn').onclick = () => all(false); });
  if (!ok) { say('Conversion annulée.'); return; }
  const ids = [...document.querySelectorAll('.gk')].filter(x => x.checked).map(x => x.value), list = S.docs.filter(d => ids.includes(d.id));
  if (!list.length) { say('Aucun document sélectionné.'); return; }
  const [q, sc] = lvl(); setBusy(true);
  try {
    const r = await reencode(list, q, sc, true, 'Niveaux de gris :');
    say(`${r.changed}/${r.total} document(s) converti(s) et allégé(s) : ${fmtSize(r.before)} → ${fmtSize(r.after)} (−${pct(r.before, r.after)} %).` + (r.changed < r.total ? ' Les autres seraient devenus plus lourds : laissés en couleur.' : '') + ' Enregistrez le classeur.');
  } catch (e) { say('Erreur de conversion : ' + e.message); }
  finally { setBusy(false); fillDocs(); }
};

async function hashDoc(d) {
  if (window.crypto && crypto.subtle) { const h = await crypto.subtle.digest('SHA-256', d.data); return [...new Uint8Array(h)].map(x => x.toString(16).padStart(2, '0')).join(''); }
  let a = 0x811c9dc5, b = 5381;
  for (let i = 0; i < d.data.length; i++) { a = Math.imul(a ^ d.data[i], 16777619); b = (Math.imul(b, 33) + d.data[i]) | 0; }
  return d.data.length + ':' + (a >>> 0) + ':' + (b >>> 0);
}
$('dup').onclick = async () => {
  if (S.docs.length < 2) { say('Moins de deux documents : pas de doublon possible.'); return; }
  setBusy(true);
  try {
    say('Recherche des doublons…');
    const gp = new Map();
    for (const d of S.docs) { d.hash = d.hash || await hashDoc(d); if (!gp.has(d.hash)) gp.set(d.hash, []); gp.get(d.hash).push(d); }
    const groups = [...gp.values()].filter(g => g.length > 1);
    if (!groups.length) { say('Aucun doublon trouvé.'); return; }
    const extra = groups.flatMap(g => g.slice(1)), saved = extra.reduce((n, d) => n + d.data.length, 0);
    const li = groups.map(g => `<li>« ${esc(g[0].name)} » : ${g.length} exemplaires` + (g.some(d => d.name !== g[0].name) ? ` (${g.map(d => '« ' + esc(d.name) + ' »').join(', ')})` : '') + '</li>').join('');
    const ok = await ask('Supprimer les doublons ?',
      `<p>${extra.length} copie(s) identique(s) trouvée(s), soit ${fmtSize(saved)} à gagner :</p><ul>${li}</ul>` +
      `<p>Le premier exemplaire de chaque groupe est conservé ; les snips des copies lui sont rattachés.</p><p><b>Action irréversible.</b></p>`, 'Supprimer');
    if (!ok) { say('Suppression annulée.'); return; }
    for (const g of groups) for (const o of g.slice(1)) { S.snips.forEach(s => { if (s.d === o.id) s.d = g[0].id; }); if (S.cur === o) S.cur = g[0]; }
    await Excel.run(async c => { extra.forEach(o => c.workbook.customXmlParts.getItem(o.pid).delete()); await c.sync(); });
    S.docs = S.docs.filter(d => !extra.includes(d));
    await saveSnips(); fillDocs();
    if (S.cur) await openDoc(S.cur.id, S.page);
    say(`${extra.length} doublon(s) supprimé(s) : ${fmtSize(saved)} libérés. Enregistrez le classeur.`);
  } catch (e) { say('Erreur : ' + e.message); }
  finally { setBusy(false); fillDocs(); }
};

/* ---------- Séparateur redimensionnable : agrandir / réduire la visionneuse ---------- */
(function () {
  const tools = $('tools'), grip = $('grip'), MIN = 0;      // 0 = zone d'outils entièrement repliable
  let h = null;                                   // null = hauteur naturelle des outils
  const store = (k, v) => { try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, v); } catch (e) {} };
  const load = k => { try { return localStorage.getItem(k); } catch (e) { return null; } };
  const natural = () => { const s = tools.style.height; tools.style.height = ''; const n = tools.offsetHeight; tools.style.height = s; return n; };
  function apply() {
    tools.style.height = '';
    const n = tools.offsetHeight, lim = Math.max(MIN, Math.min(n, innerHeight - 260));   // laisse toujours de la place à la visionneuse
    const t = Math.min(h == null ? n : h, lim);
    if (t < n) tools.style.height = Math.max(MIN, t) + 'px';
  }
  const saved = parseInt(load('xlsnip.toolsH'), 10); if (!isNaN(saved)) h = saved;
  apply();
  let st = null;
  grip.addEventListener('pointerdown', e => {
    if (e.button !== 0) return;
    grip.setPointerCapture(e.pointerId); st = { y: e.clientY, h: tools.offsetHeight, n: natural() };
    grip.classList.add('drag'); document.body.classList.add('dragging'); e.preventDefault();
  });
  grip.addEventListener('pointermove', e => {
    if (!st) return;
    const v = Math.max(MIN, Math.min(st.n, st.h + e.clientY - st.y));
    h = v >= st.n - 2 ? null : Math.round(v); apply();
  });
  const end = () => { if (!st) return; st = null; grip.classList.remove('drag'); document.body.classList.remove('dragging'); store('xlsnip.toolsH', h); };
  grip.addEventListener('pointerup', end); grip.addEventListener('pointercancel', end);
  grip.addEventListener('dblclick', () => { h = (h == null) ? MIN : null; apply(); store('xlsnip.toolsH', h); });
  grip.addEventListener('keydown', e => {
    const cur = tools.offsetHeight, n = natural(); let v = null;
    if (e.key === 'ArrowUp') v = cur - 30; else if (e.key === 'ArrowDown') v = cur + 30;
    else if (e.key === 'Home') v = MIN; else if (e.key === 'End') v = n; else return;
    e.preventDefault(); v = Math.max(MIN, Math.min(n, v)); h = v >= n - 2 ? null : v; apply(); store('xlsnip.toolsH', h);
  });
  addEventListener('resize', apply);
  // Bloc « Alléger » replié par défaut
  const lg = $('lg'), alg = $('alg');
  const setAlg = open => { alg.hidden = !open; lg.setAttribute('aria-expanded', open); apply(); store('xlsnip.alg', open ? '1' : null); };
  lg.onclick = () => setAlg(alg.hidden);
  if (load('xlsnip.alg') === '1') setAlg(true);
})();


/* ---------- Initialisation du complément ---------- */
let APP_READY = false;
async function initXLSnip() {
  try {
    say('Chargement des documents…');
    await loadAll();
    fillDocs();
    try {
      await Excel.run(async c => {
        const ws = c.workbook.worksheets.getActiveWorksheet();
        ws.onSelectionChanged.add(onSel);
        await c.sync();
      });
    } catch (e) { /* le suivi de source reste disponible dès que l'hôte expose l'événement */ }
    renderOrgTree();
    APP_READY = true;
    say(S.docs.length ? `${S.docs.length} document(s) chargé(s). Prêt.` : 'Prêt. Aucun document importé.');
  } catch (e) {
    console.error('XLSnip initialisation', e);
    say(`Impossible de charger les documents : ${e.message || e}`);
  }
}
if (typeof Office !== 'undefined' && Office.onReady) {
  Office.onReady(info => {
    if (info && info.host && info.host !== Office.HostType.Excel) { say('XLSnip doit être utilisé dans Excel.'); return; }
    initXLSnip();
  });
} else {
  window.addEventListener('load', initXLSnip);
}

/* ---------- Organisation des documents ---------- */
async function saveFolders() {
  await Excel.run(async c => {
    const ps=c.workbook.customXmlParts.getByNamespace(NS_ORG); ps.load('items'); await c.sync();
    ps.items.forEach(p=>p.delete());
    c.workbook.customXmlParts.add(`<o xmlns="${NS_ORG}">${jb64(S.folders)}</o>`); await c.sync();
  });
}
function folderById(id){ return S.folders.find(f=>f.id===id) || S.folders[0]; }
function folderChildren(id){ return S.folders.filter(f=>f.parent===id); }
function clearOrgDropTargets(){
  document.querySelectorAll('.orgDropTarget').forEach(el=>el.classList.remove('orgDropTarget'));
}
function dragDocIdFromEvent(e){
  const id=e.dataTransfer && e.dataTransfer.getData('text/x-xlsnip-doc');
  return id || (e.dataTransfer && e.dataTransfer.getData('text/plain')) || '';
}
async function moveDocumentToFolder(docId, folderId){
  const d=S.docs.find(x=>x.id===docId);
  const f=folderById(folderId);
  if(!d||!f)return;
  const current=d.folder||'root';
  if(current===f.id){
    S.orgDocSel=d.id;
    S.orgSel=f.id;
    renderOrgTree();
    return say(`« ${d.name} » est déjà dans « ${f.name} ».`);
  }
  d.folder=f.id;
  S.orgDocSel=d.id;
  S.orgSel=f.id;
  await rewriteDoc(d);
  fillDocs();
  renderOrgTree();
  say(`« ${d.name} » déplacé dans « ${f.name} ».`);
}
let ORG_DRAG=null;
function clearOrgDropTargets(){
  document.querySelectorAll('.orgDropTarget').forEach(el=>el.classList.remove('orgDropTarget'));
}
function dragDocIdFromEvent(e){
  const id=e.dataTransfer && e.dataTransfer.getData('text/x-xlsnip-doc');
  return id || (e.dataTransfer && e.dataTransfer.getData('text/plain')) || '';
}
function endPointerOrgDrag(){
  if(!ORG_DRAG)return;
  const ghost=ORG_DRAG.ghost;
  if(ghost&&ghost.parentNode)ghost.parentNode.removeChild(ghost);
  if(ORG_DRAG.el)ORG_DRAG.el.classList.remove('orgDragging');
  clearOrgDropTargets();
  const dropFolder=ORG_DRAG.dropFolder;
  const docId=ORG_DRAG.docId;
  const wasDragging=ORG_DRAG.active;
  ORG_DRAG=null;
  if(wasDragging&&docId&&dropFolder){
    moveDocumentToFolder(docId,dropFolder).catch(err=>say(`Erreur lors du déplacement : ${err.message||err}`));
    return true;
  }
  return false;
}
function onPointerOrgMove(e){
  if(!ORG_DRAG)return;
  const dx=e.clientX-ORG_DRAG.x, dy=e.clientY-ORG_DRAG.y;
  if(!ORG_DRAG.active && Math.hypot(dx,dy)<6)return;
  if(!ORG_DRAG.active){
    ORG_DRAG.active=true;
    ORG_DRAG.el.classList.add('orgDragging');
    const ghost=document.createElement('div');
    ghost.className='orgDragGhost';
    ghost.textContent=ORG_DRAG.name;
    document.body.appendChild(ghost);
    ORG_DRAG.ghost=ghost;
  }
  e.preventDefault();
  if(ORG_DRAG.ghost){
    ORG_DRAG.ghost.style.left=(e.clientX+12)+'px';
    ORG_DRAG.ghost.style.top=(e.clientY+12)+'px';
  }
  clearOrgDropTargets();
  const under=document.elementFromPoint(e.clientX,e.clientY);
  const folder=under&&under.closest('.orgFolder');
  if(folder&&$('orgTree')&&$('orgTree').contains(folder)){
    ORG_DRAG.dropFolder=folder.dataset.folderId||null;
    folder.classList.add('orgDropTarget');
  }else{
    ORG_DRAG.dropFolder=null;
  }
}
function startPointerOrgDrag(e,d,id){
  if(e.button!==undefined&&e.button!==0)return;
  ORG_DRAG={
    docId:d.id,name:d.name,el:e.currentTarget,
    x:e.clientX,y:e.clientY,active:false,ghost:null,dropFolder:null
  };
  S.orgDocSel=d.id;
  S.orgSel=id;
}
function renderOrgTree(){
  const root=$('orgTree'); if(!root)return; root.innerHTML='';
  if(!S.folders.length) S.folders=[{id:'root',name:'Documents',parent:null}];

  const walk=(id,depth)=>{
    const f=folderById(id); if(!f)return;
    const docs=S.docs.filter(d=>(d.folder||'root')===id);
    const children=folderChildren(id);

    const el=document.createElement('div');
    el.className='orgItem orgFolder'+(S.orgSel===id?' sel':'');
    el.style.paddingLeft=(7+depth*18)+'px';
    el.dataset.folderId=f.id;
    el.innerHTML=`<span class="ico">📁</span><span class="orgName">${esc(f.name)}</span><span class="orgCount">${docs.length}</span>`;
    el.onclick=()=>{S.orgSel=id;renderOrgTree();};

    // Pointer-based drop target: works reliably inside an Excel task pane/webview.
    el.ondragover=e=>{e.preventDefault();if(e.dataTransfer)e.dataTransfer.dropEffect='move';};
    el.ondragenter=()=>el.classList.add('orgDropTarget');
    el.ondragleave=e=>{if(!el.contains(e.relatedTarget))el.classList.remove('orgDropTarget');};
    el.ondrop=async e=>{
      e.preventDefault();e.stopPropagation();
      const docId=dragDocIdFromEvent(e);
      clearOrgDropTargets();
      if(docId)try{await moveDocumentToFolder(docId,f.id)}catch(err){say(`Erreur lors du déplacement : ${err.message||err}`)};
    };
    root.appendChild(el);

    const docsWrap=document.createElement('div');
    docsWrap.className='orgDocs';

    docs.forEach(d=>{
      const de=document.createElement('div');
      de.className='orgItem orgDoc'+(S.orgDocSel===d.id?' selDoc':'');
      de.style.paddingLeft=(31+depth*18)+'px';
      de.dataset.docId=d.id;
      de.draggable=true;
      de.innerHTML=`<span class="ico">📄</span><span class="orgName">${esc(d.name)}</span><span class="orgCount">${fmtSize(d.data.length)}</span>`;

      de.onclick=()=>{
        if(ORG_DRAG&&ORG_DRAG.active){return;}
        S.orgDocSel=d.id;
        S.orgSel=id;
        renderOrgTree();
      };
      de.ondblclick=async()=>{
        if(ORG_DRAG)return;
        S.orgDocSel=d.id;
        await openDoc(d.id);
        renderOrgTree();
      };

      // Native HTML5 DnD retained as a fallback.
      de.ondragstart=e=>{
        S.orgDocSel=d.id;S.orgSel=id;
        if(e.dataTransfer){
          e.dataTransfer.effectAllowed='move';
          e.dataTransfer.setData('text/x-xlsnip-doc',d.id);
          e.dataTransfer.setData('text/plain',d.id);
        }
        de.classList.add('orgDragging');
      };
      de.ondragend=()=>{
        de.classList.remove('orgDragging');
        clearOrgDropTargets();
      };

      // Pointer DnD is the primary mechanism in Excel's task pane.
      de.onpointerdown=e=>startPointerOrgDrag(e,d,id);
      docsWrap.appendChild(de);
    });

    root.appendChild(docsWrap);
    children.forEach(c=>walk(c.id,depth+1));
  };

  walk('root',0);
  if(!S.docs.length){
    const e=document.createElement('div');
    e.className='muted';
    e.style.padding='12px 8px';
    e.textContent='Aucun document importé. Utilisez « Importer ici » pour ajouter un PDF ou une image.';
    root.appendChild(e);
  }
  fillDocs();   // garde le menu des dossiers et la liste des documents synchronisés
}

async function askFolderName(title, defaultName=''){
  return new Promise(resolve=>{
    const body=`<label for="orgFolderName">${esc(title)}</label><input id="orgFolderName" class="fullInput" type="text" value="${esc(defaultName)}" placeholder="Nom du dossier" autocomplete="off">`;
    $('mt').textContent=title;
    $('mb').innerHTML=body;
    $('mok').textContent='Valider';
    $('modal').hidden=false;
    const input=$('orgFolderName');
    const close=ok=>{
      $('modal').hidden=true;
      $('mok').onclick=null;
      $('mno').onclick=null;
      resolve(ok ? input.value.trim() : null);
    };
    $('mok').onclick=()=>close(true);
    $('mno').onclick=()=>close(false);
    input.focus();
    input.select();
    input.onkeydown=e=>{
      if(e.key==='Enter'){e.preventDefault();close(true);}
      if(e.key==='Escape'){e.preventDefault();close(false);}
    };
  });
}
async function renameFolder(){
  const f=folderById(S.orgSel);
  if(!f||f.id==='root')return say('Le dossier racine ne peut pas être renommé.');
  const n=await askFolderName('Renommer le dossier',f.name);
  if(!n)return;
  f.name=n;
  await saveFolders();
  renderOrgTree();
  say(`Dossier renommé en « ${n} ».`);
}
async function newFolder(parent){
  const p=folderById(parent);
  const name=await askFolderName(parent==='root'?'Nouveau dossier':'Nouveau sous-dossier');
  if(!name)return;
  const id='f'+newId();
  S.folders.push({id,name,parent:p ? p.id : 'root'});
  S.orgSel=id;
  S.orgDocSel=null;
  await saveFolders();
  renderOrgTree();
  say(`Dossier « ${name} » créé.`);
}
async function deleteFolder(){ const f=folderById(S.orgSel); if(!f||f.id==='root')return; if(folderChildren(f.id).length||S.docs.some(d=>(d.folder||'root')===f.id))return say('Le dossier doit être vide avant suppression.'); S.folders=S.folders.filter(x=>x.id!==f.id); S.orgSel=f.parent||'root'; await saveFolders(); renderOrgTree(); }
async function moveSelectedToFolder(){
  const d=S.docs.find(x=>x.id===S.orgDocSel);
  if(!d)return say('Sélectionnez d’abord un document dans l’organisation.');
  const f=folderById(S.orgSel);
  if(!f)return say('Sélectionnez un dossier de destination.');
  await moveDocumentToFolder(d.id,f.id);
}
async function rewriteDoc(d){ await Excel.run(async c=>{ const old=d.pid; const xml=`<d xmlns="${NS_DOC}" id="${d.id}" name="${esc(d.name)}" folder="${esc(d.folder||'root')}" wf="json"><f>${b64(d.data)}</f><w>${escT(JSON.stringify(d.words))}</w></d>`; const p=c.workbook.customXmlParts.add(xml); p.load('id'); await c.sync(); c.workbook.customXmlParts.getItem(old).delete(); await c.sync(); d.pid=p.id; }); }
$('orgBtn').onclick=async()=>{
  $('orgModal').hidden=false;
  renderOrgTree();
  if(!APP_READY){ await initXLSnip(); renderOrgTree(); }
};
$('orgClose').onclick=()=>{$('orgModal').hidden=true};
$('orgNewFolder').onclick=async()=>{try{await newFolder('root')}catch(e){say(`Erreur : ${e.message||e}`)}};
$('orgNewSub').onclick=async()=>{try{await newFolder(S.orgSel||'root')}catch(e){say(`Erreur : ${e.message||e}`)}};
$('orgRename').onclick=async()=>{try{await renameFolder()}catch(e){say(`Erreur : ${e.message||e}`)}};
$('orgDelete').onclick=async()=>{try{await deleteFolder()}catch(e){say(`Erreur : ${e.message||e}`)}};
$('orgImport').onclick=()=>{ if(!APP_READY){say('Le chargement n’est pas terminé.'); return;} $('orgFile').click(); };
$('orgMove').onclick=async()=>{try{await moveSelectedToFolder()}catch(e){say(`Erreur : ${e.message||e}`)}};
document.addEventListener('pointermove',onPointerOrgMove,{passive:false});
document.addEventListener('pointerup',()=>{endPointerOrgDrag()},{passive:true});
document.addEventListener('pointercancel',()=>{endPointerOrgDrag()},{passive:true});
$('orgFile').onchange=async e=>{
  const target=S.orgSel||'root'; const files=[...e.target.files]; e.target.value='';
  for(const f of files){try{await importFile(f,{open:false}); const d=S.docs[S.docs.length-1]; if(d){d.folder=target;S.orgDocSel=d.id;await rewriteDoc(d);}}catch(err){say(`Échec de l'import de « ${f.name} » : ${err.message}`)}}
  renderOrgTree(); fillDocs();
};

/* ---------- Snips automatiques en série ---------- */
function batchFolderPath(id){
  const parts=[];let f=folderById(id),guard=0;
  while(f&&guard++<100){parts.unshift(f.name);f=S.folders.find(x=>x.id===f.parent);}
  return parts.join(' / ');
}
function batchDescendantFolders(id){return [id,...folderChildren(id).flatMap(f=>batchDescendantFolders(f.id))];}
function batchDocsForFolder(id){
  const ids=new Set(batchDescendantFolders(id));
  return S.docs.filter(d=>ids.has(d.folder||'root')).sort((a,b)=>String(a.name).localeCompare(String(b.name),'fr',{numeric:true,sensitivity:'base'}));
}
function renderBatchFolders(){
  const el=$('batchFolderTree');if(!el)return;el.innerHTML='';
  const walk=(id,depth)=>{
    const f=folderById(id);if(!f)return;
    const docs=batchDocsForFolder(id);
    const row=document.createElement('button');row.type='button';row.className='batchFolderRow'+(S.batch.folderId===id?' selected':'');
    row.style.paddingLeft=(10+depth*18)+'px';row.innerHTML=`<span>📁</span><span class="batchFolderName">${esc(f.name)}</span><small>${docs.length} document${docs.length===1?'':'s'} dans ce dossier et ses sous-dossiers</small>`;
    row.onclick=()=>{S.batch.folderId=id;renderBatchFolders();const info=$('batchFolderInfo');if(info)info.textContent=`Dossier choisi : ${batchFolderPath(id)} · ${docs.length} document(s), triés alphabétiquement et numériquement.`;};
    el.appendChild(row);folderChildren(id).slice().sort((a,b)=>a.name.localeCompare(b.name,'fr',{numeric:true,sensitivity:'base'})).forEach(c=>walk(c.id,depth+1));
  };
  walk('root',0);
}
async function openBatchSnips(){
  if(!S.docs.length){say('Importez d’abord les documents à traiter.');return;}
  S.batch={active:false,folderId:S.orgSel||'root',orientation:'vertical',docs:[],startRow:0,startCol:0,sheetName:'',page:1};
  $('batchVertical').checked=true;$('batchHorizontal').checked=false;
  renderBatchFolders();$('batchFolderInfo').textContent='Choisissez le dossier qui contient les documents au même format.';
  $('batchModal').hidden=false;
}
async function startBatchSnips(){
  const folderId=S.batch.folderId||'root',docs=batchDocsForFolder(folderId);
  if(!docs.length){say('Ce dossier ne contient aucun document à traiter.');return;}
  if(['table','img'].includes(S.mode)){say('Pour les snips en série, choisissez un mode qui produit une valeur dans une seule cellule (Texte, Nombre, Somme ou Pointage).');return;}
  let start;
  await Excel.run(async c=>{
    const rg=c.workbook.getSelectedRange();rg.load('rowIndex,columnIndex,worksheet/name,address');await c.sync();
    start={rowIndex:rg.rowIndex,columnIndex:rg.columnIndex,sheetName:rg.worksheet.name,address:rg.address};
  });
  S.batch={active:true,folderId,orientation:$('batchHorizontal').checked?'horizontal':'vertical',docs,startRow:start.rowIndex,startCol:start.columnIndex,sheetName:start.sheetName,page:1};
  $('batchModal').hidden=true;
  await openDoc(docs[0].id,1);
  say(`Modèle prêt : ${docs[0].name}. Tracez le premier snip sur le document. Les ${docs.length} documents seront traités dans l’ordre alphabétique et numérique.`);
}
async function runBatchSnips(rect){
  if(!S.batch.active)return;
  const batch=S.batch;batch.active=false;
  const docs=batch.docs.slice(),page=S.page,mode=S.mode;
  if(!docs.length)return;
  let completed=0,failed=[];
  try{
    for(let i=0;i<docs.length;i++){
      const d=docs[i];
      try{
        say(`Snips en série : ${i+1}/${docs.length} — ${d.name}`);
        await openDoc(d.id,Math.min(page,d.pdfObj?.numPages||page));
        const target={sheetName:batch.sheetName,rowIndex:batch.startRow+(batch.orientation==='vertical'?i:0),columnIndex:batch.startCol+(batch.orientation==='horizontal'?i:0)};
        const ok=await snip(rect,false,target);
        if(ok)completed++;else failed.push(`${d.name}: aucun texte exploitable dans la zone`);
      }catch(err){failed.push(`${d.name}: ${err.message||err}`);}
    }
    if(completed){await saveSnips();say(`${completed}/${docs.length} snip(s) créé(s) ${batch.orientation==='vertical'?'verticalement':'horizontalement'} à partir de ${batchFolderPath(batch.folderId)}.${failed.length?' Échecs : '+failed.join(' ; '):''}`);}
    else say('Aucun snip créé. Vérifiez que la zone contient du texte dans le même emplacement de chaque document.');
  }finally{
    S.batch.active=false;
  }
}
$('batchSnipBtn').onclick=openBatchSnips;
$('batchCancel').onclick=()=>{$('batchModal').hidden=true;S.batch.active=false;};
$('batchStart').onclick=async()=>{try{await startBatchSnips();}catch(e){S.batch.active=false;say('Impossible de démarrer les snips en série : '+(e.message||e));}};
