/* XLSnip – complément Excel d'extraction de données depuis des documents (PDF / images) */
const NS_DOC = 'urn:xlsnip:doc', NS_SNIP = 'urn:xlsnip:snips';
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

const S = { docs: [], cur: null, pdf: null, page: 1, zoom: 1.3, mode: 'text', snips: [], focus: null, task: null, ocr: null, busy: false, rf: false, again: false, lastRf: 0 };
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
    const xml = `<d xmlns="${NS_DOC}" id="${d.id}" name="${esc(d.name)}" wf="json"><f>${b64(d.data)}</f><w>${escT(JSON.stringify(d.words))}</w></d>`;
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
    S.docs = xs.map((x, i) => {
      const e = new DOMParser().parseFromString(x.value, 'text/xml').documentElement;
      return { pid: parts.items[i].id, id: e.getAttribute('id'), name: e.getAttribute('name'),
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
async function importFile(f) {
  let buf;
  if (/^image\//.test(f.type)) buf = await imgToPdf(f);
  else if (/\.pdf$/i.test(f.name) || f.type === 'application/pdf') buf = new Uint8Array(await f.arrayBuffer());
  else { say(`« ${f.name} » : format non pris en charge. Exportez-le d'abord en PDF.`); return; }
  const pdf = await pdfjsLib.getDocument({ data: buf.slice() }).promise;
  const d = { id: 'd' + Date.now() + Math.floor(Math.random() * 1000), name: f.name, data: buf, words: await analyse(pdf, f.name) };
  say('Enregistrement dans le classeur…');
  await saveDoc(d); S.docs.push(d); d.pdfObj = pdf;
  fillDocs(); await openDoc(d.id);
  say(`« ${f.name} » importé (${pdf.numPages} page(s)).`);
}

/* ---------- Visionneuse ---------- */
function fillDocs() {
  $('docs').innerHTML = S.docs.map(d => `<option value="${d.id}">${esc(d.name)} (${fmtSize(d.data.length)})</option>`).join('');
  if (S.cur) $('docs').value = S.cur.id;
}
async function openDoc(id, page) {
  const d = S.docs.find(x => x.id === id); if (!d) return;
  S.cur = d; d.pdfObj = d.pdfObj || await pdfjsLib.getDocument({ data: d.data.slice() }).promise;
  S.pdf = d.pdfObj; $('docs').value = id; $('empty').hidden = true; $('wrap').hidden = false;
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
  const ov = $('ov'); ov.querySelectorAll('.mark').forEach(e => e.remove());
  const add = (r, cls) => { const e = document.createElement('div'); e.className = 'mark ' + cls; const z = S.zoom;
    Object.assign(e.style, { left: r[0] * z + 'px', top: r[1] * z + 'px', width: (r[2] - r[0]) * z + 'px', height: (r[3] - r[1]) * z + 'px' }); ov.appendChild(e); };
  S.snips.filter(s => s.d === S.cur.id && s.p === S.page).forEach(s => add(s.r, (s.k || '') + (S.focusN && s.n === S.focusN ? ' f' : '')));
}

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

async function snip(r, cont) {
  await refreshSnips();
  const chk = S.mode === 'ok' || S.mode === 'nok';
  let grid, raw = null, num = null;
  if (chk) grid = [[S.mode === 'ok' ? '\u2713' : '\u2717']];
  else {
    const ws = (S.cur.words[S.page - 1] || []).filter(w => { const cx = (w[1] + w[3]) / 2, cy = (w[2] + w[4]) / 2; return cx >= r[0] && cx <= r[2] && cy >= r[1] && cy <= r[3]; });
    if (!ws.length) { say('Aucun texte détecté dans cette zone.'); return; }
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
      const sh = c.workbook.worksheets.getActiveWorksheet(); sh.load('name');
      const act = c.workbook.getActiveCell(); act.load('address,values'); await c.sync();
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
    await saveSnips(); drawMarks(); say(msg);
  } catch (e) { say('Erreur Excel : ' + e.message); }
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
$('docs').onchange = e => openDoc(e.target.value);
$('del').onclick = async () => {
  if (!S.cur || !(await ask('Retirer ce document ?', `<p>« ${esc(S.cur.name)} » et ses snips seront retirés du classeur.</p>`, 'Retirer'))) return;
  const d = S.cur;
  await Excel.run(async c => { c.workbook.customXmlParts.getItem(d.pid).delete(); await c.sync(); });
  S.docs = S.docs.filter(x => x !== d); S.snips = S.snips.filter(s => s.d !== d.id); await saveSnips();
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
async function cropImage(d, page, r) {
  d.pdfObj = d.pdfObj || await pdfjsLib.getDocument({ data: d.data.slice() }).promise;
  const pg = await d.pdfObj.getPage(page), sc = 3, vp = pg.getViewport({ scale: sc });
  const full = document.createElement('canvas'); full.width = vp.width; full.height = vp.height;
  await pg.render({ canvasContext: full.getContext('2d'), viewport: vp }).promise;
  const x = Math.max(0, r[0] * sc), y = Math.max(0, r[1] * sc), w = Math.min(full.width - x, (r[2] - r[0]) * sc), h = Math.min(full.height - y, (r[3] - r[1]) * sc);
  const out = document.createElement('canvas'); out.width = Math.max(1, Math.round(w)); out.height = Math.max(1, Math.round(h));
  out.getContext('2d').drawImage(full, x, y, w, h, 0, 0, out.width, out.height);
  return out.toDataURL('image/png').split(',')[1];
}
async function insertImage(s) {
  try {
    const d = S.docs.find(x => x.id === s.d); say('Création de l\u2019image\u2026');
    const b = await cropImage(d, s.p, s.r);
    await Excel.run(async c => {
      const ws = c.workbook.worksheets.getActiveWorksheet(), cell = c.workbook.getActiveCell(); cell.load('left,top'); await c.sync();
      const w = s.r[2] - s.r[0], h = s.r[3] - s.r[1], k = Math.min(1, 360 / w);
      const sh = ws.shapes.addImage(b); sh.name = 'XLSnip_img_' + newId(); sh.lockAspectRatio = false;
      sh.left = cell.left; sh.top = cell.top; sh.width = w * k; sh.height = h * k; await c.sync();
    });
    say('Image insérée à la cellule active.');
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
    `<p>Cochez les documents à convertir.</p><div class="sel"><button type="button" id="gsa">Tout sélectionner</button><button type="button" id="gsn">Tout désélectionner</button></div>` +
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

Office.onReady(async info => {
  if (info.host !== Office.HostType.Excel) { say('Ouvrez ce volet depuis Excel.'); return; }
  try {
    await loadAll(); await refreshSnips(); fillDocs();
    if (S.docs.length) await openDoc(S.docs[0].id);
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSel);
    try {
      await Excel.run(async c => {
        c.workbook.worksheets.onChanged.add(async () => schedRefresh());
        c.workbook.worksheets.onDeleted.add(async () => schedRefresh());
        await c.sync();
      });
    } catch (e) { /* événements indisponibles : le rafraîchissement se fait à la sélection */ }
    say(S.docs.length ? 'Classeur chargé.' : 'Prêt.');
  } catch (e) { say('Erreur d\'initialisation : ' + e.message); }
});
