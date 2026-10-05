/* XLSnip – complément Excel d'extraction de données depuis des documents (PDF / images) */
const NS_DOC = 'urn:xlsnip:doc', NS_SNIP = 'urn:xlsnip:snips';
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

const S = { docs: [], cur: null, pdf: null, page: 1, zoom: 1.3, mode: 'text', snips: [], focus: null, task: null, ocr: null };
const $ = id => document.getElementById(id);
const say = t => { $('st').textContent = t; };
const esc = s => s.replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const b64 = u => { let s = ''; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); };
const unb64 = s => { const b = atob(s), u = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); return u; };
const jb64 = o => btoa(unescape(encodeURIComponent(JSON.stringify(o))));
const bj64 = s => JSON.parse(decodeURIComponent(escape(atob(s))));
const r1 = x => Math.round(x * 10) / 10;

/* ---------- Persistance dans le classeur (parties XML personnalisées) ---------- */
async function saveDoc(d) {
  await Excel.run(async c => {
    const xml = `<d xmlns="${NS_DOC}" id="${d.id}" name="${esc(d.name)}"><f>${b64(d.data)}</f><w>${jb64(d.words)}</w></d>`;
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
        data: unb64(e.getElementsByTagName('f')[0].textContent), words: bj64(e.getElementsByTagName('w')[0].textContent) };
    });
    S.snips = ss.length ? bj64(new DOMParser().parseFromString(ss[0].value, 'text/xml').documentElement.textContent) : [];
  });
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
  $('docs').innerHTML = S.docs.map(d => `<option value="${d.id}">${esc(d.name)}</option>`).join('');
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
  if (S.task) { try { S.task.cancel(); } catch (e) {} }
  const pg = await S.pdf.getPage(n), vp = pg.getViewport({ scale: S.zoom }), cv = $('cv');
  cv.width = vp.width; cv.height = vp.height;
  S.task = pg.render({ canvasContext: cv.getContext('2d'), viewport: vp });
  try { await S.task.promise; } catch (e) { return; }
  $('pn').textContent = `${n} / ${S.pdf.numPages}`; drawMarks();
}
function drawMarks() {
  const ov = $('ov'); ov.querySelectorAll('.mark').forEach(e => e.remove());
  const add = (r, cls) => { const e = document.createElement('div'); e.className = 'mark ' + cls; const z = S.zoom;
    Object.assign(e.style, { left: r[0] * z + 'px', top: r[1] * z + 'px', width: (r[2] - r[0]) * z + 'px', height: (r[3] - r[1]) * z + 'px' }); ov.appendChild(e); };
  S.snips.filter(s => s.d === S.cur.id && s.p === S.page).forEach(s => add(s.r, s.k || ''));
  if (S.focus && S.focus.d === S.cur.id && S.focus.p === S.page) add(S.focus.r, 'f');
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

async function snip(r) {
  const chk = S.mode === 'ok' || S.mode === 'nok';
  let grid;
  if (chk) grid = [[S.mode === 'ok' ? '\u2713' : '\u2717']];
  else {
    const ws = (S.cur.words[S.page - 1] || []).filter(w => { const cx = (w[1] + w[3]) / 2, cy = (w[2] + w[4]) / 2; return cx >= r[0] && cx <= r[2] && cy >= r[1] && cy <= r[3]; });
    if (!ws.length) { say('Aucun texte détecté dans cette zone.'); return; }
    const text = lines(ws).rows.map(l => l.map(w => w[0]).join(' ')).join(' ');
    if (S.mode === 'table') grid = toGrid(ws).map(row => row.map(t => (looksNum(t) && parseNum(t) !== null) ? parseNum(t) : safe(t)));
    else if (S.mode === 'num') { const n = parseNum(text); if (n === null) { say(`« ${text} » n'est pas un nombre.`); return; } grid = [[n]]; }
    else grid = [[safe(text)]];
  }
  try {
    await Excel.run(async c => {
      const sh = c.workbook.worksheets.getActiveWorksheet(); sh.load('name');
      const rg = c.workbook.getActiveCell().getResizedRange(grid.length - 1, grid[0].length - 1);
      rg.load('address'); rg.values = grid; await c.sync();
      const a = rg.address.slice(rg.address.lastIndexOf('!') + 1).replace(/\$/g, '');
      const prev = S.snips.find(s => s.sh === sh.name && s.a === a);
      if (chk) {
        const ok = S.mode === 'ok';
        rg.format.fill.color = ok ? '#C6EFCE' : '#FFC7CE';
        rg.format.font.color = ok ? '#006100' : '#9C0006';
        rg.format.font.bold = true; rg.format.horizontalAlignment = 'Center';
      } else {
        if ($('hl').checked) rg.format.fill.color = '#E2F0D9';
        if (prev && prev.k) { rg.format.font.color = '#000000'; rg.format.font.bold = false; rg.format.horizontalAlignment = 'General'; }
      }
      await c.sync();
      S.snips = S.snips.filter(s => !(s.sh === sh.name && s.a === a));
      S.snips.push(Object.assign({ sh: sh.name, a, d: S.cur.id, p: S.page, r: r.map(r1) }, chk ? { k: S.mode } : {}));
    });
    await saveSnips(); drawMarks();
    say(chk ? (S.mode === 'ok' ? 'Pointé : valide \u2713' : 'Pointé : invalide \u2717') : S.mode === 'table' ? `Tableau ${grid.length}×${grid[0].length} inséré.` : 'Valeur insérée.');
  } catch (e) { say('Erreur Excel : ' + e.message); }
}

/* ---------- Retrouver la source d'une cellule ---------- */
const colN = s => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
function parseA1(a) {
  const m = a.replace(/\$/g, '').match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/); if (!m) return null;
  return { c1: colN(m[1]), r1: +m[2], c2: colN(m[3] || m[1]), r2: +(m[4] || m[2]) };
}
async function onSel() {
  if (!$('follow').checked || !S.snips.length) return;
  try {
    await Excel.run(async c => {
      const r = c.workbook.getSelectedRange(); r.load('address'); await c.sync();
      const i = r.address.lastIndexOf('!'), sh = r.address.slice(0, i).replace(/^'|'$/g, '').replace(/''/g, "'"), a = parseA1(r.address.slice(i + 1));
      if (!a) return;
      const s = S.snips.find(s => { const q = parseA1(s.a); return s.sh === sh && q && a.c1 >= q.c1 && a.c1 <= q.c2 && a.r1 >= q.r1 && a.r1 <= q.r2; });
      if (!s) { if (S.focus) { S.focus = null; if (S.cur) drawMarks(); } return; }
      S.focus = { d: s.d, p: s.p, r: s.r };
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
  if (!S.cur || !confirm(`Retirer « ${S.cur.name} » du classeur ?`)) return;
  const d = S.cur;
  await Excel.run(async c => { c.workbook.customXmlParts.getItem(d.pid).delete(); await c.sync(); });
  S.docs = S.docs.filter(x => x !== d); S.snips = S.snips.filter(s => s.d !== d.id); await saveSnips();
  S.cur = null; S.pdf = null; fillDocs();
  if (S.docs.length) await openDoc(S.docs[0].id); else { $('wrap').hidden = true; $('empty').hidden = false; $('pn').textContent = '–'; }
};
$('modes').onclick = e => { const m = e.target.dataset.m; if (!m) return; S.mode = m; [...$('modes').children].forEach(b => b.classList.toggle('on', b === e.target)); };
$('prev').onclick = () => S.pdf && showPage(S.page - 1);
$('next').onclick = () => S.pdf && showPage(S.page + 1);
$('zi').onclick = () => { if (S.pdf) { S.zoom *= 1.2; showPage(S.page); } };
$('zo').onclick = () => { if (S.pdf) { S.zoom /= 1.2; showPage(S.page); } };
$('fit').onclick = async () => { if (!S.pdf) return; const p = await S.pdf.getPage(S.page); S.zoom = ($('view').clientWidth - 24) / p.getViewport({ scale: 1 }).width; showPage(S.page); };

let drag = null;
const ov = $('ov');
ov.addEventListener('pointerdown', e => {
  if (!S.pdf) return; ov.setPointerCapture(e.pointerId);
  const b = ov.getBoundingClientRect(); drag = { x: e.clientX - b.left, y: e.clientY - b.top, b };
  drag.el = document.createElement('div'); drag.el.className = 'sel'; ov.appendChild(drag.el);
});
ov.addEventListener('pointermove', e => {
  if (!drag) return; const x = e.clientX - drag.b.left, y = e.clientY - drag.b.top; drag.x2 = x; drag.y2 = y;
  Object.assign(drag.el.style, { left: Math.min(drag.x, x) + 'px', top: Math.min(drag.y, y) + 'px', width: Math.abs(x - drag.x) + 'px', height: Math.abs(y - drag.y) + 'px' });
});
ov.addEventListener('pointerup', async () => {
  if (!drag) return; const d = drag; drag = null; d.el.remove(); if (d.x2 == null) return;
  const z = S.zoom, r = [Math.min(d.x, d.x2) / z, Math.min(d.y, d.y2) / z, Math.max(d.x, d.x2) / z, Math.max(d.y, d.y2) / z];
  if (r[2] - r[0] < 3 || r[3] - r[1] < 3) return;
  await snip(r);
});

Office.onReady(async info => {
  if (info.host !== Office.HostType.Excel) { say('Ouvrez ce volet depuis Excel.'); return; }
  try {
    await loadAll(); fillDocs();
    if (S.docs.length) await openDoc(S.docs[0].id);
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSel);
    say(S.docs.length ? 'Classeur chargé.' : 'Prêt.');
  } catch (e) { say('Erreur d\'initialisation : ' + e.message); }
});
