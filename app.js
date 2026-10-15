/* Campaign Tracker — retail marketing elements report, per campaign.
   All data stays on the device (IndexedDB). No server needed. */
(function () {
  'use strict';

  // ---------- storage (IndexedDB) ----------
  const DB_NAME = 'campaign-tracker', STORE = 'campaigns';
  function openDb() {
    return new Promise((res, rej) => {
      const r = indexedDB.open(DB_NAME, 1);
      r.onupgradeneeded = () => r.result.createObjectStore(STORE, { keyPath: 'id' });
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
  }
  async function tx(mode, fn) {
    const db = await openDb();
    return new Promise((res, rej) => {
      const t = db.transaction(STORE, mode);
      const out = fn(t.objectStore(STORE));
      t.oncomplete = () => res(out && out.result !== undefined ? out.result : out);
      t.onerror = () => rej(t.error);
    });
  }
  const dbAll = () => tx('readonly', s => s.getAll());
  const dbGet = id => tx('readonly', s => s.get(id));
  const dbPut = c => tx('readwrite', s => s.put(c));
  const dbDel = id => tx('readwrite', s => s.delete(id));

  // ---------- helpers ----------
  const $ = id => document.getElementById(id);
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const nf = new Intl.NumberFormat('en-IN');
  const fmt = n => (n ? nf.format(n) : '–');
  const norm = v => String(v == null ? '' : v).trim().replace(/\s+/g, ' ');
  const keyOf = v => norm(v).toUpperCase();
  const toNum = v => {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    const n = Number(String(v).replace(/,/g, '').trim());
    return isFinite(n) ? n : 0;
  };
  const isNumLike = v => typeof v === 'number' || (typeof v === 'string' && v.trim() !== '' && isFinite(Number(v.replace(/,/g, ''))));
  let toastTimer;
  function toast(msg) {
    const t = $('toast'); t.textContent = msg; t.classList.remove('hidden');
    clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.add('hidden'), 3200);
  }
  const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
  const CAT_ORDER = ['S', 'A', 'B', 'C', 'D'];

  // ---------- views ----------
  function show(view) {
    ['home', 'upload', 'report'].forEach(v => $('view-' + v).classList.toggle('hidden', v !== view));
    window.scrollTo(0, 0);
  }

  // ---------- home ----------
  async function renderHome() {
    const list = (await dbAll()).sort((a, b) => b.updated - a.updated);
    $('emptyState').classList.toggle('hidden', list.length > 0);
    $('campaignList').innerHTML = list.map(c => {
      const st = summarize(c, null);
      return `<button class="card" data-id="${esc(c.id)}">
        <span class="card-title">${esc(c.name)}</span>
        <span class="muted small">${esc(c.fileName)} · updated ${new Date(c.updated).toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' })}</span>
        <span class="card-stats">
          <span><b>${fmt(st.stores)}</b><span>Stores</span></span>
          <span><b>${fmt(st.covered)}</b><span>Covered</span></span>
          <span><b>${fmt(st.qty)}</b><span>Element qty</span></span>
        </span></button>`;
    }).join('');
    $('campaignList').querySelectorAll('.card').forEach(b => b.onclick = () => openReport(b.dataset.id));
    show('home');
  }

  // ---------- upload & column mapping ----------
  let pending = null; // {wb, fileName, editing: campaign|null, replace: bool}

  function startUpload(editing, replace) {
    pending = { editing: editing || null, replace: !!replace };
    $('uploadTitle').textContent = editing ? (replace ? 'Replace data file' : 'Edit columns') : 'New campaign';
    $('campName').value = editing ? editing.name : '';
    $('fileInput').value = '';
    $('fileInfo').textContent = '';
    $('mapping').classList.add('hidden');
    $('fileInput').closest('.field').classList.toggle('hidden', !!editing && !replace);
    show('upload');
    if (editing && !replace) {
      // edit mapping using the stored raw table
      pending.raw = editing.raw; pending.fileName = editing.fileName;
      buildMappingFromRaw(editing.mapping);
    }
  }

  async function onFile(file) {
    if (!file) return;
    $('fileInfo').textContent = 'Reading ' + file.name + '…';
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array', cellDates: false });
      pending.wb = wb; pending.fileName = file.name;
      if (!$('campName').value.trim()) $('campName').value = file.name.replace(/\.[^.]+$/, '').replace(/[_]+/g, ' ');
      const sel = $('sheetSel');
      sel.innerHTML = wb.SheetNames.map((n, i) => `<option value="${i}">${esc(n)}</option>`).join('');
      // choose the sheet with most rows
      let best = 0, bestRows = -1;
      wb.SheetNames.forEach((n, i) => { const r = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, blankrows: false }).length; if (r > bestRows) { best = i; bestRows = r; } });
      sel.value = best;
      loadSheet(best);
    } catch (e) {
      console.error(e);
      $('fileInfo').textContent = 'Could not read this file. Please upload an .xlsx, .xls or .csv file.';
    }
  }

  function loadSheet(i) {
    const ws = pending.wb.Sheets[pending.wb.SheetNames[i]];
    pending.raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: false, raw: true });
    $('fileInfo').textContent = `${pending.fileName} · ${pending.raw.length.toLocaleString('en-IN')} rows read`;
    buildMappingFromRaw(null);
  }

  function detectHeaderRow(raw) {
    let best = 0, bestScore = -1;
    for (let r = 0; r < Math.min(20, raw.length); r++) {
      const row = raw[r] || [];
      const score = row.filter(v => typeof v === 'string' && v.trim() && !isNumLike(v)).length;
      if (score > bestScore) { best = r; bestScore = score; }
    }
    return best;
  }

  function columnsFor(raw, h) {
    const header = raw[h] || [];
    const width = Math.max(header.length, ...raw.slice(h + 1, h + 200).map(r => r.length));
    const seen = {};
    const cols = [];
    for (let c = 0; c < width; c++) {
      let name = norm(header[c]) || `Column ${c + 1}`;
      if (seen[name]) name += ` (${++seen[name]})`; else seen[name] = 1;
      let filled = 0, nums = 0; const distinct = new Set();
      for (let r = h + 1; r < raw.length; r++) {
        const v = raw[r][c];
        if (v == null || norm(v) === '') continue;
        filled++; if (isNumLike(v)) nums++;
        if (distinct.size <= 80) distinct.add(keyOf(v));
      }
      cols.push({ idx: c, name, filled, numeric: filled > 0 && nums / filled >= 0.9, distinct: distinct.size });
    }
    return cols;
  }

  function buildMappingFromRaw(existing) {
    const raw = pending.raw;
    const hSel = $('headerSel');
    const h = existing ? existing.headerRow : detectHeaderRow(raw);
    hSel.innerHTML = raw.slice(0, 20).map((row, i) => `<option value="${i}">Row ${i + 1}: ${esc((row || []).filter(v => v != null).slice(0, 4).join(' | ').slice(0, 60))}</option>`).join('');
    hSel.value = h;
    renderMapping(h, existing);
    $('mapping').classList.remove('hidden');
    $('sheetSel').closest('.field').classList.toggle('hidden', !pending.wb);
  }

  function renderMapping(h, existing) {
    const cols = columnsFor(pending.raw, h);
    pending.cols = cols; pending.headerRow = h;
    const opt = (c, sel) => `<option value="${c.idx}" ${sel ? 'selected' : ''}>${esc(c.name)}</option>`;
    const find = re => cols.find(c => re.test(c.name));
    const idCol = existing ? cols[existing.idCol] : (find(/^(ret|store|outlet|retailer)\s*(code|id)$/i) || find(/code|id/i));
    const zoneCol = existing ? cols[existing.zoneCol] : (find(/^zone/i) || find(/zone|region/i));
    $('idSel').innerHTML = '<option value="-1">(none — count rows)</option>' + cols.map(c => opt(c, idCol && c.idx === idCol.idx)).join('');
    $('zoneSel').innerHTML = '<option value="-1">(no zones)</option>' + cols.map(c => opt(c, zoneCol && c.idx === zoneCol.idx)).join('');

    // dimension candidates: low-cardinality text columns
    const prefer = [/^xi\s*state$/i, /^ret(ailer)?\s*type$/i, /^category$/i];
    const rank = c => { const i = prefer.findIndex(re => re.test(c.name)); return i === -1 ? 99 : i; };
    const dimCands = cols.filter(c => c.filled > 0 && !c.numeric && c.distinct <= 60).sort((a, b) => rank(a) - rank(b) || a.idx - b.idx);
    let dimSel = existing ? existing.dims : prefer.map(re => (dimCands.find(c => re.test(c.name)) || {}).idx).filter(v => v != null);
    if (!existing && !dimSel.length) dimSel = dimCands.filter(c => !zoneCol || c.idx !== zoneCol.idx).slice(0, 3).map(c => c.idx);
    $('dimChecks').innerHTML = dimCands.length ? dimCands.map(c => `<label class="check"><input type="checkbox" value="${c.idx}" ${dimSel.includes(c.idx) ? 'checked' : ''}>${esc(c.name)}</label>`).join('') : '<span class="muted small">No suitable text columns found.</span>';

    // element candidates: numeric columns, not codes; plus empty columns sitting after the first element column
    const notCode = c => !/code|\bpin\b|zip|\bid\b|latitude|longitude|^s\.?\s*no|sr\.?\s*no|serial/i.test(c.name);
    const numCols = cols.filter(c => c.numeric && notCode(c));
    const firstEl = numCols.length ? numCols[0].idx : Infinity;
    const elCands = cols.filter(c => notCode(c) && (c.numeric || (c.filled === 0 && c.idx > firstEl)));
    const elSel = existing ? existing.elements : elCands.map(c => c.idx);
    $('elChecks').innerHTML = elCands.length ? elCands.map(c => `<label class="check"><input type="checkbox" value="${c.idx}" ${elSel.includes(c.idx) ? 'checked' : ''}>${esc(c.name)}</label>`).join('') : '<span class="muted small">No numeric columns found.</span>';
    $('mapWarn').classList.add('hidden');
  }

  async function saveCampaign() {
    const name = $('campName').value.trim();
    const warn = m => { $('mapWarn').textContent = m; $('mapWarn').classList.remove('hidden'); };
    if (!name) return warn('Please enter a campaign name.');
    if (!pending || !pending.raw) return warn('Please upload a data file.');
    const checked = id => [...$(id).querySelectorAll('input:checked')].map(i => +i.value);
    const mapping = {
      headerRow: pending.headerRow,
      idCol: +$('idSel').value,
      zoneCol: +$('zoneSel').value,
      dims: checked('dimChecks'),
      elements: checked('elChecks'),
    };
    if (!mapping.elements.length) return warn('Select at least one element column.');
    const c = pending.editing && !pending.replace ? { ...pending.editing } : (pending.editing ? { ...pending.editing } : { id: uid(), created: Date.now() });
    c.name = name;
    c.fileName = pending.fileName;
    c.raw = pending.raw;
    c.mapping = mapping;
    c.updated = Date.now();
    delete c._model;
    try { await dbPut(c); } catch (e) { console.error(e); return warn('Could not save — the device may be out of storage.'); }
    toast('Campaign saved');
    openReport(c.id);
  }

  // ---------- model ----------
  // Build the working model from the stored raw table + mapping.
  function model(c) {
    if (c._model) return c._model;
    const m = c.mapping, cols = columnsFor(c.raw, m.headerRow);
    const name = i => (cols[i] || {}).name || `Column ${i + 1}`;
    const rows = [];
    const labelVotes = {}; // per dim: key -> {label: count}
    for (let r = m.headerRow + 1; r < c.raw.length; r++) {
      const row = c.raw[r];
      if (!row || row.every(v => v == null || norm(v) === '')) continue;
      const id = m.idCol >= 0 ? norm(row[m.idCol]) : String(r);
      if (m.idCol >= 0 && !id) continue; // skip total / blank rows without a store code
      const zone = m.zoneCol >= 0 ? keyOf(row[m.zoneCol]) || '(BLANK)' : 'ALL';
      const dims = m.dims.map(d => {
        const k = keyOf(row[d]) || '(BLANK)';
        const lv = (labelVotes[d] = labelVotes[d] || {});
        const lab = norm(row[d]) || '(Blank)';
        (lv[k] = lv[k] || {})[lab] = ((lv[k] || {})[lab] || 0) + 1;
        return k;
      });
      const els = m.elements.map(e => toNum(row[e]));
      rows.push({ id, zone, dims, els, total: els.reduce((a, b) => a + b, 0) });
    }
    const label = (d, k) => { const v = (labelVotes[d] || {})[k]; return v ? Object.keys(v).sort((a, b) => v[b] - v[a])[0] : k; };
    const zoneVotes = {};
    if (m.zoneCol >= 0) for (let r = m.headerRow + 1; r < c.raw.length; r++) { const row = c.raw[r]; if (!row) continue; const k = keyOf(row[m.zoneCol]); if (!k) continue; const l = norm(row[m.zoneCol]); (zoneVotes[k] = zoneVotes[k] || {})[l] = (zoneVotes[k][l] || 0) + 1; }
    const zoneLabel = k => zoneVotes[k] ? Object.keys(zoneVotes[k]).sort((a, b) => zoneVotes[k][b] - zoneVotes[k][a])[0] : k;
    const zones = [...new Set(rows.map(r => r.zone))].sort();
    c._model = { rows, zones, zoneLabel, dimNames: m.dims.map(name), elNames: m.elements.map(name), label: (i, k) => label(m.dims[i], k) };
    return c._model;
  }

  function summarize(c, zone) {
    const md = model(c);
    const rows = zone ? md.rows.filter(r => r.zone === zone) : md.rows;
    const ids = new Set(), cov = new Set(); let qty = 0;
    rows.forEach(r => { ids.add(r.id); if (r.total > 0) cov.add(r.id); qty += r.total; });
    return { stores: ids.size, covered: cov.size, qty, rows };
  }

  // aggregate rows by a key function -> Map key -> {stores, covered, qty[], st[]}
  function aggregate(rows, nEl, keyFn) {
    const out = new Map();
    for (const r of rows) {
      const k = keyFn(r);
      let g = out.get(k);
      if (!g) { g = { ids: new Set(), cov: new Set(), qty: new Array(nEl).fill(0), st: new Array(nEl).fill(0).map(() => new Set()) }; out.set(k, g); }
      g.ids.add(r.id); if (r.total > 0) g.cov.add(r.id);
      r.els.forEach((v, i) => { if (v) { g.qty[i] += v; if (v > 0) g.st[i].add(r.id); } });
    }
    for (const g of out.values()) { g.stores = g.ids.size; g.covered = g.cov.size; g.stc = g.st.map(s => s.size); }
    return out;
  }

  // ---------- report ----------
  const state = { id: null, zone: null, dim: 0 };

  async function openReport(id) {
    const c = await dbGet(id);
    if (!c) return renderHome();
    state.c = c; state.id = id;
    const md = model(c);
    if (!md.zones.includes(state.zone)) state.zone = null;
    if (state.dim >= md.dimNames.length) state.dim = 0;
    $('repTitle').textContent = c.name;
    $('repMeta').textContent = `${c.fileName} · updated ${new Date(c.updated).toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' })}`;
    const chips = [{ k: null, l: 'All zones' }].concat(md.zones.length > 1 ? md.zones.map(z => ({ k: z, l: md.zoneLabel(z) })) : []);
    $('zoneChips').innerHTML = chips.map(z => `<button class="chip" role="tab" data-z="${esc(z.k == null ? '' : z.k)}" aria-selected="${z.k === state.zone}">${esc(z.l)}</button>`).join('');
    $('zoneChips').querySelectorAll('.chip').forEach(b => b.onclick = () => { state.zone = b.dataset.z || null; openReport(id); });
    $('dimTabs').innerHTML = md.dimNames.map((n, i) => `<button class="tab" role="tab" data-i="${i}" aria-selected="${i === state.dim}">${esc(n)}</button>`).join('') || '<span class="muted small" style="padding:6px">No “report by” columns selected</span>';
    $('dimTabs').querySelectorAll('.tab').forEach(b => b.onclick = () => { state.dim = +b.dataset.i; openReport(id); });
    renderReport();
    show('report');
  }

  function activeElements(c, rows) {
    const md = model(c);
    const totals = md.elNames.map((_, i) => rows.reduce((a, r) => a + r.els[i], 0));
    // hide elements with no quantity anywhere in the campaign
    const all = md.elNames.map((_, i) => md.rows.reduce((a, r) => a + r.els[i], 0));
    return { idx: md.elNames.map((_, i) => i).filter(i => all[i] !== 0), hidden: md.elNames.filter((_, i) => all[i] === 0), totals };
  }

  function sortKeys(keys, dimName) {
    const isCat = /categ/i.test(dimName);
    return keys.sort((a, b) => {
      if (isCat) { const ia = CAT_ORDER.indexOf(a), ib = CAT_ORDER.indexOf(b); if (ia !== -1 || ib !== -1) return (ia === -1 ? 99 : ia) - (ib === -1 ? 99 : ib); }
      return a.localeCompare(b);
    });
  }

  function renderReport() {
    const c = state.c, md = model(c);
    const s = summarize(c, state.zone);
    $('kStores').textContent = fmt(s.stores);
    $('kCovered').textContent = fmt(s.covered);
    $('kCovPct').textContent = s.stores ? Math.round(s.covered / s.stores * 100) + '% of stores' : '';
    $('kQty').textContent = fmt(s.qty);

    const ae = activeElements(c, s.rows);
    const tot = aggregate(s.rows, md.elNames.length, () => 'T').get('T') || { qty: [], stc: [] };
    const max = Math.max(1, ...ae.idx.map(i => tot.qty[i] || 0));
    $('elBars').innerHTML = ae.idx.map(i => {
      const q = tot.qty[i] || 0, st = tot.stc[i] || 0;
      return `<div class="bar-item"><span class="bar-label" title="${esc(md.elNames[i])}">${esc(md.elNames[i])}</span>
        <span class="bar-val">${fmt(q)} <span>· ${fmt(st)} stores</span></span>
        <span class="bar-track"><span class="bar-fill" style="display:block;width:${(q / max * 100).toFixed(1)}%"></span></span></div>`;
    }).join('');

    // table
    const show = $('showSel').value, orient = $('orientSel').value;
    const t = $('repTable');
    if (!md.dimNames.length) { t.innerHTML = ''; $('tableNote').textContent = ''; return; }
    const d = state.dim, dimName = md.dimNames[d];
    const groups = aggregate(s.rows, md.elNames.length, r => r.dims[d]);
    const keys = sortKeys([...groups.keys()], dimName);
    const cellHtml = (q, st) => {
      if (!q && !st) return '<td class="zero">–</td>';
      if (show === 'qty') return `<td>${fmt(q)}</td>`;
      if (show === 'stores') return `<td>${fmt(st)}</td>`;
      return `<td>${fmt(q)} <span class="s">(${fmt(st)})</span></td>`;
    };
    const cols = keys.map(k => ({ label: md.label(d, k), g: groups.get(k) })).concat([{ label: 'Total', g: tot.ids ? tot : aggregate(s.rows, md.elNames.length, () => 'T').get('T'), total: true }]);
    if (orient === 'el-rows') {
      let h = `<thead><tr><th>Element</th>${cols.map(c2 => `<th>${esc(c2.label)}</th>`).join('')}</tr></thead><tbody>`;
      h += `<tr class="sub"><td>Total stores</td>${cols.map(c2 => `<td>${fmt(c2.g ? c2.g.stores : 0)}</td>`).join('')}</tr>`;
      h += `<tr class="sub"><td>Covered stores</td>${cols.map(c2 => `<td>${fmt(c2.g ? c2.g.covered : 0)}</td>`).join('')}</tr>`;
      ae.idx.forEach(i => { h += `<tr><td>${esc(md.elNames[i])}</td>${cols.map(c2 => c2.g ? cellHtml(c2.g.qty[i], c2.g.stc[i]) : '<td class="zero">–</td>').join('')}</tr>`; });
      t.innerHTML = h + '</tbody>';
    } else {
      let h = `<thead><tr><th>${esc(dimName)}</th><th>Total stores</th><th>Covered stores</th>${ae.idx.map(i => `<th>${esc(md.elNames[i])}</th>`).join('')}</tr></thead><tbody>`;
      cols.forEach(c2 => {
        const g = c2.g || { stores: 0, covered: 0, qty: [], stc: [] };
        h += `<tr class="${c2.total ? 'total' : ''}"><td>${esc(c2.label)}</td><td>${fmt(g.stores)}</td><td>${fmt(g.covered)}</td>${ae.idx.map(i => cellHtml(g.qty[i], g.stc[i])).join('')}</tr>`;
      });
      t.innerHTML = h + '</tbody>';
    }
    const notes = [];
    if (show === 'both') notes.push('Each cell: element qty (stores receiving it).');
    notes.push('Covered stores = stores receiving at least one element.');
    notes.push('Names that differ only in capital letters are merged.');
    if (ae.hidden.length) notes.push('Not shown (no quantity planned): ' + ae.hidden.join(', ') + '.');
    $('tableNote').textContent = notes.join(' ');
  }

  // ---------- export ----------
  function buildWorkbook(c, zone) {
    const md = model(c), s = summarize(c, zone);
    const ae = activeElements(c, s.rows);
    const wb = XLSX.utils.book_new();
    const zl = zone ? md.zoneLabel(zone) : 'All zones';
    const tot = aggregate(s.rows, md.elNames.length, () => 'T').get('T');
    const aoa = [[`${c.name} — ${zl}`], [`Total stores: ${s.stores}`, '', `Covered stores: ${s.covered}`, '', `Total element qty: ${s.qty}`], []];
    md.dimNames.forEach((dn, d) => {
      const groups = aggregate(s.rows, md.elNames.length, r => r.dims[d]);
      const keys = sortKeys([...groups.keys()], dn);
      aoa.push([`${dn} wise — element qty & store count`]);
      aoa.push([dn, 'Total stores', 'Covered stores'].concat(...ae.idx.map(i => [md.elNames[i] + ' Qty', md.elNames[i] + ' Stores'])));
      const line = (lab, g) => [lab, g.stores, g.covered].concat(...ae.idx.map(i => [g.qty[i], g.stc[i]]));
      keys.forEach(k => aoa.push(line(md.label(d, k), groups.get(k))));
      if (tot) aoa.push(line('Grand Total', tot));
      aoa.push([]);
    });
    if (ae.hidden.length) aoa.push(['Not shown (no quantity planned): ' + ae.hidden.join(', ')]);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 26 }, { wch: 12 }, { wch: 14 }].concat(ae.idx.flatMap(() => [{ wch: 10 }, { wch: 10 }]));
    XLSX.utils.book_append_sheet(wb, ws, 'Summary');
    // store data for this zone
    const m = c.mapping, header = c.raw[m.headerRow] || [];
    const data = [header].concat(c.raw.slice(m.headerRow + 1).filter(r => {
      if (!r || (m.idCol >= 0 && !norm(r[m.idCol]))) return false;
      return !zone || (m.zoneCol >= 0 && keyOf(r[m.zoneCol]) === zone);
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data), 'Store Data');
    return wb;
  }

  async function saveFile(fileName, base64, mime) {
    const cap = window.Capacitor;
    if (cap && cap.isNativePlatform && cap.isNativePlatform() && cap.Plugins && cap.Plugins.Filesystem) {
      const { Filesystem, Share } = cap.Plugins;
      const res = await Filesystem.writeFile({ path: fileName, data: base64, directory: 'CACHE' });
      if (Share) await Share.share({ title: fileName, files: [res.uri], dialogTitle: 'Save or share ' + fileName });
      else toast('Saved to app storage');
      return;
    }
    const bin = atob(base64), bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    const url = URL.createObjectURL(new Blob([bytes], { type: mime }));
    const a = document.createElement('a'); a.href = url; a.download = fileName; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }
  const safeName = s => s.replace(/[\\/:*?"<>|]+/g, '-').trim();
  const b64utf8 = s => btoa(unescape(encodeURIComponent(s)));

  // ---------- events ----------
  $('homeBtn').onclick = renderHome;
  $('newBtn').onclick = () => startUpload(null, false);
  $('fileInput').onchange = e => onFile(e.target.files[0]);
  $('sheetSel').onchange = e => loadSheet(+e.target.value);
  $('headerSel').onchange = e => renderMapping(+e.target.value, null);
  $('cancelUpload').onclick = () => (pending && pending.editing ? openReport(pending.editing.id) : renderHome());
  $('saveCampaign').onclick = saveCampaign;
  $('showSel').onchange = renderReport;
  $('orientSel').onchange = renderReport;
  $('editBtn').onclick = () => startUpload(state.c, false);
  $('replaceBtn').onclick = () => startUpload(state.c, true);
  $('deleteBtn').onclick = async () => {
    if (!confirm(`Delete campaign “${state.c.name}” from this device? This cannot be undone.`)) return;
    await dbDel(state.id); toast('Campaign deleted'); renderHome();
  };
  $('exportBtn').onclick = async () => {
    try {
      const md = model(state.c);
      const zl = state.zone ? md.zoneLabel(state.zone) : 'All zones';
      const wb = buildWorkbook(state.c, state.zone);
      const b64 = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });
      await saveFile(safeName(`${state.c.name} - ${zl}.xlsx`), b64, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    } catch (e) { console.error(e); toast('Export failed'); }
  };
  $('backupBtn').onclick = async () => {
    const all = (await dbAll()).map(({ _model, ...c }) => c);
    if (!all.length) return toast('No campaigns to back up');
    const json = JSON.stringify({ app: 'campaign-tracker', version: 1, exported: new Date().toISOString(), campaigns: all });
    await saveFile(`campaign-tracker-backup-${new Date().toISOString().slice(0, 10)}.json`, b64utf8(json), 'application/json');
  };
  $('importInput').onchange = async e => {
    const f = e.target.files[0]; if (!f) return;
    try {
      const data = JSON.parse(await f.text());
      if (data.app !== 'campaign-tracker' || !Array.isArray(data.campaigns)) throw new Error('bad file');
      for (const c of data.campaigns) { if (c.id && c.raw && c.mapping) await dbPut(c); }
      toast(`Imported ${data.campaigns.length} campaign(s)`);
      renderHome();
    } catch (err) { toast('This is not a Campaign Tracker backup file'); }
    e.target.value = '';
  };

  if ('serviceWorker' in navigator && location.protocol === 'https:') {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }
  renderHome();
})();
