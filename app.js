/**
 * Literature Screening Pipeline — Client-Side Engine (v2)
 *
 * Two-file merge workflow:
 *   - Master file (accumulated history, optional)
 *   - Weekly input (new batch, required)
 *   - Three-type duplicate detection
 *   - Process only genuinely new entries via APIs
 *   - Generate updated master Excel for next week
 */

/* ===== CONSTANTS ===== */
const CROSSREF_API = 'https://api.crossref.org/works/';
const UNPAYWALL_API = 'https://api.unpaywall.org/v2/';
const REQUEST_DELAY = 600;

/* ===== STATE ===== */
let masterFile = null;
let weeklyFile = null;
let pipelineData = null;

/* ===== INIT ===== */
document.addEventListener('DOMContentLoaded', () => {
  setupDropzone('master-dropzone', 'master-file-input', 'master-file-name', 'master');
  setupDropzone('weekly-dropzone', 'weekly-file-input', 'weekly-file-name', 'weekly');
  setupTabs();
  setupDownloads();
  document.getElementById('email-input').addEventListener('input', updateRunBtn);
  document.getElementById('run-btn').addEventListener('click', handleRun);
});

/* ===== DROPZONE ===== */
function setupDropzone(dzId, fiId, fnId, type) {
  const dz = document.getElementById(dzId);
  const fi = document.getElementById(fiId);

  dz.addEventListener('click', () => fi.click());
  dz.addEventListener('keydown', e => { if (e.key === 'Enter') fi.click(); });
  dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
  dz.addEventListener('drop', e => {
    e.preventDefault();
    dz.classList.remove('drag-over');
    if (e.dataTransfer.files.length) pickFile(type, e.dataTransfer.files[0], fnId);
  });
  fi.addEventListener('change', () => { if (fi.files.length) pickFile(type, fi.files[0], fnId); });
}

function pickFile(type, file, fnId) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx', 'xls', 'csv'].includes(ext)) {
    showToast('Please select an .xlsx, .xls, or .csv file.');
    return;
  }
  if (type === 'master') masterFile = file;
  else weeklyFile = file;

  const el = document.getElementById(fnId);
  el.textContent = '✅ ' + file.name;
  el.style.display = 'block';
  updateRunBtn();
}

function updateRunBtn() {
  document.getElementById('run-btn').disabled =
    !(weeklyFile && document.getElementById('email-input').value.trim());
}

/* ===== TABS ===== */
function setupTabs() {
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById('tab-' + btn.dataset.tab).classList.add('active');
    });
  });
}

/* ===== DOWNLOADS ===== */
function setupDownloads() {
  document.getElementById('dl-master').addEventListener('click', e => {
    e.preventDefault(); if (pipelineData) generateUpdatedMasterExcel(pipelineData);
  });
  document.getElementById('dl-screening').addEventListener('click', e => {
    e.preventDefault(); if (pipelineData) generateScreeningExcel(pipelineData);
  });
  document.getElementById('dl-log').addEventListener('click', e => {
    e.preventDefault(); if (pipelineData) generateLogExcel(pipelineData);
  });
  document.getElementById('dl-purchase').addEventListener('click', e => {
    e.preventDefault(); if (pipelineData) generatePurchaseExcel(pipelineData);
  });
}

/* ===== MAIN RUN HANDLER ===== */
async function handleRun() {
  if (!weeklyFile) return;
  const email = document.getElementById('email-input').value.trim();
  if (!email) return;
  const batchName = document.getElementById('batch-input').value.trim()
    || `Batch ${new Date().toLocaleDateString()}`;

  const overlay = document.getElementById('processing-overlay');
  const status = document.getElementById('processing-status');
  const bar = document.getElementById('progress-bar');

  overlay.classList.add('active');
  status.textContent = 'Reading weekly input…';
  bar.style.width = '0%';
  document.getElementById('results-section').classList.remove('visible');

  try {
    /* 1. Parse files */
    const weeklyRecords = await parseFile(weeklyFile);
    if (!weeklyRecords.length) throw new Error('No records found in weekly input.');

    let masterRecords = [];
    if (masterFile) {
      status.textContent = 'Reading master file…';
      masterRecords = await parseFile(masterFile, false); // don't require strict columns
    }
    bar.style.width = '5%';

    /* 2. Merge & detect duplicates */
    status.textContent = 'Merging and detecting duplicates…';
    const merged = mergeAndDetectDuplicates(masterRecords, weeklyRecords, batchName);
    bar.style.width = '10%';

    /* 3. Process only new entries via APIs */
    const toProcess = merged.newEntries.filter(r => !r._skipProcessing);
    for (let i = 0; i < toProcess.length; i++) {
      const row = toProcess[i];
      status.textContent = `[${i + 1}/${toProcess.length}] Processing ${row.UI}…`;
      bar.style.width = `${10 + (i / toProcess.length) * 80}%`;

      const result = await processRecord(row, email);
      // Merge API results back into the record
      Object.assign(row, result);

      if (i < toProcess.length - 1) await sleep(REQUEST_DELAY);
    }

    /* 4. Build purchase list (new entries only) */
    const purchase = merged.newEntries
      .filter(r => r.Status === 'PURCHASE_REQUIRED')
      .map(r => ({
        UI: r.UI, Title: r.Title, DOI: r.DOI,
        Publisher: r.Publisher || '',
        'Website/site link for purchase': r['Publisher URL'] || '',
        'Purchase Cost': r['Purchase Cost'] || '',
      }));

    bar.style.width = '95%';
    status.textContent = 'Rendering results…';

    pipelineData = { ...merged, purchase };
    renderResults(pipelineData);

    bar.style.width = '100%';
    await sleep(300);
    overlay.classList.remove('active');

    const rs = document.getElementById('results-section');
    rs.classList.add('visible');
    rs.scrollIntoView({ behavior: 'smooth', block: 'start' });

  } catch (err) {
    overlay.classList.remove('active');
    showToast(err.message);
    console.error(err);
  }
}

/* ===== FILE PARSING (SheetJS) ===== */
function parseFile(file, strict = true) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const sheetInput = document.getElementById('sheet-input').value.trim();
        const sheetName = sheetInput || wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        if (!ws) {
          reject(new Error(`Sheet "${sheetName}" not found. Available: ${wb.SheetNames.join(', ')}`));
          return;
        }
        const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
        const records = json.map(row => {
          const obj = {};
          for (const [key, val] of Object.entries(row)) {
            obj[key.trim()] = String(val).trim();
          }
          return obj;
        });
        if (strict && records.length > 0) {
          const cols = Object.keys(records[0]);
          const required = ['UI', 'Title', 'DOI'];
          const missing = required.filter(c => !cols.includes(c));
          if (missing.length) {
            reject(new Error(`Missing required columns: ${missing.join(', ')}. Found: ${cols.join(', ')}`));
            return;
          }
        }
        resolve(records);
      } catch (err) {
        reject(new Error('Failed to parse file: ' + err.message));
      }
    };
    reader.onerror = () => reject(new Error('Failed to read file.'));
    reader.readAsArrayBuffer(file);
  });
}

/* ===== NORMALIZE ===== */
function normalizeDoi(val) {
  let doi = (val || '').trim().toLowerCase();
  if (!doi) return '';
  for (const pfx of [
    'https://doi.org/', 'http://doi.org/',
    'https://dx.doi.org/', 'http://dx.doi.org/',
  ]) { doi = doi.replace(pfx, ''); }
  doi = doi.replace(/^\s*doi\s*:\s*/, '');
  return doi.trim();
}

function normalizeTitle(val) {
  let t = (val || '').trim().toLowerCase();
  t = t.replace(/[^\w\s]/g, ' ');
  return t.replace(/\s+/g, ' ').trim();
}

/* ===== MERGE & DUPLICATE DETECTION ===== */
function mergeAndDetectDuplicates(masterRecords, weeklyRecords, batchName) {
  const today = new Date().toISOString().split('T')[0];

  // Build master lookup sets
  const masterDoiSet = new Set();
  const masterTitleSet = new Set();
  masterRecords.forEach(r => {
    const nd = normalizeDoi(r.DOI);
    const nt = normalizeTitle(r.Title);
    if (nd) masterDoiSet.add(nd);
    if (nt) masterTitleSet.add(nt);
  });

  // Prepare master rows (gray, already processed)
  const processedMaster = masterRecords.map(r => ({
    ...r,
    _source: 'Master',
    _rowType: 'master',
    _skipProcessing: true,
  }));

  // Normalize weekly records
  const weeklyNorm = weeklyRecords.map(r => ({
    ...r,
    _ndoi: normalizeDoi(r.DOI),
    _ntitle: normalizeTitle(r.Title),
    _source: 'New',
    'Date Added': today,
    Batch: batchName,
  }));

  // Count within-batch occurrences
  const bDoiCnt = {}, bTitleCnt = {};
  weeklyNorm.forEach(r => {
    if (r._ndoi) bDoiCnt[r._ndoi] = (bDoiCnt[r._ndoi] || 0) + 1;
    if (r._ntitle) bTitleCnt[r._ntitle] = (bTitleCnt[r._ntitle] || 0) + 1;
  });

  // Track first-seen for batch dups
  const seenDois = new Set();
  const seenTitles = new Set();

  let masterDupCount = 0, batchDupCount = 0, newCount = 0;

  weeklyNorm.forEach(r => {
    const nd = r._ndoi, nt = r._ntitle;

    // 1. Check against master
    const inMasterDoi = nd && masterDoiSet.has(nd);
    const inMasterTitle = nt && masterTitleSet.has(nt);
    if (inMasterDoi || inMasterTitle) {
      r._rowType = 'masterDup';
      r._skipProcessing = true;
      r.Status = 'ALREADY_IN_MASTER';
      r.Notes = 'Already exists in master file';
      r.Duplicate = true;
      const reasons = [];
      if (inMasterDoi) reasons.push('DOI in master');
      if (inMasterTitle) reasons.push('Title in master');
      r['Duplicate Reason'] = reasons.join('; ');
      masterDupCount++;
      return;
    }

    // 2. Check within batch
    const bDoiDup = nd && (bDoiCnt[nd] > 1);
    const bTitleDup = nt && (bTitleCnt[nt] > 1);
    if (bDoiDup || bTitleDup) {
      r._rowType = 'batchDup';
      r.Duplicate = true;
      const reasons = [];
      if (bDoiDup) reasons.push('Duplicate DOI in batch');
      if (bTitleDup) reasons.push('Duplicate Title in batch');
      r['Duplicate Reason'] = reasons.join('; ');

      // Process first occurrence, skip subsequent
      const firstDoi = nd && !seenDois.has(nd);
      const firstTitle = nt && !seenTitles.has(nt);
      if (firstDoi || firstTitle) {
        r._skipProcessing = false;
        if (nd) seenDois.add(nd);
        if (nt) seenTitles.add(nt);
      } else {
        r._skipProcessing = true;
        r.Status = 'BATCH_DUPLICATE';
        r.Notes = 'Duplicate within this batch — see earlier entry';
      }
      batchDupCount++;
      return;
    }

    // 3. Genuinely new
    if (nd) seenDois.add(nd);
    if (nt) seenTitles.add(nt);
    r._rowType = 'new';
    r._skipProcessing = false;
    r.Duplicate = false;
    r['Duplicate Reason'] = '';
    newCount++;
  });

  // Combine: master first, then weekly entries
  const allRecords = [...processedMaster, ...weeklyNorm];
  const newEntries = weeklyNorm; // all weekly entries (includes masterDups, batchDups, new)

  const stats = {
    masterCount: masterRecords.length,
    weeklyCount: weeklyRecords.length,
    masterDupCount,
    batchDupCount,
    newCount,
    toProcessCount: weeklyNorm.filter(r => !r._skipProcessing).length,
    totalCombined: allRecords.length,
  };

  return { allRecords, newEntries, stats, batchName };
}

/* ===== API LOOKUPS ===== */
async function crossrefLookup(doi) {
  if (!doi) return null;
  try {
    const r = await fetch(CROSSREF_API + encodeURIComponent(doi), {
      headers: { 'Accept': 'application/json' },
    });
    if (!r.ok) return null;
    const json = await r.json();
    return json.message || null;
  } catch { return null; }
}

async function unpaywallLookup(doi, email) {
  if (!doi) return null;
  try {
    const r = await fetch(
      `${UNPAYWALL_API}${encodeURIComponent(doi)}?email=${encodeURIComponent(email)}`,
      { headers: { 'Accept': 'application/json' } }
    );
    if (!r.ok) return null;
    return await r.json();
  } catch { return null; }
}

function extractOaCandidates(data) {
  if (!data) return [];
  const out = [];
  const best = data.best_oa_location;
  if (best && best.url_for_pdf) {
    out.push({ url: best.url_for_pdf, source: 'Unpaywall best OA location' });
  }
  for (const loc of data.oa_locations || []) {
    const url = loc.url_for_pdf;
    if (url && !out.some(x => x.url === url)) {
      out.push({ url, source: loc.host_type || 'Unpaywall OA location' });
    }
  }
  return out;
}

/* ===== PROCESS SINGLE RECORD ===== */
async function processRecord(row, email) {
  const doi = normalizeDoi(row.DOI);
  const result = {
    Status: '', 'OA Source': '', 'PDF URL': '',
    Publisher: '', 'Publisher URL': '', 'Purchase Cost': '', Notes: '',
  };

  if (!doi) {
    result.Status = 'MANUAL_REVIEW';
    result.Notes = 'DOI missing';
    return result;
  }

  const crossref = await crossrefLookup(doi);
  if (crossref) {
    result.Publisher = crossref.publisher || '';
    result['Publisher URL'] = crossref.URL || `https://doi.org/${doi}`;
  } else {
    result['Publisher URL'] = `https://doi.org/${doi}`;
  }

  const oa = await unpaywallLookup(doi, email);
  if (!oa) {
    result.Status = 'MANUAL_REVIEW';
    result.Notes = 'Unable to determine OA status';
    return result;
  }

  const candidates = extractOaCandidates(oa);
  if (candidates.length > 0) {
    result.Status = 'OA_AVAILABLE';
    result['OA Source'] = candidates[0].source;
    result['PDF URL'] = candidates[0].url;
    result.Notes = 'Open-access PDF link available.';
    return result;
  }

  if (oa.is_oa === false) {
    result.Status = 'PURCHASE_REQUIRED';
    result['Purchase Cost'] = 'Price not publicly available';
    result.Notes = 'No legitimate open-access location found.';
    return result;
  }

  result.Status = 'MANUAL_REVIEW';
  result.Notes = 'No downloadable PDF URL was found.';
  return result;
}

/* ===== RENDER RESULTS ===== */
function renderResults(data) {
  renderStats(data.stats);
  renderAllTable(data.allRecords);
  renderNewTable(data.newEntries);
  renderPurchaseTable(data.purchase);
}

function renderStats(s) {
  const bar = document.getElementById('stats-bar');
  bar.innerHTML = '';
  addStat(bar, s.totalCombined, 'Total Combined');
  addStat(bar, s.masterCount, 'In Master');
  addStat(bar, s.weeklyCount, 'Weekly Input');
  addStat(bar, s.masterDupCount, '🟠 In Master Already');
  addStat(bar, s.batchDupCount, '🟡 Batch Duplicates');
  addStat(bar, s.newCount, '🆕 Genuinely New');
  addStat(bar, s.toProcessCount, 'API Lookups');
}

function addStat(container, count, label) {
  const d = document.createElement('div');
  d.className = 'stat-item';
  d.innerHTML = `<span class="stat-count">${count}</span><span class="stat-label">${label}</span>`;
  container.appendChild(d);
}

function rowClass(row) {
  const t = row._rowType;
  if (t === 'master') return 'master-row';
  if (t === 'masterDup') return 'master-dup-row';
  if (t === 'batchDup') return 'batch-dup-row';
  return '';
}

function sourceBadge(row) {
  const t = row._rowType;
  if (t === 'master') return '<span class="badge badge-master">Master</span>';
  if (t === 'masterDup') return '<span class="badge badge-master-dup">In Master</span>';
  if (t === 'batchDup') return '<span class="badge badge-batch-dup">Batch Dup</span>';
  return '<span class="badge badge-downloaded">New</span>';
}

function statusBadge(status) {
  if (!status) return '—';
  return `<span class="badge ${badgeClass(status)}">${formatStatus(status)}</span>`;
}

function renderAllTable(rows) {
  const body = document.getElementById('all-body');
  body.innerHTML = '';
  rows.forEach(row => {
    const tr = document.createElement('tr');
    const cls = rowClass(row);
    if (cls) tr.className = cls;
    tr.innerHTML = `
      <td>${sourceBadge(row)}</td>
      <td>${esc(row.UI)}</td>
      <td>${esc(row.Title)}</td>
      <td>${doiLink(row.DOI)}</td>
      <td>${statusBadge(row.Status)}</td>
      <td>${row.Duplicate ? '⚠️' : '—'}</td>
      <td>${esc(row['Duplicate Reason'] || '')}</td>
      <td>${esc(row.Batch || '')}</td>`;
    body.appendChild(tr);
  });
}

function renderNewTable(rows) {
  const body = document.getElementById('new-body');
  body.innerHTML = '';
  rows.forEach(row => {
    const tr = document.createElement('tr');
    const cls = rowClass(row);
    if (cls) tr.className = cls;
    tr.innerHTML = `
      <td>${esc(row.UI)}</td>
      <td>${esc(row.Title)}</td>
      <td>${doiLink(row.DOI)}</td>
      <td>${statusBadge(row.Status)}</td>
      <td>${esc(row['OA Source'] || '')}</td>
      <td>${esc(row.Publisher || '')}</td>
      <td>${esc(row.Notes || '')}</td>`;
    body.appendChild(tr);
  });
}

function renderPurchaseTable(rows) {
  const body = document.getElementById('purchase-body');
  const wrap = document.getElementById('purchase-wrapper');
  const empty = document.getElementById('purchase-empty');
  body.innerHTML = '';
  if (!rows.length) {
    wrap.style.display = 'none';
    empty.style.display = 'block';
    return;
  }
  wrap.style.display = 'block';
  empty.style.display = 'none';
  rows.forEach(row => {
    const tr = document.createElement('tr');
    const link = row['Website/site link for purchase'] || '';
    tr.innerHTML = `
      <td>${esc(row.UI)}</td>
      <td>${esc(row.Title)}</td>
      <td>${doiLink(row.DOI)}</td>
      <td>${esc(row.Publisher || '')}</td>
      <td>${link ? `<a href="${esc(link)}" target="_blank" rel="noopener" style="color:var(--accent)">${esc(link)}</a>` : '—'}</td>
      <td>${esc(row['Purchase Cost'] || '')}</td>`;
    body.appendChild(tr);
  });
}

/* ===== EXCEL EXPORT ===== */
const HEADER_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2D3142' } };
const HEADER_FONT = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
const YELLOW_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
const ORANGE_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE0CC' } };
const GRAY_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E8E8' } };

function fillForRow(row) {
  const t = row._rowType;
  if (t === 'masterDup') return ORANGE_FILL;
  if (t === 'batchDup') return YELLOW_FILL;
  if (t === 'master') return GRAY_FILL;
  return null;
}

async function generateUpdatedMasterExcel(data) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Master');
  ws.columns = [
    { header: 'UI', key: 'UI', width: 18 },
    { header: 'Title', key: 'Title', width: 80 },
    { header: 'DOI', key: 'DOI', width: 45 },
    { header: 'Status', key: 'Status', width: 22 },
    { header: 'Publisher', key: 'Publisher', width: 30 },
    { header: 'OA Source', key: 'OASource', width: 30 },
    { header: 'Date Added', key: 'DateAdded', width: 14 },
    { header: 'Batch', key: 'Batch', width: 20 },
  ];
  styleHeader(ws);

  // Only include master rows + genuinely new/processed entries (skip master dups & batch dup copies)
  const seenDois = new Set();
  data.allRecords.forEach(row => {
    const nd = normalizeDoi(row.DOI);
    if (row._rowType === 'masterDup') return; // already in master, skip
    if (row._rowType === 'batchDup' && row._skipProcessing) return; // duplicate copy, skip
    if (nd && seenDois.has(nd)) return;
    if (nd) seenDois.add(nd);

    ws.addRow({
      UI: row.UI, Title: row.Title, DOI: row.DOI || '',
      Status: row.Status || '', Publisher: row.Publisher || '',
      OASource: row['OA Source'] || '',
      DateAdded: row['Date Added'] || '', Batch: row.Batch || '',
    });
  });

  ws.views = [{ state: 'frozen', ySplit: 1 }];
  ws.autoFilter = { from: 'A1', to: `H${ws.rowCount}` };
  saveWorkbook(wb, 'Updated_Master.xlsx');
}

async function generateScreeningExcel(data) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('All Records');
  ws.columns = [
    { header: 'Source', key: 'Source', width: 16 },
    { header: 'UI', key: 'UI', width: 18 },
    { header: 'Title', key: 'Title', width: 80 },
    { header: 'DOI', key: 'DOI', width: 45 },
    { header: 'Status', key: 'Status', width: 22 },
    { header: 'Duplicate', key: 'Duplicate', width: 12 },
    { header: 'Duplicate Reason', key: 'DupReason', width: 30 },
    { header: 'Batch', key: 'Batch', width: 20 },
  ];
  styleHeader(ws);

  data.allRecords.forEach(row => {
    const r = ws.addRow({
      Source: row._rowType === 'master' ? 'Master' : 'New',
      UI: row.UI, Title: row.Title, DOI: row.DOI || '',
      Status: row.Status || '',
      Duplicate: row.Duplicate ? 'Yes' : 'No',
      DupReason: row['Duplicate Reason'] || '',
      Batch: row.Batch || '',
    });
    const fill = fillForRow(row);
    if (fill) r.eachCell(cell => { cell.fill = fill; });
  });

  ws.views = [{ state: 'frozen', ySplit: 1 }];
  ws.autoFilter = { from: 'A1', to: `H${ws.rowCount}` };
  saveWorkbook(wb, 'Screening_Output.xlsx');
}

async function generateLogExcel(data) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('New Entries Log');
  ws.columns = [
    { header: 'UI', key: 'UI', width: 18 },
    { header: 'Title', key: 'Title', width: 80 },
    { header: 'DOI', key: 'DOI', width: 45 },
    { header: 'Status', key: 'Status', width: 22 },
    { header: 'OA Source', key: 'OASource', width: 30 },
    { header: 'Publisher', key: 'Publisher', width: 30 },
    { header: 'Publisher URL', key: 'PubURL', width: 50 },
    { header: 'Purchase Cost', key: 'Cost', width: 25 },
    { header: 'Notes', key: 'Notes', width: 50 },
  ];
  styleHeader(ws);

  data.newEntries.forEach(row => {
    const r = ws.addRow({
      UI: row.UI, Title: row.Title, DOI: row.DOI || '',
      Status: row.Status || '', OASource: row['OA Source'] || '',
      Publisher: row.Publisher || '', PubURL: row['Publisher URL'] || '',
      Cost: row['Purchase Cost'] || '', Notes: row.Notes || '',
    });
    const fill = fillForRow(row);
    if (fill) r.eachCell(cell => { cell.fill = fill; });
  });

  ws.views = [{ state: 'frozen', ySplit: 1 }];
  saveWorkbook(wb, 'New_Entries_Log.xlsx');
}

async function generatePurchaseExcel(data) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Purchase List');
  ws.columns = [
    { header: 'UI', key: 'UI', width: 18 },
    { header: 'Title', key: 'Title', width: 80 },
    { header: 'DOI', key: 'DOI', width: 45 },
    { header: 'Publisher', key: 'Publisher', width: 30 },
    { header: 'Website/site link for purchase', key: 'Link', width: 70 },
    { header: 'Purchase Cost', key: 'Cost', width: 30 },
  ];
  styleHeader(ws);

  (data.purchase || []).forEach(row => {
    ws.addRow({
      UI: row.UI, Title: row.Title, DOI: row.DOI || '',
      Publisher: row.Publisher || '',
      Link: row['Website/site link for purchase'] || '',
      Cost: row['Purchase Cost'] || '',
    });
  });

  ws.views = [{ state: 'frozen', ySplit: 1 }];
  saveWorkbook(wb, 'Purchase_List.xlsx');
}

function styleHeader(ws) {
  ws.getRow(1).eachCell(cell => {
    cell.fill = HEADER_FILL;
    cell.font = HEADER_FONT;
  });
}

async function saveWorkbook(wb, filename) {
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/* ===== UTILITIES ===== */
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function esc(str) {
  const d = document.createElement('div');
  d.textContent = str || '';
  return d.innerHTML;
}

function doiLink(doi) {
  if (!doi) return '—';
  const clean = esc(doi);
  return `<a href="https://doi.org/${clean}" target="_blank" rel="noopener"
            style="color:var(--accent);font-family:var(--font-mono);font-size:0.8rem">${clean}</a>`;
}

function badgeClass(status) {
  const map = {
    'DOWNLOADED':        'badge-downloaded',
    'OA_AVAILABLE':      'badge-downloaded',
    'PURCHASE_REQUIRED': 'badge-purchase',
    'MANUAL_REVIEW':     'badge-review',
    'OA_DOWNLOAD_FAILED':'badge-oa-failed',
    'ERROR':             'badge-error',
    'ALREADY_IN_MASTER': 'badge-in-master',
    'BATCH_DUPLICATE':   'badge-batch-dup',
  };
  return map[status] || 'badge-review';
}

function formatStatus(s) {
  return (s || '').replace(/_/g, ' ')
    .replace(/\b\w/g, c => c.toUpperCase())
    .replace('Oa ', 'OA ');
}

function showToast(msg) {
  document.querySelectorAll('.toast').forEach(t => t.remove());
  const t = document.createElement('div');
  t.className = 'toast';
  t.textContent = '❌ ' + msg;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), 6000);
}
