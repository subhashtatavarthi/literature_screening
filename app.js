/**
 * Literature Screening Pipeline — Client-Side Engine
 *
 * Runs entirely in the browser:
 *   - Reads Excel/CSV with SheetJS
 *   - Detects duplicates locally
 *   - Calls Crossref & Unpaywall APIs (CORS-enabled)
 *   - Generates styled Excel downloads with ExcelJS
 */

/* ===== CONSTANTS ===== */
const CROSSREF_API = 'https://api.crossref.org/works/';
const UNPAYWALL_API = 'https://api.unpaywall.org/v2/';
const REQUEST_DELAY = 600; // ms between API calls

/* ===== STATE ===== */
let selectedFile = null;
let pipelineData = null; // { screening, log, purchase }

/* ===== INIT ===== */
document.addEventListener('DOMContentLoaded', () => {
  setupDropzone();
  setupTabs();
  setupDownloads();
  document.getElementById('email-input').addEventListener('input', updateRunBtn);
  document.getElementById('run-btn').addEventListener('click', handleRun);
});

/* ===== DROPZONE ===== */
function setupDropzone() {
  const dz = document.getElementById('dropzone');
  const fi = document.getElementById('file-input');

  dz.addEventListener('click', () => fi.click());
  dz.addEventListener('keydown', e => { if (e.key === 'Enter') fi.click(); });
  dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
  dz.addEventListener('drop', e => {
    e.preventDefault();
    dz.classList.remove('drag-over');
    if (e.dataTransfer.files.length) selectFile(e.dataTransfer.files[0]);
  });
  fi.addEventListener('change', () => { if (fi.files.length) selectFile(fi.files[0]); });
}

function selectFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx', 'xls', 'csv'].includes(ext)) {
    showToast('Please select an .xlsx, .xls, or .csv file.');
    return;
  }
  selectedFile = file;
  const el = document.getElementById('file-name');
  el.textContent = '✅ ' + file.name;
  el.style.display = 'block';
  updateRunBtn();
}

function updateRunBtn() {
  document.getElementById('run-btn').disabled =
    !(selectedFile && document.getElementById('email-input').value.trim());
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

/* ===== DOWNLOAD BUTTONS ===== */
function setupDownloads() {
  document.getElementById('dl-screening').addEventListener('click', e => {
    e.preventDefault();
    if (pipelineData) generateScreeningExcel(pipelineData.screening);
  });
  document.getElementById('dl-log').addEventListener('click', e => {
    e.preventDefault();
    if (pipelineData) generateLogExcel(pipelineData.log);
  });
  document.getElementById('dl-purchase').addEventListener('click', e => {
    e.preventDefault();
    if (pipelineData) generatePurchaseExcel(pipelineData.purchase);
  });
}

/* ===== MAIN RUN HANDLER ===== */
async function handleRun() {
  if (!selectedFile) return;
  const email = document.getElementById('email-input').value.trim();
  if (!email) return;

  const overlay = document.getElementById('processing-overlay');
  const status = document.getElementById('processing-status');
  const bar = document.getElementById('progress-bar');

  overlay.classList.add('active');
  status.textContent = 'Reading file…';
  bar.style.width = '0%';
  document.getElementById('results-section').classList.remove('visible');

  try {
    /* 1. Parse the uploaded file */
    const records = await parseFile(selectedFile);
    if (!records.length) throw new Error('No records found in the file.');
    status.textContent = `Parsed ${records.length} records. Detecting duplicates…`;
    bar.style.width = '5%';

    /* 2. Detect duplicates */
    const screening = detectDuplicates(records);
    status.textContent = `Found duplicates. Running API lookups…`;
    bar.style.width = '10%';

    /* 3. Process each record (API lookups) */
    const log = [];
    for (let i = 0; i < screening.length; i++) {
      const row = screening[i];
      status.textContent = `[${i + 1}/${screening.length}] Processing ${row.UI}…`;
      bar.style.width = `${10 + (i / screening.length) * 80}%`;

      const result = await processRecord(row, email);
      log.push(result);

      if (i < screening.length - 1) await sleep(REQUEST_DELAY);
    }

    /* 4. Build purchase list */
    const purchase = log
      .filter(r => r.Status === 'PURCHASE_REQUIRED')
      .map(r => ({
        UI: r.UI, Title: r.Title, DOI: r.DOI,
        Publisher: r.Publisher,
        'Website/site link for purchase': r['Publisher URL'],
        'Purchase Cost': r['Purchase Cost'],
      }));

    bar.style.width = '95%';
    status.textContent = 'Rendering results…';

    pipelineData = { screening, log, purchase };
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
function parseFile(file) {
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
        // Normalize column names
        const records = json.map(row => {
          const obj = {};
          for (const [key, val] of Object.entries(row)) {
            obj[key.trim()] = String(val).trim();
          }
          return obj;
        });
        // Validate required columns
        if (records.length > 0) {
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

/* ===== DUPLICATE DETECTION ===== */
function normalizeDoi(val) {
  let doi = (val || '').trim().toLowerCase();
  if (!doi) return '';
  for (const pfx of [
    'https://doi.org/', 'http://doi.org/',
    'https://dx.doi.org/', 'http://dx.doi.org/',
  ]) {
    doi = doi.replace(pfx, '');
  }
  doi = doi.replace(/^\s*doi\s*:\s*/, '');
  return doi.trim();
}

function normalizeTitle(val) {
  let t = (val || '').trim().toLowerCase();
  t = t.replace(/[^\w\s]/g, ' ');
  return t.replace(/\s+/g, ' ').trim();
}

function detectDuplicates(records) {
  const withNorm = records.map(r => ({
    ...r,
    'Normalized DOI': normalizeDoi(r.DOI),
    'Normalized Title': normalizeTitle(r.Title),
  }));

  // Count occurrences
  const doiCount = {};
  const titleCount = {};
  withNorm.forEach(r => {
    if (r['Normalized DOI']) doiCount[r['Normalized DOI']] = (doiCount[r['Normalized DOI']] || 0) + 1;
    if (r['Normalized Title']) titleCount[r['Normalized Title']] = (titleCount[r['Normalized Title']] || 0) + 1;
  });

  return withNorm.map(r => {
    const doiDup = r['Normalized DOI'] && (doiCount[r['Normalized DOI']] || 0) > 1;
    const titleDup = r['Normalized Title'] && (titleCount[r['Normalized Title']] || 0) > 1;
    const reasons = [];
    if (doiDup) reasons.push('Duplicate DOI');
    if (titleDup) reasons.push('Duplicate Title');
    return {
      ...r,
      Duplicate: doiDup || titleDup,
      'Duplicate Reason': reasons.join('; '),
    };
  });
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
    UI: row.UI, Title: row.Title, DOI: doi,
    Duplicate: !!row.Duplicate,
    Status: '', 'OA Source': '', 'PDF URL': '',
    Publisher: '', 'Publisher URL': '', 'Purchase Cost': '', Notes: '',
  };

  if (!doi) {
    result.Status = 'MANUAL_REVIEW';
    result.Notes = 'DOI missing';
    return result;
  }

  // Crossref
  const crossref = await crossrefLookup(doi);
  if (crossref) {
    result.Publisher = crossref.publisher || '';
    result['Publisher URL'] = crossref.URL || `https://doi.org/${doi}`;
  } else {
    result['Publisher URL'] = `https://doi.org/${doi}`;
  }

  // Unpaywall
  const oa = await unpaywallLookup(doi, email);
  if (!oa) {
    result.Status = 'MANUAL_REVIEW';
    result.Notes = 'Unable to determine OA status';
    return result;
  }

  const candidates = extractOaCandidates(oa);
  if (candidates.length > 0) {
    // In browser we can't download PDFs due to CORS, but we provide the link
    result.Status = 'OA_AVAILABLE';
    result['OA Source'] = candidates[0].source;
    result['PDF URL'] = candidates[0].url;
    result.Notes = 'Open-access PDF link available (click DOI or see PDF URL).';
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
  renderStats(data);
  renderScreeningTable(data.screening);
  renderLogTable(data.log);
  renderPurchaseTable(data.purchase);
}

function renderStats(data) {
  const bar = document.getElementById('stats-bar');
  bar.innerHTML = '';
  const counts = {};
  data.log.forEach(r => { counts[r.Status] = (counts[r.Status] || 0) + 1; });
  const totalDups = data.screening.filter(r => r.Duplicate).length;

  addStat(bar, data.screening.length, 'Total');
  addStat(bar, totalDups, 'Duplicates');
  Object.entries(counts).forEach(([s, c]) => addStat(bar, c, formatStatus(s)));
}

function addStat(container, count, label) {
  const d = document.createElement('div');
  d.className = 'stat-item';
  d.innerHTML = `<span class="stat-count">${count}</span><span class="stat-label">${label}</span>`;
  container.appendChild(d);
}

function renderScreeningTable(rows) {
  const body = document.getElementById('screening-body');
  body.innerHTML = '';
  rows.forEach(row => {
    const tr = document.createElement('tr');
    if (row.Duplicate) tr.classList.add('duplicate-row');
    tr.innerHTML = `
      <td>${esc(row.UI)}</td>
      <td>${esc(row.Title)}</td>
      <td>${doiLink(row.DOI)}</td>
      <td>${row.Duplicate ? '<span class="badge badge-duplicate">Duplicate</span>' : '—'}</td>
      <td>${esc(row['Duplicate Reason'] || '')}</td>`;
    body.appendChild(tr);
  });
}

function renderLogTable(rows) {
  const body = document.getElementById('log-body');
  body.innerHTML = '';
  rows.forEach(row => {
    const tr = document.createElement('tr');
    if (row.Duplicate) tr.classList.add('duplicate-row');
    tr.innerHTML = `
      <td>${esc(row.UI)}</td>
      <td>${esc(row.Title)}</td>
      <td>${doiLink(row.DOI)}</td>
      <td>${row.Duplicate ? '<span class="badge badge-duplicate">Duplicate</span>' : '—'}</td>
      <td><span class="badge ${badgeClass(row.Status)}">${formatStatus(row.Status)}</span></td>
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

/* ===== EXCEL EXPORT (ExcelJS) ===== */
async function generateScreeningExcel(rows) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Titles for Screening');
  const cols = ['UI', 'Title', 'DOI', 'Duplicate', 'Duplicate Reason'];
  ws.columns = [
    { header: 'UI', key: 'UI', width: 18 },
    { header: 'Title', key: 'Title', width: 80 },
    { header: 'DOI', key: 'DOI', width: 45 },
    { header: 'Duplicate', key: 'Duplicate', width: 15 },
    { header: 'Duplicate Reason', key: 'DuplicateReason', width: 30 },
  ];

  // Style header
  const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2D3142' } };
  const headerFont = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
  ws.getRow(1).eachCell(cell => {
    cell.fill = headerFill;
    cell.font = headerFont;
  });

  const dupFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
  rows.forEach(row => {
    const r = ws.addRow({
      UI: row.UI, Title: row.Title, DOI: row.DOI || row['Normalized DOI'] || '',
      Duplicate: row.Duplicate ? 'Yes' : 'No',
      DuplicateReason: row['Duplicate Reason'] || '',
    });
    if (row.Duplicate) {
      r.eachCell(cell => { cell.fill = dupFill; });
    }
  });

  ws.views = [{ state: 'frozen', ySplit: 1 }];
  ws.autoFilter = { from: 'A1', to: `E${rows.length + 1}` };
  saveWorkbook(wb, 'Screening_Output.xlsx');
}

async function generateLogExcel(rows) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Download Log');
  ws.columns = [
    { header: 'UI', key: 'UI', width: 18 },
    { header: 'Title', key: 'Title', width: 80 },
    { header: 'DOI', key: 'DOI', width: 45 },
    { header: 'Duplicate', key: 'Duplicate', width: 12 },
    { header: 'Status', key: 'Status', width: 22 },
    { header: 'OA Source', key: 'OASource', width: 30 },
    { header: 'PDF URL', key: 'PDFURL', width: 50 },
    { header: 'Publisher', key: 'Publisher', width: 30 },
    { header: 'Publisher URL', key: 'PublisherURL', width: 50 },
    { header: 'Purchase Cost', key: 'PurchaseCost', width: 25 },
    { header: 'Notes', key: 'Notes', width: 50 },
  ];

  const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2D3142' } };
  const headerFont = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
  ws.getRow(1).eachCell(cell => { cell.fill = headerFill; cell.font = headerFont; });

  const dupFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
  rows.forEach(row => {
    const r = ws.addRow({
      UI: row.UI, Title: row.Title, DOI: row.DOI,
      Duplicate: row.Duplicate ? 'Yes' : 'No',
      Status: row.Status, OASource: row['OA Source'],
      PDFURL: row['PDF URL'], Publisher: row.Publisher,
      PublisherURL: row['Publisher URL'],
      PurchaseCost: row['Purchase Cost'], Notes: row.Notes,
    });
    if (row.Duplicate) {
      r.eachCell(cell => { cell.fill = dupFill; });
    }
  });

  ws.views = [{ state: 'frozen', ySplit: 1 }];
  saveWorkbook(wb, 'Download_Log.xlsx');
}

async function generatePurchaseExcel(rows) {
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

  const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2D3142' } };
  const headerFont = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
  ws.getRow(1).eachCell(cell => { cell.fill = headerFill; cell.font = headerFont; });

  rows.forEach(row => {
    ws.addRow({
      UI: row.UI, Title: row.Title, DOI: row.DOI,
      Publisher: row.Publisher,
      Link: row['Website/site link for purchase'],
      Cost: row['Purchase Cost'],
    });
  });

  ws.views = [{ state: 'frozen', ySplit: 1 }];
  saveWorkbook(wb, 'Purchase_List.xlsx');
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
