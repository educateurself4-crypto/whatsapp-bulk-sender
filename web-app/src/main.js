/**
 * WhatsApp Bulk Sender — Web App
 * Backend credentials come from .env (VITE_N8N_BASE_URL, VITE_N8N_API_KEY).
 * The user never sees n8n URLs or API keys.
 */

import './index.css';
import * as XLSX from 'xlsx';

// ============================================================
// Backend config — injected from .env at build time
// ============================================================

const N8N_BASE_URL = (import.meta.env.VITE_N8N_BASE_URL || '').replace(/\/+$/, '');
const N8N_API_KEY  = import.meta.env.VITE_N8N_API_KEY  || '';
const DEFAULT_CC   = (import.meta.env.VITE_DEFAULT_CC  || '91').replace(/^\+/, '');

// ============================================================
// Helpers
// ============================================================

const $ = (sel) => document.querySelector(sel);
const CHUNK_SIZE = 50;

function show(el) { el.classList.remove('hidden'); }
function hide(el) { el.classList.add('hidden'); }
function toggleVis(el, visible) { visible ? show(el) : hide(el); }

// ============================================================
// Toast Notifications
// ============================================================

function toast(msg, type = 'info', duration = 4000) {
  const container = $('#toastContainer');
  const el = document.createElement('div');
  el.className = `toast toast--${type}`;
  el.textContent = msg;
  container.appendChild(el);
  setTimeout(() => {
    el.classList.add('toast--exit');
    el.addEventListener('animationend', () => el.remove());
  }, duration);
}

// ============================================================
// Log Console
// ============================================================

const logConsole = $('#logConsole');

function logLine(msg, cls = '') {
  const ts = new Date().toLocaleTimeString('en-GB', { hour12: false });
  const div = document.createElement('div');
  div.className = `log-line ${cls}`;
  div.innerHTML = `<span class="log-ts">[${ts}]</span> ${escapeHtml(msg)}`;
  logConsole.appendChild(div);
  logConsole.scrollTop = logConsole.scrollHeight;
}

function escapeHtml(s) {
  const d = document.createElement('div');
  d.textContent = s;
  return d.innerHTML;
}

// ============================================================
// Backend Client — reads credentials from .env, not the UI
// ============================================================

class BackendClient {
  constructor() {
    this.base    = N8N_BASE_URL;
    this.headers = { 'x-api-key': N8N_API_KEY, 'Content-Type': 'application/json' };
  }

  async _post(path, payload) {
    if (!this.base) throw new Error('Backend URL is not configured. Edit .env and restart the dev server.');
    const url = `${this.base}/webhook/${path}`;
    let res;
    try {
      res = await fetch(url, {
        method: 'POST',
        headers: this.headers,
        body: JSON.stringify(payload),
      });
    } catch (e) {
      throw new Error(`Cannot reach the backend server: ${e.message}`);
    }
    if (res.status === 401 || res.status === 403) {
      throw new Error('The server rejected the API key. Check VITE_N8N_API_KEY in .env.');
    }
    if (res.status === 404) {
      throw new Error(`Endpoint '${path}' not found. Make sure the backend workflow is active.`);
    }
    if (res.status >= 400) {
      const txt = await res.text();
      throw new Error(`Server error ${res.status}: ${txt.slice(0, 300)}`);
    }
    const text = await res.text();
    if (!text) return {};
    try { return JSON.parse(text); } catch { return {}; }
  }

  async listTemplates() {
    const d = await this._post('wa-list-templates', {});
    return d.templates || [];
  }

  async sendChunk(campaignId, template, contacts) {
    await this._post('wa-send-campaign', {
      campaignId,
      templateName: template.name,
      language: template.language,
      contacts,
    });
  }

  async campaignStatus(campaignId) {
    const d = await this._post('wa-campaign-status', { campaignId });
    for (const k of ['processed', 'accepted', 'send_failed', 'delivered', 'read', 'delivery_failed']) {
      d[k] = parseInt(d[k] || 0, 10) || 0;
    }
    d.failures = d.failures || [];
    return d;
  }

  async optouts() {
    const d = await this._post('wa-optouts', {});
    return new Set(d.phones || []);
  }

  async importContacts(contacts, sheetTab = 'Contacts') {
    const d = await this._post('wa-import-contacts', { contacts, sheetTab });
    return parseInt(d.imported || contacts.length, 10);
  }
}

// Single shared client instance — created once from env vars
const client = new BackendClient();

function getCC() {
  return ($('#inputCC').value || '').trim().replace(/^\+/, '') || DEFAULT_CC;
}

// ============================================================
// Phone Normalisation  (port of core.py normalize_phone)
// ============================================================

function normalizePhone(raw, defaultCC = '91') {
  const s = (raw || '').trim();
  if (!s) return [null, 'empty phone number'];

  const international = s.startsWith('+') || s.startsWith('00');
  let digits = s.replace(/\D/g, '');
  if (s.startsWith('00')) digits = digits.slice(2);
  if (!digits) return [null, 'no digits in phone number'];

  if (!international) {
    digits = digits.replace(/^0+/, '');
    if (digits.length === 10) digits = defaultCC + digits;
  }

  if (digits.length < 8 || digits.length > 15) {
    return [null, `invalid length (${digits.length} digits)`];
  }

  if (digits.startsWith('91')) {
    const national = digits.slice(2);
    if (national.length !== 10 || !'6789'.includes(national[0])) {
      return [null, 'invalid Indian mobile number'];
    }
  }
  return [digits, ''];
}

// ============================================================
// Template Parsing  (port of core.py parse_template)
// ============================================================

function parseTemplate(t) {
  let body = '';
  let supported = true;
  let reason = '';

  for (const comp of (t.components || [])) {
    const ctype = (comp.type || '').toUpperCase();
    if (ctype === 'BODY') {
      body = comp.text || '';
    } else if (ctype === 'FOOTER') {
      // ok
    } else if (ctype === 'HEADER') {
      const fmt = (comp.format || '').toUpperCase();
      const re = /\{\{\s*(\w+)\s*\}\}/g;
      if (fmt !== 'TEXT' || re.test(comp.text || '')) {
        supported = false;
        reason = `header (${fmt || 'variable'}) needs extra parameters`;
      }
    } else if (ctype === 'BUTTONS') {
      for (const b of (comp.buttons || [])) {
        const btype = (b.type || '').toUpperCase();
        if (btype === 'QUICK_REPLY' || btype === 'PHONE_NUMBER') continue;
        if (btype === 'URL' && !(b.url || '').includes('{{')) continue;
        supported = false;
        reason = `button type ${btype} needs extra parameters`;
      }
    } else {
      supported = false;
      reason = `component ${ctype} is not supported`;
    }
  }

  const names = [];
  let m;
  const re = /\{\{\s*(\w+)\s*\}\}/g;
  while ((m = re.exec(body)) !== null) names.push(m[1]);

  let varCount = 0;
  if (names.length) {
    if (names.every(n => /^\d+$/.test(n))) {
      varCount = Math.max(...names.map(Number));
    } else {
      supported = false;
      reason = 'named variables are not supported';
    }
  }

  const flag = supported ? '' : '  [not supported by this app]';
  return {
    name: t.name || '',
    language: t.language || '',
    category: t.category || '',
    body,
    varCount,
    supported,
    reason,
    label: `${t.name || ''}  (${t.language || ''}, ${t.category || ''})${flag}`,
  };
}

// ============================================================
// Contact Validation  (port of core.py build_contacts)
// ============================================================

function cleanParam(value) {
  let v = (value || '').replace(/[\r\n\t]+/g, ' ');
  v = v.replace(/ {2,}/g, ' ');
  return v.trim();
}

function buildContacts(rows, phoneIdx, varIdx, defaultCC, optedOut) {
  const rep = { valid: [], invalid: [], duplicates: 0, optedOut: 0 };
  const seen = new Set();

  rows.forEach((row, i) => {
    const rowNum = i + 2;
    const rawPhone = row[phoneIdx] || '';
    const [phone, err] = normalizePhone(rawPhone, defaultCC);
    if (err) {
      rep.invalid.push([rowNum, rawPhone, err]);
      return;
    }

    const params = varIdx.map(idx => cleanParam(row[idx] || ''));
    const empty = params.reduce((acc, p, k) => p === '' ? [...acc, k + 1] : acc, []);
    if (empty.length) {
      rep.invalid.push([rowNum, rawPhone, `empty value for {{${empty.join('}}, {{')}}}`]);
      return;
    }

    if (optedOut.has(phone)) { rep.optedOut++; return; }
    if (seen.has(phone)) { rep.duplicates++; return; }
    seen.add(phone);
    rep.valid.push({ phone, params });
  });

  return rep;
}

// ============================================================
// CSV Export Helpers
// ============================================================

function downloadCSV(filename, headerRow, dataRows) {
  const lines = [headerRow, ...dataRows].map(r =>
    r.map(c => `"${String(c).replace(/"/g, '""')}"`).join(',')
  );
  const blob = new Blob(['\uFEFF' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

// ============================================================
// Campaign ID
// ============================================================

function makeCampaignId() {
  const now = new Date();
  const pad = (n, w = 2) => String(n).padStart(w, '0');
  return `c_${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
}

// ============================================================
// App State
// ============================================================

const state = {
  templates: [],
  headers: [],
  rows: [],
  report: null,
  lastCampaignId: '',
  lastStatus: null,
  cancelled: false,
  running: false,
};

// ============================================================
// DOM References
// ============================================================

const dom = {
  templateLoading:   $('#templateLoading'),
  templateSelectWrap:$('#templateSelectWrap'),
  btnRefreshTemplates:$('#btnRefreshTemplates'),
  selectTemplate:    $('#selectTemplate'),
  templatePreview:   $('#templatePreview'),
  fileDrop:          $('#fileDrop'),
  fileInput:         $('#fileInput'),
  fileInfo:          $('#fileInfo'),
  fileInfoText:      $('#fileInfoText'),
  phonePickerWrap:   $('#phonePickerWrap'),
  selectPhone:       $('#selectPhone'),
  inputCC:           $('#inputCC'),
  mappingRow:        $('#mappingRow'),
  dataPreviewWrap:   $('#dataPreviewWrap'),
  dataPreviewHead:   $('#dataPreviewHead'),
  dataPreviewBody:   $('#dataPreviewBody'),
  btnValidate:       $('#btnValidate'),
  btnImportSheets:   $('#btnImportSheets'),
  btnExportInvalid:  $('#btnExportInvalid'),
  validationSummary: $('#validationSummary'),
  checkConsent:      $('#checkConsent'),
  inputMax:          $('#inputMax'),
  inputTestPhone:    $('#inputTestPhone'),
  btnTest:           $('#btnTest'),
  btnStart:          $('#btnStart'),
  btnStop:           $('#btnStop'),
  btnRefresh:        $('#btnRefresh'),
  btnExportFails:    $('#btnExportFails'),
  progressWrap:      $('#progressWrap'),
  progressFill:      $('#progressFill'),
  progressText:      $('#progressText'),
  progressPct:       $('#progressPct'),
  statsBar:          $('#statsBar'),
};

// Pre-fill country code from env
dom.inputCC.value = DEFAULT_CC;

// ============================================================
// Step 1: Load Templates (auto on page load + refresh button)
// ============================================================

async function loadTemplates() {
  show(dom.templateLoading);
  hide(dom.templateSelectWrap);
  dom.btnRefreshTemplates.disabled = true;

  try {
    const raw = await client.listTemplates();
    state.templates = raw
      .filter(t => (t.status || '').toUpperCase() === 'APPROVED')
      .map(parseTemplate)
      .sort((a, b) => {
        if (a.supported !== b.supported) return a.supported ? -1 : 1;
        return a.name.localeCompare(b.name);
      });

    dom.selectTemplate.innerHTML = '<option value="">— Select a template —</option>';
    state.templates.forEach((t, i) => {
      const opt = document.createElement('option');
      opt.value = i;
      opt.textContent = t.label;
      dom.selectTemplate.appendChild(opt);
    });
    dom.selectTemplate.disabled = false;

    logLine(`Loaded ${state.templates.length} approved template(s).`, 'log-line--success');
    if (!state.templates.length) toast('No APPROVED templates found on the account.', 'info');
  } catch (e) {
    logLine('ERROR loading templates: ' + e.message, 'log-line--error');
    toast(e.message, 'error');
  } finally {
    hide(dom.templateLoading);
    show(dom.templateSelectWrap);
    dom.btnRefreshTemplates.disabled = false;
  }
}

dom.btnRefreshTemplates.addEventListener('click', loadTemplates);

// ============================================================
// Step 1: Select Template
// ============================================================

function selectedTemplate() {
  const i = parseInt(dom.selectTemplate.value, 10);
  return (i >= 0 && i < state.templates.length) ? state.templates[i] : null;
}

dom.selectTemplate.addEventListener('change', () => {
  const t = selectedTemplate();
  if (!t) { dom.templatePreview.innerHTML = ''; rebuildMapping(); invalidate(); return; }

  let html = escapeHtml(t.body || '(no body)');
  if (!t.supported) {
    html += `<span class="unsupported">⚠ NOT SUPPORTED: ${escapeHtml(t.reason)}</span>`;
  }
  dom.templatePreview.innerHTML = html;
  rebuildMapping();
  invalidate();
});

// ============================================================
// Step 2: File Reading
// ============================================================

dom.fileDrop.addEventListener('dragover', (e) => { e.preventDefault(); dom.fileDrop.classList.add('dragover'); });
dom.fileDrop.addEventListener('dragleave', () => dom.fileDrop.classList.remove('dragover'));
dom.fileDrop.addEventListener('drop', (e) => {
  e.preventDefault();
  dom.fileDrop.classList.remove('dragover');
  if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});
dom.fileInput.addEventListener('change', (e) => {
  if (e.target.files.length) handleFile(e.target.files[0]);
});

function cellToText(v) {
  if (v == null) return '';
  if (typeof v === 'boolean') return String(v);
  if (typeof v === 'number') {
    if (Number.isInteger(v)) return String(v);
    return v.toFixed(6).replace(/0+$/, '').replace(/\.$/, '');
  }
  if (v instanceof Date) {
    const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return `${String(v.getDate()).padStart(2,'0')} ${months[v.getMonth()]} ${v.getFullYear()}`;
  }
  return String(v).trim();
}

function dedupeHeaders(raw) {
  const seen = {};
  return raw.map((h, i) => {
    let name = (h || '').trim() || `Column ${i + 1}`;
    if (seen[name]) { seen[name]++; name = `${name} (${seen[name]})`; }
    else seen[name] = 1;
    return name;
  });
}

function handleFile(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith('.csv')) {
    readCSVFile(file);
  } else if (name.endsWith('.xlsx') || name.endsWith('.xlsm')) {
    readExcelFile(file);
  } else {
    toast('Unsupported file type. Use .xlsx, .xlsm or .csv', 'error');
  }
}

function readCSVFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    const clean = e.target.result.replace(/^\uFEFF/, '');
    const lines = [];
    let current = [], inQuote = false, field = '';
    for (let ci = 0; ci < clean.length; ci++) {
      const ch = clean[ci];
      if (inQuote) {
        if (ch === '"') {
          if (ci + 1 < clean.length && clean[ci + 1] === '"') { field += '"'; ci++; }
          else inQuote = false;
        } else field += ch;
      } else {
        if (ch === '"') inQuote = true;
        else if (ch === ',') { current.push(field); field = ''; }
        else if (ch === '\n') { current.push(field); field = ''; lines.push(current); current = []; }
        else if (ch === '\r') { /* skip */ }
        else field += ch;
      }
    }
    if (field || current.length) { current.push(field); lines.push(current); }

    const nonBlank = lines.filter(r => r.some(c => c.trim() !== ''));
    if (nonBlank.length < 2) { toast('The file needs a header row and at least one data row.', 'error'); return; }

    const headers = dedupeHeaders(nonBlank[0].map(c => cellToText(c)));
    const width = headers.length;
    const rows = nonBlank.slice(1).map(r => {
      const padded = [...r.map(c => cellToText(c)), ...Array(width).fill('')];
      return padded.slice(0, width);
    });
    onFileLoaded(file.name, headers, rows);
  };
  reader.readAsText(file, 'utf-8');
}

function readExcelFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: 'array', cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });

      const nonBlank = rawRows.filter(r => r.some(c => cellToText(c) !== ''));
      if (nonBlank.length < 2) { toast('The file needs a header row and at least one data row.', 'error'); return; }

      const headers = dedupeHeaders(nonBlank[0].map(c => cellToText(c)));
      const width = headers.length;
      const rows = nonBlank.slice(1).map(r => {
        const padded = [...r.map(c => cellToText(c)), ...Array(width).fill('')];
        return padded.slice(0, width);
      });
      onFileLoaded(file.name, headers, rows);
    } catch (err) {
      toast('Cannot read Excel file: ' + err.message, 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

function onFileLoaded(filename, headers, rows) {
  state.headers = headers;
  state.rows = rows;

  dom.fileInfoText.textContent = `${filename}  —  ${rows.length} rows`;
  show(dom.fileInfo);

  dom.selectPhone.innerHTML = headers.map((h, i) => `<option value="${i}">${escapeHtml(h)}</option>`).join('');
  const guessIdx = headers.findIndex(h => /phone|mobile|whatsapp|number|contact/i.test(h));
  dom.selectPhone.value = guessIdx >= 0 ? guessIdx : 0;
  show(dom.phonePickerWrap);

  dom.dataPreviewHead.innerHTML = '<tr>' + headers.map(h => `<th>${escapeHtml(h)}</th>`).join('') + '</tr>';
  dom.dataPreviewBody.innerHTML = rows.slice(0, 5).map(r =>
    '<tr>' + r.map(c => `<td title="${escapeHtml(c)}">${escapeHtml(c)}</td>`).join('') + '</tr>'
  ).join('');
  show(dom.dataPreviewWrap);

  // Enable Import to Sheets as soon as a file is loaded
  dom.btnImportSheets.disabled = false;

  rebuildMapping();
  invalidate();
  logLine(`Loaded ${rows.length} rows from ${filename}`);
}

// ============================================================
// Variable Mapping
// ============================================================

function rebuildMapping() {
  dom.mappingRow.innerHTML = '';
  const t = selectedTemplate();
  if (!t || !state.headers.length || t.varCount === 0) { hide(dom.mappingRow); return; }

  const phoneSel = dom.selectPhone.value;
  const others = state.headers.filter((_, i) => String(i) !== phoneSel);

  for (let n = 1; n <= t.varCount; n++) {
    const item = document.createElement('div');
    item.className = 'mapping-item';
    item.innerHTML = `
      <span class="mapping-item__label">{{${n}}}</span>
      <span class="mapping-item__arrow">←</span>
    `;
    const sel = document.createElement('select');
    sel.innerHTML = '<option value="">— column —</option>' +
      state.headers.map((h, i) => `<option value="${i}">${escapeHtml(h)}</option>`).join('');
    if (n - 1 < others.length) {
      const defIdx = state.headers.indexOf(others[n - 1]);
      sel.value = defIdx >= 0 ? defIdx : '';
    }
    sel.addEventListener('change', invalidate);
    item.appendChild(sel);
    dom.mappingRow.appendChild(item);
  }
  show(dom.mappingRow);
}

dom.selectPhone.addEventListener('change', () => { rebuildMapping(); invalidate(); });

function getMappingSelects() {
  return Array.from(dom.mappingRow.querySelectorAll('select'));
}

// ============================================================
// Invalidation
// ============================================================

function invalidate() {
  state.report = null;
  hide(dom.validationSummary);
  hide(dom.btnExportInvalid);
  dom.btnStart.disabled = true;
  dom.btnValidate.disabled = !(selectedTemplate() && state.rows.length);
}

// ============================================================
// Step 2: Validate Contacts
// ============================================================

dom.btnValidate.addEventListener('click', async () => {
  const t = selectedTemplate();
  if (!t) { toast('Choose a template first.', 'error'); return; }
  if (!t.supported) { toast(`This template cannot be used: ${t.reason}`, 'error'); return; }
  if (!state.rows.length) { toast('Choose a contacts file first.', 'error'); return; }

  const mapSels = getMappingSelects();
  if (mapSels.some(s => !s.value && s.value !== '0')) {
    toast('Map every {{n}} variable to a column.', 'error'); return;
  }

  logLine('Checking the opt-out list and validating contacts…');
  dom.btnValidate.disabled = true;

  try {
    const optedOut = await client.optouts();
    const phoneIdx = parseInt(dom.selectPhone.value, 10);
    const varIdx   = mapSels.map(s => parseInt(s.value, 10));
    const rep      = buildContacts(state.rows, phoneIdx, varIdx, getCC(), optedOut);
    state.report   = rep;

    dom.validationSummary.innerHTML = `
      <span class="val-chip val-chip--valid">✔ ${rep.valid.length} valid</span>
      <span class="val-chip val-chip--invalid">✖ ${rep.invalid.length} invalid</span>
      <span class="val-chip val-chip--dup">⊘ ${rep.duplicates} duplicates</span>
      <span class="val-chip val-chip--optout">🚫 ${rep.optedOut} opted out</span>
    `;
    show(dom.validationSummary);
    toggleVis(dom.btnExportInvalid, rep.invalid.length > 0);
    dom.btnStart.disabled = !rep.valid.length;

    logLine(`${rep.valid.length} valid | ${rep.invalid.length} invalid | ${rep.duplicates} duplicates | ${rep.optedOut} opted out`);
    if (rep.valid.length) toast(`${rep.valid.length} contacts ready to send`, 'success');
    else toast('No valid contacts found.', 'error');
  } catch (e) {
    logLine('ERROR: ' + e.message, 'log-line--error');
    toast(e.message, 'error');
  } finally {
    dom.btnValidate.disabled = false;
  }
});

// Export invalid
dom.btnExportInvalid.addEventListener('click', () => {
  if (!state.report || !state.report.invalid.length) return;
  downloadCSV('invalid_rows.csv', ['Excel row', 'Phone as entered', 'Problem'], state.report.invalid);
  logLine('Saved invalid_rows.csv');
  toast('Invalid rows exported', 'success');
});

// ============================================================
// Step 2: Import Contacts to Google Sheets
// ============================================================

dom.btnImportSheets.addEventListener('click', async () => {
  if (!state.rows.length) { toast('Load a contacts file first.', 'error'); return; }

  const headers = state.headers;
  // Build contact objects with all columns as keys
  const contacts = state.rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] || ''; });
    return obj;
  });

  logLine(`Importing ${contacts.length} rows to Google Sheets…`);
  dom.btnImportSheets.disabled = true;

  try {
    const imported = await client.importContacts(contacts, 'Contacts');
    logLine(`✔ ${imported} rows imported to Google Sheets (Contacts tab).`, 'log-line--success');
    toast(`${imported} rows imported to Google Sheets`, 'success');
  } catch (e) {
    logLine('ERROR: ' + e.message, 'log-line--error');
    toast(e.message, 'error');
  } finally {
    dom.btnImportSheets.disabled = state.rows.length === 0;
  }
});

// ============================================================
// Step 3: Test Message
// ============================================================

dom.btnTest.addEventListener('click', async () => {
  const t = selectedTemplate();
  if (!t || !t.supported) { toast('Choose a supported template first.', 'error'); return; }

  const [phone, err] = normalizePhone($('#inputTestPhone').value, getCC());
  if (err) { toast('Test number: ' + err, 'error'); return; }

  if (!state.rows.length) { toast('Load a file first (row 1 of data is used for values).', 'error'); return; }
  const mapSels = getMappingSelects();
  if (mapSels.some(s => !s.value && s.value !== '0')) {
    toast('Map every variable to a column first.', 'error'); return;
  }

  const varIdx = mapSels.map(s => parseInt(s.value, 10));
  const params = varIdx.map(i => cleanParam(state.rows[0][i] || '') || '-');
  const cid    = 'TEST_' + makeCampaignId();

  if (!confirm(`Send '${t.name}' to ${phone} using the values from the first data row?`)) return;

  logLine(`Sending test message to ${phone}…`);
  dom.btnTest.disabled = true;
  try {
    await client.sendChunk(cid, t, [{ phone, params }]);
    logLine(`Test message sent successfully (id ${cid}).`, 'log-line--success');
    toast('Test message sent!', 'success');
  } catch (e) {
    logLine('ERROR: ' + e.message, 'log-line--error');
    toast(e.message, 'error');
  } finally {
    dom.btnTest.disabled = false;
  }
});

// ============================================================
// Step 3: Start Campaign
// ============================================================

dom.btnStart.addEventListener('click', async () => {
  const t   = selectedTemplate();
  const rep = state.report;
  if (!t || !rep || !rep.valid.length) return;

  if (!dom.checkConsent.checked) { toast('Tick the opt-in confirmation first.', 'error'); return; }

  const limit = parseInt(dom.inputMax.value, 10);
  if (!limit || limit < 1) { toast('Max messages must be a positive whole number.', 'error'); return; }

  const contacts = rep.valid.slice(0, limit);
  const skipped  = rep.valid.length - contacts.length;

  let msg = `Send template '${t.name}' to ${contacts.length} contacts?`;
  if (skipped) msg += `\n\n${skipped} valid contacts are over your 'Max messages' limit and will NOT be sent in this run.`;
  msg += '\n\nMarketing messages are charged per delivered message by Meta.';
  if (!confirm(msg)) return;

  const cid = makeCampaignId();
  state.lastCampaignId = cid;
  state.cancelled = false;
  state.running = true;
  setRunningUI(true);

  show(dom.progressWrap);
  dom.progressFill.style.width = '0%';
  dom.progressText.textContent = `0 / ${contacts.length}`;
  dom.progressPct.textContent  = '0%';

  logLine(`Campaign ${cid} started: ${contacts.length} contacts.`);

  try {
    const total = contacts.length;
    let submitted = 0;

    for (let start = 0; start < total; start += CHUNK_SIZE) {
      if (state.cancelled) {
        logLine('Stopped by user. Chunks already submitted will still finish.');
        break;
      }

      const chunk = contacts.slice(start, start + CHUNK_SIZE);
      await client.sendChunk(cid, t, chunk);
      submitted += chunk.length;
      updateProgress(submitted, total, `Chunk submitted (${submitted}/${total}). Waiting for results…`);

      const deadlineMs = (chunk.length * 4 + 240) * 1000;
      const startTime  = Date.now();

      while (true) {
        const st = await client.campaignStatus(cid);
        if (st.processed >= submitted) {
          updateProgress(st.processed, total,
            `Processed ${st.processed}/${total} (accepted ${st.accepted}, failed at send ${st.send_failed})`
          );
          break;
        }
        if (Date.now() - startTime > deadlineMs) {
          throw new Error(
            'Timed out waiting for the backend. ' +
            'Do NOT resend until you have checked the Messages sheet, or contacts will get duplicates.'
          );
        }
        if (state.cancelled) break;
        await sleep(6000);
      }
    }

    logLine(`Finished. ${submitted} contacts were submitted.`, 'log-line--success');
    toast(`Campaign complete: ${submitted} contacts processed`, 'success');
    refreshStatus();
  } catch (e) {
    logLine('ERROR: ' + e.message, 'log-line--error');
    toast(e.message, 'error');
  } finally {
    state.running = false;
    setRunningUI(false);
  }
});

function updateProgress(done, total, msg) {
  const pct = total > 0 ? Math.round((done / total) * 100) : 0;
  dom.progressFill.style.width = pct + '%';
  dom.progressText.textContent = `${done} / ${total}`;
  dom.progressPct.textContent  = pct + '%';
  logLine(msg);
}

function sleep(ms) { return new Promise(resolve => setTimeout(resolve, ms)); }

function setRunningUI(running) {
  dom.btnStart.disabled   = running || !(state.report && state.report.valid.length);
  dom.btnStop.disabled    = !running;
  dom.btnTest.disabled    = running;
  dom.btnRefresh.disabled = !state.lastCampaignId;
  dom.btnExportFails.disabled = !state.lastStatus;
}

// ============================================================
// Step 3: Stop
// ============================================================

dom.btnStop.addEventListener('click', () => {
  state.cancelled = true;
  logLine('Stop requested — no new chunks will be sent.');
  toast('Stopping after current chunk…', 'info');
});

// ============================================================
// Step 3: Refresh Delivery Status
// ============================================================

dom.btnRefresh.addEventListener('click', refreshStatus);

async function refreshStatus() {
  if (!state.lastCampaignId) return;
  dom.btnRefresh.disabled = true;
  try {
    const st = await client.campaignStatus(state.lastCampaignId);
    state.lastStatus = st;

    $('#statAccepted').textContent      = st.accepted;
    $('#statDelivered').textContent     = st.delivered;
    $('#statRead').textContent          = st.read;
    $('#statSendFailed').textContent    = st.send_failed;
    $('#statDeliveryFailed').textContent = st.delivery_failed;

    show(dom.statsBar);
    dom.btnExportFails.disabled = !(st.failures && st.failures.length);
    logLine(`Status: accepted ${st.accepted} | delivered ${st.delivered} | read ${st.read} | send-failed ${st.send_failed} | delivery-failed ${st.delivery_failed}`);
  } catch (e) {
    logLine('ERROR: ' + e.message, 'log-line--error');
  } finally {
    dom.btnRefresh.disabled = false;
  }
}

// ============================================================
// Step 3: Export Failures
// ============================================================

dom.btnExportFails.addEventListener('click', () => {
  const fails = (state.lastStatus || {}).failures || [];
  if (!fails.length) { toast('No failures recorded (press "Refresh delivery status" first).', 'info'); return; }
  const rows = fails.map(f => [f.phone || '', f.stage || '', f.reason || '']);
  downloadCSV(`${state.lastCampaignId}_failures.csv`, ['Phone', 'Stage', 'Reason'], rows);
  logLine(`Saved ${state.lastCampaignId}_failures.csv`);
  toast('Failures exported', 'success');
});

// ============================================================
// Initialise — auto-load templates on page load
// ============================================================

logLine('WhatsApp Bulk Sender ready. Loading templates…');
loadTemplates();
