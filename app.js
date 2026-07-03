'use strict';

/* -------------------------------------------------------------------------
 * Constants & state
 * ---------------------------------------------------------------------- */
const DB_NAME = 'sparktech-db';
const DB_VERSION = 1;
const STORE_NAME = 'reports';

let previewRows = []; // { siteName, inspectedBy, diameter, date, status }

/* -------------------------------------------------------------------------
 * IndexedDB helper (thin promise wrapper, no external dependency)
 * ---------------------------------------------------------------------- */
function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: 'id' });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function dbAddReport(record) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    tx.objectStore(STORE_NAME).add(record);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function dbGetAllReports() {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readonly');
    const req = tx.objectStore(STORE_NAME).getAll();
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = () => reject(req.error);
  });
}

async function dbDeleteReport(id) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    tx.objectStore(STORE_NAME).delete(id);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

/* -------------------------------------------------------------------------
 * Toast
 * ---------------------------------------------------------------------- */
let toastTimer = null;
function showToast(message) {
  const el = document.getElementById('toast');
  el.textContent = message;
  el.classList.remove('hidden');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.add('hidden'), 3000);
}

/* -------------------------------------------------------------------------
 * Theme toggle
 * ---------------------------------------------------------------------- */
function initTheme() {
  const stored = localStorage.getItem('sparktech-theme');
  const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
  const isDark = stored ? stored === 'dark' : prefersDark;
  document.documentElement.classList.toggle('dark', isDark);
  updateThemeIcon(isDark);

  document.getElementById('theme-toggle').addEventListener('click', () => {
    const nowDark = !document.documentElement.classList.contains('dark');
    document.documentElement.classList.toggle('dark', nowDark);
    localStorage.setItem('sparktech-theme', nowDark ? 'dark' : 'light');
    updateThemeIcon(nowDark);
  });
}

function updateThemeIcon(isDark) {
  document.getElementById('theme-icon-light').classList.toggle('hidden', !isDark);
  document.getElementById('theme-icon-dark').classList.toggle('hidden', isDark);
}

/* -------------------------------------------------------------------------
 * Tabs
 * ---------------------------------------------------------------------- */
function initTabs() {
  const buttons = document.querySelectorAll('.tab-btn');
  buttons.forEach((btn) => {
    btn.addEventListener('click', () => {
      buttons.forEach((b) => b.classList.remove('tab-active'));
      btn.classList.add('tab-active');

      const target = btn.dataset.tab;
      document.getElementById('view-new-report').classList.toggle('hidden', target !== 'new-report');
      document.getElementById('view-history').classList.toggle('hidden', target !== 'history');

      if (target === 'history') renderHistory();
    });
  });
}

/* -------------------------------------------------------------------------
 * Manual form
 * ---------------------------------------------------------------------- */
function initManualForm() {
  const form = document.getElementById('manual-form');
  form.querySelector('input[name="date"]').valueAsDate = new Date();

  form.addEventListener('submit', (e) => {
    e.preventDefault();
    const data = new FormData(form);
    previewRows.push({
      siteName: data.get('siteName').trim(),
      inspectedBy: data.get('inspectedBy').trim(),
      diameter: Number(data.get('diameter')),
      date: data.get('date'),
      status: data.get('status'),
    });
    renderPreview();
    form.reset();
    form.querySelector('input[name="date"]').valueAsDate = new Date();
    form.querySelector('input[name="siteName"]').focus();
  });
}

/* -------------------------------------------------------------------------
 * Excel upload — column mapping + parsing
 * ---------------------------------------------------------------------- */
const FIELD_ALIASES = {
  siteName: ['sitename', 'site'],
  inspectedBy: ['inspectedby', 'inspector', 'inspectedbyname'],
  diameter: ['pipediametermm', 'pipediameter', 'diametermm', 'diameter'],
  date: ['inspectiondate', 'date'],
  status: ['status', 'result'],
};

function normalizeKey(key) {
  return String(key).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
}

function excelDateToIso(value) {
  if (value instanceof Date && !isNaN(value)) {
    return value.toISOString().slice(0, 10);
  }
  const parsed = new Date(value);
  if (!isNaN(parsed)) return parsed.toISOString().slice(0, 10);
  return String(value);
}

function normalizeStatus(value) {
  const v = String(value).trim().toLowerCase();
  if (v.startsWith('p')) return 'Pass';
  if (v.startsWith('f')) return 'Fail';
  return String(value).trim() || 'Pass';
}

function mapExcelRow(rawRow) {
  const normalized = {};
  Object.keys(rawRow).forEach((key) => {
    normalized[normalizeKey(key)] = rawRow[key];
  });

  const pick = (field) => {
    for (const alias of FIELD_ALIASES[field]) {
      if (normalized[alias] !== undefined && normalized[alias] !== '') return normalized[alias];
    }
    return undefined;
  };

  const siteName = pick('siteName');
  const inspectedBy = pick('inspectedBy');
  if (siteName === undefined || inspectedBy === undefined) return null; // not a usable row

  return {
    siteName: String(siteName).trim(),
    inspectedBy: String(inspectedBy).trim(),
    diameter: Number(pick('diameter')) || 0,
    date: excelDateToIso(pick('date') ?? new Date()),
    status: normalizeStatus(pick('status') ?? 'Pass'),
  };
}

function handleExcelFile(file) {
  const statusEl = document.getElementById('upload-status');
  const reader = new FileReader();

  reader.onerror = () => {
    console.error('Excel upload failed: FileReader error reading', file.name, reader.error);
    statusEl.textContent = `Could not read "${file.name}".`;
    statusEl.className = 'text-xs mt-2 text-red-600 dark:text-red-400';
    statusEl.classList.remove('hidden');
  };

  reader.onload = (e) => {
    try {
      const workbook = XLSX.read(e.target.result, { type: 'array', cellDates: true });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

      const mapped = rawRows.map(mapExcelRow).filter(Boolean);
      if (mapped.length === 0) {
        statusEl.textContent = 'No matching columns found. Expected: Site Name, Inspected By, Pipe Diameter (mm), Inspection Date, Status.';
        statusEl.className = 'text-xs mt-2 text-red-600 dark:text-red-400';
        statusEl.classList.remove('hidden');
        return;
      }

      previewRows.push(...mapped);
      renderPreview();
      statusEl.textContent = `Imported ${mapped.length} record(s) from "${file.name}".`;
      statusEl.className = 'text-xs mt-2 text-emerald-600 dark:text-emerald-400';
      statusEl.classList.remove('hidden');
    } catch (err) {
      console.error('Excel upload failed: could not parse workbook', file.name, err);
      statusEl.textContent = `Failed to parse "${file.name}". Is it a valid .xlsx file?`;
      statusEl.className = 'text-xs mt-2 text-red-600 dark:text-red-400';
      statusEl.classList.remove('hidden');
    }
  };

  reader.readAsArrayBuffer(file);
}

function initExcelUpload() {
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('file-input');

  fileInput.addEventListener('change', () => {
    if (fileInput.files[0]) handleExcelFile(fileInput.files[0]);
    fileInput.value = '';
  });

  ['dragenter', 'dragover'].forEach((evt) =>
    dropzone.addEventListener(evt, (e) => {
      e.preventDefault();
      dropzone.classList.add('drag-over');
    })
  );
  ['dragleave', 'drop'].forEach((evt) =>
    dropzone.addEventListener(evt, (e) => {
      e.preventDefault();
      dropzone.classList.remove('drag-over');
    })
  );
  dropzone.addEventListener('drop', (e) => {
    const file = e.dataTransfer.files[0];
    if (file) handleExcelFile(file);
  });
}

/* -------------------------------------------------------------------------
 * Preview table
 * ---------------------------------------------------------------------- */
function renderPreview() {
  const body = document.getElementById('preview-body');
  const empty = document.getElementById('preview-empty');
  const countEl = document.getElementById('preview-count');
  const generateBtn = document.getElementById('generate-btn');

  countEl.textContent = previewRows.length;
  generateBtn.disabled = previewRows.length === 0;
  empty.classList.toggle('hidden', previewRows.length > 0);

  body.innerHTML = previewRows
    .map((row, index) => {
      const pillClass = row.status === 'Pass' ? 'status-pass' : 'status-fail';
      return `
        <tr class="border-b border-slate-100 dark:border-slate-700/60">
          <td class="py-2 pr-3">${escapeHtml(row.siteName)}</td>
          <td class="py-2 pr-3">${escapeHtml(row.inspectedBy)}</td>
          <td class="py-2 pr-3">${row.diameter}</td>
          <td class="py-2 pr-3">${escapeHtml(row.date)}</td>
          <td class="py-2 pr-3"><span class="status-pill ${pillClass}">${escapeHtml(row.status)}</span></td>
          <td class="py-2 pr-3">
            <button data-remove-index="${index}" class="text-xs text-slate-400 hover:text-red-500 transition" title="Remove">✕</button>
          </td>
        </tr>`;
    })
    .join('');

  body.querySelectorAll('[data-remove-index]').forEach((btn) => {
    btn.addEventListener('click', () => {
      previewRows.splice(Number(btn.dataset.removeIndex), 1);
      renderPreview();
    });
  });
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/* -------------------------------------------------------------------------
 * PDF generation — vector-drawn, corporate inspection document layout
 * ---------------------------------------------------------------------- */
function buildInspectionPdf(records, reportId, logoDataUrl) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: 'pt', format: 'a4' });

  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = 40;
  let y = 0;

  const drawLetterhead = () => {
    doc.setFillColor(15, 23, 42); // slate-900
    doc.rect(0, 0, pageWidth, 70, 'F');

    let textX = margin;
    if (logoDataUrl) {
      doc.setFillColor(255, 255, 255);
      doc.roundedRect(margin - 4, 10, 50, 50, 6, 6, 'F');
      doc.addImage(logoDataUrl, 'PNG', margin, 14, 42, 42);
      textX = margin + 60;
    }

    doc.setTextColor(255, 255, 255);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(16);
    doc.text('Sparktech Industrial Services', textX, 30);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(10);
    doc.text('Pipeline & Infrastructure Inspection Report', textX, 47);
    y = 95;
  };

  const drawMeta = () => {
    doc.setTextColor(71, 85, 105); // slate-500
    doc.setFontSize(9);
    doc.text(`Report ID: ${reportId}`, margin, y);
    doc.text(`Generated: ${new Date().toLocaleString()}`, pageWidth - margin, y, { align: 'right' });
    y += 8;
    doc.setDrawColor(203, 213, 225);
    doc.line(margin, y, pageWidth - margin, y);
    y += 25;
  };

  const drawFooter = (pageNum, pageCount) => {
    doc.setFontSize(8);
    doc.setTextColor(148, 163, 184);
    doc.text('Generated by Sparktech — client-side inspection report tool', margin, pageHeight - 25);
    doc.text(`Page ${pageNum} of ${pageCount}`, pageWidth - margin, pageHeight - 25, { align: 'right' });
  };

  const ensureSpace = (needed) => {
    if (y + needed > pageHeight - 60) {
      doc.addPage();
      y = 95;
      drawMeta();
    }
  };

  drawLetterhead();
  drawMeta();

  records.forEach((record, index) => {
    ensureSpace(100);

    // Record card border
    const cardTop = y;
    doc.setDrawColor(226, 232, 240);
    doc.setFillColor(248, 250, 252);
    doc.roundedRect(margin, cardTop, pageWidth - margin * 2, 90, 4, 4, 'FD');

    doc.setTextColor(15, 23, 42);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(12);
    doc.text(record.siteName, margin + 14, cardTop + 22);

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9.5);
    doc.setTextColor(71, 85, 105);
    const col1X = margin + 14;
    const col2X = margin + (pageWidth - margin * 2) / 2 + 10;

    doc.text(`Inspected By:  ${record.inspectedBy}`, col1X, cardTop + 44);
    doc.text(`Inspection Date:  ${record.date}`, col2X, cardTop + 44);
    doc.text(`Pipe Diameter:  ${record.diameter} mm`, col1X, cardTop + 62);

    // Status pill
    const isPass = record.status === 'Pass';
    doc.setFillColor(isPass ? 220 : 254, isPass ? 252 : 226, isPass ? 231 : 226);
    doc.roundedRect(col2X, cardTop + 52, 60, 16, 8, 8, 'F');
    doc.setTextColor(isPass ? 22 : 153, isPass ? 101 : 27, isPass ? 52 : 27);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(9);
    doc.text(record.status.toUpperCase(), col2X + 30, cardTop + 63, { align: 'center' });

    y = cardTop + 90 + 16;

    if (index < records.length - 1) {
      ensureSpace(20);
    }
  });

  ensureSpace(80);
  y += 20;
  doc.setDrawColor(203, 213, 225);
  doc.line(margin, y, margin + 200, y);
  doc.line(pageWidth - margin - 200, y, pageWidth - margin, y);
  doc.setFontSize(8.5);
  doc.setTextColor(100, 116, 139);
  doc.text('Inspector Signature', margin, y + 12);
  doc.text('Reviewed By', pageWidth - margin - 200, y + 12);

  const pageCount = doc.internal.getNumberOfPages();
  for (let i = 1; i <= pageCount; i++) {
    doc.setPage(i);
    drawFooter(i, pageCount);
  }

  return doc;
}

/* -------------------------------------------------------------------------
 * Logo — rasterized once from logo.svg and reused in every generated PDF
 * ---------------------------------------------------------------------- */
let cachedLogoDataUrl = null;
async function getLogoDataUrl() {
  if (cachedLogoDataUrl) return cachedLogoDataUrl;

  const response = await fetch('logo.svg');
  const svgText = await response.text();
  const img = new Image();
  const svgUrl = URL.createObjectURL(new Blob([svgText], { type: 'image/svg+xml' }));

  await new Promise((resolve, reject) => {
    img.onload = resolve;
    img.onerror = reject;
    img.src = svgUrl;
  });

  const canvas = document.createElement('canvas');
  canvas.width = 160;
  canvas.height = 160;
  canvas.getContext('2d').drawImage(img, 0, 0, 160, 160);
  URL.revokeObjectURL(svgUrl);

  cachedLogoDataUrl = canvas.toDataURL('image/png');
  return cachedLogoDataUrl;
}

/* -------------------------------------------------------------------------
 * Generate report → save to IndexedDB → download
 * ---------------------------------------------------------------------- */
function initGenerateButton() {
  document.getElementById('generate-btn').addEventListener('click', async () => {
    if (previewRows.length === 0) return;

    const reportId = `INS-${Date.now()}`;
    const logoDataUrl = await getLogoDataUrl().catch((err) => {
      console.error('Failed to load logo.svg for PDF letterhead', err);
      return null;
    });

    let doc;
    try {
      doc = buildInspectionPdf(previewRows, reportId, logoDataUrl);
    } catch (err) {
      console.error('PDF generation failed for report', reportId, err);
      showToast('Failed to generate PDF. See console for details.');
      return;
    }

    const blob = doc.output('blob');
    const fileName = `${reportId}.pdf`;
    const first = previewRows[0];
    const record = {
      id: reportId,
      generatedAt: new Date().toISOString(),
      siteName: previewRows.length > 1 ? `${first.siteName} (+${previewRows.length - 1} more)` : first.siteName,
      inspectedBy: first.inspectedBy,
      recordCount: previewRows.length,
      fileName,
      pdfBlob: blob,
    };

    try {
      await dbAddReport(record);
    } catch (err) {
      console.error('Failed to save report to IndexedDB', reportId, err);
      showToast('Report generated, but could not be saved to history.');
    }

    triggerDownload(blob, fileName);
    showToast('Report generated and saved to history.');

    previewRows = [];
    renderPreview();
    document.getElementById('upload-status').classList.add('hidden');
  });
}

function triggerDownload(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/* -------------------------------------------------------------------------
 * History view
 * ---------------------------------------------------------------------- */
async function renderHistory() {
  const body = document.getElementById('history-body');
  const empty = document.getElementById('history-empty');

  let reports = [];
  try {
    reports = await dbGetAllReports();
  } catch (err) {
    console.error('Failed to load report history from IndexedDB', err);
    showToast('Could not load report history.');
  }

  reports.sort((a, b) => new Date(b.generatedAt) - new Date(a.generatedAt));
  empty.classList.toggle('hidden', reports.length > 0);

  body.innerHTML = reports
    .map(
      (r) => `
        <tr class="border-b border-slate-100 dark:border-slate-700/60">
          <td class="py-2 pr-3">${new Date(r.generatedAt).toLocaleString()}</td>
          <td class="py-2 pr-3">${escapeHtml(r.siteName)}</td>
          <td class="py-2 pr-3">${escapeHtml(r.inspectedBy)}</td>
          <td class="py-2 pr-3">${r.recordCount}</td>
          <td class="py-2 pr-3 flex gap-3">
            <button data-redownload="${r.id}" class="text-xs font-medium text-sky-600 dark:text-sky-400 hover:underline">Redownload PDF</button>
            <button data-delete="${r.id}" class="text-xs text-slate-400 hover:text-red-500 transition" title="Delete">✕</button>
          </td>
        </tr>`
    )
    .join('');

  body.querySelectorAll('[data-redownload]').forEach((btn) => {
    btn.addEventListener('click', async () => {
      const id = btn.dataset.redownload;
      const report = reports.find((r) => r.id === id);
      if (report) triggerDownload(report.pdfBlob, report.fileName);
    });
  });

  body.querySelectorAll('[data-delete]').forEach((btn) => {
    btn.addEventListener('click', async () => {
      const id = btn.dataset.delete;
      if (!confirm('Delete this report from history? This cannot be undone.')) return;
      try {
        await dbDeleteReport(id);
        renderHistory();
      } catch (err) {
        console.error('Failed to delete report', id, err);
        showToast('Could not delete report.');
      }
    });
  });
}

/* -------------------------------------------------------------------------
 * Init
 * ---------------------------------------------------------------------- */
window.addEventListener('DOMContentLoaded', () => {
  initTheme();
  initTabs();
  initManualForm();
  initExcelUpload();
  initGenerateButton();
  renderPreview();
});
