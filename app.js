// ══════════════════════════════════════════════════════════════════
//  Excel PDF Maker — app.js
//  Firebase Realtime DB + SheetJS + jsPDF (B&W / TNR)
// ══════════════════════════════════════════════════════════════════

// ── Firebase Configuration ──────────────────────────────────────
// Replace the values below with your own Firebase project config.
const FIREBASE_CONFIG = {
  apiKey: "AIzaSyBc4oXLBFulQzd6Co7GP6HCj3JuCUdzU5o",
  authDomain: "excel-ea4b9.firebaseapp.com",
  databaseURL: "https://excel-ea4b9-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "excel-ea4b9",
  storageBucket: "excel-ea4b9.firebasestorage.app",
  messagingSenderId: "185343682456",
  appId: "1:185343682456:web:8b17c13e61f179e663aded"
};

// ── State ────────────────────────────────────────────────────────
let db = null;
let useFirebase = false;
let studentData = [];
let formData = {};

const COMPANIES_KEY = 'excelPDF_companies';
const ATTENDEES_KEY = 'excelPDF_attendees';

// ── Init ─────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  initFirebase();
  setupDragDrop();
  setupCompanyAutofill();
});

// ── Firebase Init ────────────────────────────────────────────────
function initFirebase() {
  try {
    if (FIREBASE_CONFIG.apiKey === 'YOUR_API_KEY') throw new Error('not configured');
    firebase.initializeApp(FIREBASE_CONFIG);
    db = firebase.database();
    useFirebase = true;
    document.getElementById('firebaseBadge').textContent = '🔥 Firebase Connected';
  } catch (err) {
    useFirebase = false;
    const badge = document.getElementById('firebaseBadge');
    badge.textContent = '💾 Local Storage Mode';
    badge.style.background = 'rgba(148,163,184,0.1)';
    badge.style.color = '#94a3b8';
  }
  loadCompaniesIntoDatalist();
  loadAttendeesIntoDatalist();
}

// ── Storage ──────────────────────────────────────────────────────
async function saveCompany(name, id, password) {
  const key = sanitizeKey(name);
  const payload = { name, id, password };
  if (useFirebase) await db.ref(`companies/${key}`).set(payload);
  else { const d = getLocal(COMPANIES_KEY); d[key] = payload; setLocal(COMPANIES_KEY, d); }
}

async function saveAttendee(name) {
  const key = sanitizeKey(name);
  if (useFirebase) await db.ref(`attendees/${key}`).set({ name });
  else { const d = getLocal(ATTENDEES_KEY); d[key] = { name }; setLocal(ATTENDEES_KEY, d); }
}

async function getAllCompanies() {
  if (useFirebase) { const s = await db.ref('companies').once('value'); return s.val() || {}; }
  return getLocal(COMPANIES_KEY);
}

async function getAllAttendees() {
  if (useFirebase) { const s = await db.ref('attendees').once('value'); return s.val() || {}; }
  return getLocal(ATTENDEES_KEY);
}

function getLocal(k) { try { return JSON.parse(localStorage.getItem(k)) || {}; } catch { return {}; } }
function setLocal(k, v) { localStorage.setItem(k, JSON.stringify(v)); }
function sanitizeKey(s) { return String(s).replace(/[.#$/\[\]]/g, '_').trim(); }

// ── Datalists ─────────────────────────────────────────────────────
async function loadCompaniesIntoDatalist() {
  try {
    const cos = await getAllCompanies();
    const dl = document.getElementById('companyList');
    dl.innerHTML = Object.values(cos).map(c => `<option value="${esc(c.name)}"></option>`).join('');
  } catch (e) { }
}

async function loadAttendeesIntoDatalist() {
  try {
    const att = await getAllAttendees();
    const dl = document.getElementById('attendeeList');
    dl.innerHTML = Object.values(att).map(a => `<option value="${esc(a.name)}"></option>`).join('');
  } catch (e) { }
}

// ── Company Autofill ─────────────────────────────────────────────
function setupCompanyAutofill() {
  document.getElementById('companyName').addEventListener('change', async () => {
    const typed = document.getElementById('companyName').value.trim();
    if (!typed) return;
    try {
      const companies = await getAllCompanies();
      const hit = companies[sanitizeKey(typed)];
      if (hit) {
        document.getElementById('companyId').value = hit.id;
        document.getElementById('companyPassword').value = hit.password;
        showToast('Company auto-filled from saved data!', 'info');
      }
    } catch (e) { }
  });
}

// ── Smart Text Parser ────────────────────────────────────────────
// Parses pasted batch text and fills form fields
function smartParseFill() {
  const raw = document.getElementById('smartPasteText').value;
  if (!raw.trim()) { showToast('Please paste some batch details first', 'error'); return; }

  let filled = 0;

  // IP / Company Name
  const ipMatch = raw.match(/ip\s*name\s*[:\-]\s*(.+)/i);
  if (ipMatch) {
    document.getElementById('companyName').value = ipMatch[1].trim();
    // Trigger autofill check
    document.getElementById('companyName').dispatchEvent(new Event('change'));
    filled++;
  }

  // Batch ID
  const batchIdMatch = raw.match(/batch\s*id\s*[:\-]\s*(\S+)/i);
  if (batchIdMatch) {
    document.getElementById('batchId').value = batchIdMatch[1].trim();
    filled++;
  }

  // Batch Start Date  (DD-MM-YYYY)
  const startMatch = raw.match(/batch\s*start\s*date\s*[:\-]\s*(\d{1,2}[-\/]\d{1,2}[-\/]\d{4})/i);
  if (startMatch) {
    document.getElementById('batchStartDate').value = parseDateToISO(startMatch[1]);
    filled++;
  }

  // Batch End Date  (DD-MM-YYYY)
  const endMatch = raw.match(/batch\s*end\s*date\s*[:\-]\s*(\d{1,2}[-\/]\d{1,2}[-\/]\d{4})/i);
  if (endMatch) {
    document.getElementById('batchEndDate').value = parseDateToISO(endMatch[1]);
    filled++;
  }

  // Batch Timings  (HH:MM:SS-HH:MM:SS  or  HH:MM-HH:MM)
  const timingMatch = raw.match(/batch\s*timings?\s*[:\-]\s*(\d{1,2}:\d{2}(?::\d{2})?)\s*[-–to]+\s*(\d{1,2}:\d{2}(?::\d{2})?)/i);
  if (timingMatch) {
    document.getElementById('batchStartTime').value = timingMatch[1].substring(0, 5); // HH:MM
    document.getElementById('batchEndTime').value = timingMatch[2].substring(0, 5);
    filled++;
  }

  // Total Students — not directly a form field; just show info toast
  const studentMatch = raw.match(/total\s*students\s*[:\-]\s*(\d+)/i);
  if (studentMatch) {
    showToast(`ℹ️ Total Students expected: ${studentMatch[1]} (will be counted from Excel)`, 'info');
  }

  if (filled > 0) {
    const res = document.getElementById('parseResult');
    res.style.display = 'inline';
    setTimeout(() => res.style.display = 'none', 4000);
    showToast(`✅ ${filled} field(s) auto-filled from pasted text!`, 'success');
  } else {
    showToast('Could not parse any fields. Check the text format.', 'error');
  }
}

// Convert DD-MM-YYYY or DD/MM/YYYY → YYYY-MM-DD (for <input type=date>)
function parseDateToISO(str) {
  const parts = str.split(/[-\/]/);
  if (parts.length !== 3) return '';
  const [d, m, y] = parts;
  return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
}

// ── Step Navigation ──────────────────────────────────────────────
function goToStep(n) {
  document.querySelectorAll('.step-section').forEach((s, i) => s.classList.toggle('active', i + 1 === n));
  document.querySelectorAll('.step-item').forEach((t, i) => {
    t.classList.remove('active', 'completed');
    if (i + 1 < n) { t.classList.add('completed'); t.querySelector('.step-num').textContent = '✓'; }
    if (i + 1 === n) { t.classList.add('active'); t.querySelector('.step-num').textContent = i + 1; }
    if (i + 1 > n) { t.querySelector('.step-num').textContent = i + 1; }
  });
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ── Step 1 → 2 ──────────────────────────────────────────────────
async function handleStep1Next() {
  const required = [
    ['companyName', 'IP / Company Name'], ['companyId', 'Company / IDP ID'],
    ['companyPassword', 'Company Password'], ['batchId', 'Batch ID'],
    ['batchStartDate', 'Batch Start Date'], ['batchEndDate', 'Batch End Date'],
    ['batchStartTime', 'Batch Start Time'], ['batchEndTime', 'Batch End Time'],
    ['attendeeName', 'Attendee Name'], ['attendeeStartDate', 'Attendee Start Date'],
  ];
  for (const [id, label] of required) {
    if (!document.getElementById(id).value.trim()) {
      document.getElementById(id).focus();
      showToast(`Please fill in: ${label}`, 'error');
      return;
    }
  }

  try {
    await saveCompany(
      document.getElementById('companyName').value.trim(),
      document.getElementById('companyId').value.trim(),
      document.getElementById('companyPassword').value.trim()
    );
    await saveAttendee(document.getElementById('attendeeName').value.trim());
    await loadCompaniesIntoDatalist();
    await loadAttendeesIntoDatalist();
    showToast('Details saved successfully!', 'success');
  } catch (e) {
    showToast('Could not save to database, but proceeding…', 'info');
  }

  const startTime = document.getElementById('batchStartTime').value; // HH:MM
  const endTime = document.getElementById('batchEndTime').value;

  formData = {
    companyName: document.getElementById('companyName').value.trim(),
    companyId: document.getElementById('companyId').value.trim(),
    companyPassword: document.getElementById('companyPassword').value.trim(),
    batchId: document.getElementById('batchId').value.trim(),
    batchStartDate: document.getElementById('batchStartDate').value,
    batchEndDate: document.getElementById('batchEndDate').value,
    batchStartTime: startTime,
    batchEndTime: endTime,
    batchTimings: `${startTime} – ${endTime}`,
    attendeeName: document.getElementById('attendeeName').value.trim(),
    attendeeStartDate: document.getElementById('attendeeStartDate').value,
    batchHours: calcDurationHours(startTime, endTime),
  };

  formData.expectedEndDate = calcExpectedEndDate(formData.batchStartDate, formData.batchHours);
  goToStep(2);
}

// ── Duration & Expected End ───────────────────────────────────────
function calcDurationHours(startTime, endTime) {
  if (!startTime || !endTime) return 0;
  const [sh, sm] = startTime.split(':').map(Number);
  const [eh, em] = endTime.split(':').map(Number);
  return (eh * 60 + em - sh * 60 - sm) / 60;
}

function calcExpectedEndDate(startDateStr, hours) {
  if (!startDateStr || !hours) return 'N/A';
  let days = 0;
  if (hours === 9) days = 30;
  else if (hours === 6) days = 40;
  else return `N/A (${hours} hrs/day — only 6 or 9 supported)`;
  const d = new Date(startDateStr);
  d.setDate(d.getDate() + days);
  return formatDate(d);
}

// ── File Select ──────────────────────────────────────────────────
function handleFileSelect(e) {
  const file = e.target.files[0];
  if (file) processFile(file);
}

function setupDragDrop() {
  const zone = document.getElementById('uploadZone');
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', e => {
    e.preventDefault(); zone.classList.remove('drag-over');
    if (e.dataTransfer.files[0]) processFile(e.dataTransfer.files[0]);
  });
}

function processFile(file) {
  if (!file.name.match(/\.(xlsx|xls)$/i)) { showToast('Please upload a valid .xlsx or .xls file', 'error'); return; }
  document.getElementById('selectedFileName').textContent = file.name;
  document.getElementById('selectedFileSize').textContent = formatFileSize(file.size);
  document.getElementById('fileSelectedBadge').classList.add('show');
  document.getElementById('btn-next-2').disabled = false;
  const reader = new FileReader();
  reader.onload = e => parseExcel(e.target.result);
  reader.readAsArrayBuffer(file);
}

function clearFile() {
  document.getElementById('excelFile').value = '';
  document.getElementById('fileSelectedBadge').classList.remove('show');
  document.getElementById('btn-next-2').disabled = true;
  studentData = [];
}

// ── Excel Parsing ────────────────────────────────────────────────
function parseExcel(buffer) {
  try {
    const wb = XLSX.read(buffer, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    studentData = [];
    for (let i = 2; i < rows.length; i++) {
      const row = rows[i];
      const name = String(row[1] || '').trim();
      const refId = String(row[2] || '').trim();
      if (!name && !refId) continue;
      studentData.push({
        sno: row[0] !== '' ? row[0] : studentData.length + 1,
        name,
        refId,
        trimmedRefId: trimRefId(refId),
      });
    }
    // ── Sort by Attendance Ref.ID numerically 0→9 ──
    studentData.sort((a, b) => {
      const na = parseInt(a.trimmedRefId, 10);
      const nb = parseInt(b.trimmedRefId, 10);
      if (!isNaN(na) && !isNaN(nb)) return na - nb;
      return a.trimmedRefId.localeCompare(b.trimmedRefId);
    });
    // Re-number S.No after sort
    studentData.forEach((s, i) => s.sno = i + 1);
    showToast(`Parsed & sorted ${studentData.length} students!`, 'success');
  } catch (e) {
    showToast('Error reading Excel: ' + e.message, 'error');
  }
}

function trimRefId(refId) {
  if (!refId) return '';
  const parts = refId.split('/');
  return parts[parts.length - 1].trim();
}

// ── Step 2 → 3 ──────────────────────────────────────────────────
function handleStep2Next() {
  if (!studentData.length) { showToast('No student data found in the Excel file.', 'error'); return; }
  buildPreview();
  goToStep(3);
}

// ── Build Preview ────────────────────────────────────────────────
function buildPreview() {
  document.getElementById('summaryChips').innerHTML = `
    <div class="stat-chip">📦 Batch <span class="stat-value">&nbsp;${esc(formData.batchId)}</span></div>
    <div class="stat-chip">🏢 <span class="stat-value">${esc(formData.companyName)}</span></div>
    <div class="stat-chip">👥 Students <span class="stat-value">&nbsp;${studentData.length}</span></div>
    <div class="stat-chip">📅 Expected End <span class="stat-value">&nbsp;${esc(formData.expectedEndDate)}</span></div>
    <div class="stat-chip">⏱ Timings <span class="stat-value">&nbsp;${esc(formData.batchTimings)}</span></div>
  `;

  const details = [
    ['🏢 IP / Company Name', formData.companyName],
    ['🆔 Company / IDP ID', formData.companyId],
    ['🔑 Company Password', formData.companyPassword],
    ['📦 Batch ID', formData.batchId],
    ['📅 Start Date', formatDate(new Date(formData.batchStartDate))],
    ['📅 End Date', formatDate(new Date(formData.batchEndDate))],
    ['⏱ Batch Timings', formData.batchTimings],
    ['📅 Expected End Date', formData.expectedEndDate],
    ['👥 No. of Students', studentData.length],
    ['👤 Attendee Name', formData.attendeeName],
    ['📅 Attendee Start', formatDate(new Date(formData.attendeeStartDate))],
  ];
  document.getElementById('detailsSummary').innerHTML = details.map(([label, val]) =>
    `<div style="display:flex;flex-direction:column;gap:2px">
       <span style="font-size:0.7rem;color:#64748b;font-weight:600;text-transform:uppercase;letter-spacing:.05em">${label}</span>
       <span style="font-size:0.88rem;color:#f0f4ff;font-weight:500">${esc(String(val))}</span>
     </div>`
  ).join('');

  document.getElementById('previewTableBody').innerHTML = studentData.map(s =>
    `<tr>
       <td>${s.sno}</td>
       <td>${esc(s.name)}</td>
       <td>${esc(s.trimmedRefId)}</td>
       <td></td><td></td><td></td>
     </tr>`
  ).join('');
}

// ── PDF Generation ───────────────────────────────────────────────
// B&W · Times New Roman · white table · sorted by Ref.ID
async function generatePDF() {
  if (!studentData.length) { showToast('No data to generate PDF', 'error'); return; }

  const btn = document.getElementById('btn-generate');
  const txt = document.getElementById('generateBtnText');
  const spin = document.getElementById('generateSpinner');
  btn.disabled = true; txt.textContent = 'Generating…'; spin.style.display = 'block';
  await new Promise(r => setTimeout(r, 50));

  try {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    const PW = 297, PM = 12;

    // ─ Fonts: Times New Roman ─
    doc.setFont('times', 'normal');

    // ─ HEADER BOX ─
    const headerH = 72;
    // Outer box
    doc.setDrawColor(0);
    doc.setLineWidth(0.8);
    doc.rect(PM, 10, PW - PM * 2, headerH);

    // Title
    doc.setFont('times', 'bold');
    doc.setFontSize(15);
    doc.setTextColor(0);
    doc.text(`${formData.batchId}  —  ${formData.companyName}`, PW / 2, 19, { align: 'center', maxWidth: PW - PM * 4 });

    // Thin divider under title
    doc.setLineWidth(0.4);
    doc.line(PM, 21, PW - PM, 21);

    // Highlighted key row (Company Name, Batch ID, Start Date, Timings) — B&W: gray fill
    const keyY = 23, keyH = 11;
    const keyFields = [
      `Company/IP: ${formData.companyName}`,
      `Batch ID: ${formData.batchId}`,
      `Start Date: ${formatDate(new Date(formData.batchStartDate))}`,
      `Timings: ${formData.batchTimings}`,
    ];
    const keyW = (PW - PM * 2) / keyFields.length;
    keyFields.forEach((text, i) => {
      const x = PM + i * keyW;
      // Gray fill for highlighted fields
      doc.setFillColor(220, 220, 220);
      doc.rect(x, keyY, keyW, keyH, 'F');
      doc.setDrawColor(0);
      doc.setLineWidth(0.3);
      doc.rect(x, keyY, keyW, keyH);
      doc.setFont('times', 'bold');
      doc.setFontSize(8.5);
      doc.setTextColor(0);
      doc.text(text, x + keyW / 2, keyY + 7.2, { align: 'center', maxWidth: keyW - 4 });
    });

    // Details grid (2 rows × 4 cols)
    const gridY = keyY + keyH + 2;
    const detailItems = [
      ['Company / IDP ID', formData.companyId],
      ['Company Password', formData.companyPassword],
      ['Batch End Date', formatDate(new Date(formData.batchEndDate))],
      ['Expected End Date', formData.expectedEndDate],
      ['No. of Students', String(studentData.length)],
      ['Attendee Name', formData.attendeeName],
      ['Attendee Start Date', formatDate(new Date(formData.attendeeStartDate))],
      ['', ''],
    ];
    const dCols = 4;
    const dW = (PW - PM * 2) / dCols;
    detailItems.forEach(([label, val], i) => {
      const col = i % dCols;
      const row = Math.floor(i / dCols);
      const x = PM + col * dW + 3;
      const y = gridY + row * 11;
      if (label) {
        doc.setFont('times', 'bold');
        doc.setFontSize(6.5);
        doc.setTextColor(80);
        doc.text(label.toUpperCase(), x, y + 4);
        doc.setFont('times', 'normal');
        doc.setFontSize(8.5);
        doc.setTextColor(0);
        doc.text(String(val), x, y + 9.5, { maxWidth: dW - 6 });
      }
    });

    // ─ STUDENT TABLE ─
    const tableStartY = 10 + headerH + 5;
    const tableData = studentData.map(s => [s.sno, s.name, s.trimmedRefId, '', '', '']);

    doc.autoTable({
      head: [['S.No', 'Student Name', 'Attendance Ref. ID', 'Absent', 'Eye Number', 'Remark']],
      body: tableData,
      startY: tableStartY,
      margin: { left: PM, right: PM },
      styles: {
        font: 'times',
        fontSize: 9,
        textColor: [0, 0, 0],
        fillColor: [255, 255, 255],   // white background
        lineColor: [0, 0, 0],
        lineWidth: 0.3,
        cellPadding: { top: 3, bottom: 3, left: 4, right: 4 },
        overflow: 'linebreak',
      },
      headStyles: {
        fillColor: [0, 0, 0],         // black header
        textColor: [255, 255, 255],   // white text
        fontStyle: 'bold',
        fontSize: 9,
        halign: 'center',
      },
      alternateRowStyles: {
        fillColor: [255, 255, 255],     // keep white (no alternating for B&W print)
      },
      columnStyles: {
        0: { cellWidth: 14, halign: 'center' },
        1: { cellWidth: 80, fontStyle: 'bold' },
        2: { cellWidth: 42, halign: 'center', font: 'courier' },
        3: { cellWidth: 22, halign: 'center' },
        4: { cellWidth: 22, halign: 'center' },
        5: { cellWidth: 'auto' },
      },
      didDrawPage: ({ pageNumber }) => {
        const total = doc.internal.getNumberOfPages();
        doc.setFont('times', 'italic');
        doc.setFontSize(7);
        doc.setTextColor(100);
        doc.text(
          `${formData.companyName}  ·  Batch: ${formData.batchId}  ·  Page ${pageNumber} of ${total}`,
          PW / 2, doc.internal.pageSize.height - 5,
          { align: 'center' }
        );
      },
    });

    // ─ Save ─
    const safeCo = formData.companyName.replace(/[^a-zA-Z0-9 \-]/g, '').trim();
    const safeBatch = formData.batchId.replace(/[^a-zA-Z0-9 \-]/g, '').trim();
    doc.save(`${safeBatch} - ${safeCo}.pdf`);
    showToast(`✅ PDF saved as "${safeBatch} - ${safeCo}.pdf"`, 'success');

  } catch (e) {
    console.error(e);
    showToast('PDF generation failed: ' + e.message, 'error');
  } finally {
    btn.disabled = false;
    txt.textContent = '🖨️ Generate PDF';
    spin.style.display = 'none';
  }
}

// ── Utilities ────────────────────────────────────────────────────
function formatDate(d) {
  if (!d || isNaN(d)) return '';
  return d.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
}

function formatFileSize(b) {
  if (b < 1024) return b + ' B';
  if (b < 1048576) return (b / 1024).toFixed(1) + ' KB';
  return (b / 1048576).toFixed(1) + ' MB';
}

function esc(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ── Toast ────────────────────────────────────────────────────────
function showToast(msg, type = 'info') {
  const c = document.getElementById('toast-container');
  const t = document.createElement('div');
  t.className = `toast toast-${type}`;
  t.innerHTML = `<span>${{ success: '✅', error: '❌', info: 'ℹ️' }[type] || 'ℹ️'}</span><span>${msg}</span>`;
  c.appendChild(t);
  setTimeout(() => { t.style.opacity = '0'; t.style.transform = 'translateX(120%)'; setTimeout(() => t.remove(), 300); }, 3500);
}
