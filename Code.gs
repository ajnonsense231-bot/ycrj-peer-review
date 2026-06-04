// ═══════════════════════════════════════════════════════════════
// YCRJ and Associates — Peer Review Tool
// Google Apps Script · Code.gs · v3.0
// ═══════════════════════════════════════════════════════════════

const FIRM     = 'YCRJ and Associates';
const SENDER   = 'office.ycrj@gmail.com';
const PARTNERS = [
  { name: 'Vijayendra R Nayak',  email: 'ajay.j.231@gmail.com' },
  { name: 'Yashavanth Khanderi', email: 'navimk18@gmail.com'   }
];

// ── All checklist items per Peer Review Tag ──
const SA_ITEMS = ['Statutory Audit Appointment Letter','ADT 1','Engagement Letter','MRL','Checklists','Accountant\'s Compilation Report','Statutory Audit Report / Other Audit Reports','LFAR (Bank)','Financial Statements / Results','AGM Notice & Directors Report','Director\'s Declaration (Loan given is not out of Borrowed Funds)','DIR 8','MBP 1','Signed Audit Notes / ML','Tally data'];
const IA_ITEMS = ['Appointment Letter','Signed Report','Signed Audit Notes / ML'];

const CHECKLISTS = {
  'SA':   SA_ITEMS,
  'SBA':  SA_ITEMS,
  'CSA':  SA_ITEMS,
  'LR':   SA_ITEMS,
  'GA':   SA_ITEMS,
  'TA':   ['Tax Audit Appointment Letter (Signed copy)','Tax Audit MRL (Signed copy)','Computation Sheet','ITR','IT Form','Form 29B','Tax Audit Report (Form 3CD)','Form 10B / 10BB','Tally data'],
  'Cert': ['Certificate Request Letter','Signed Copy of Certificate issued'],
  'IA':   IA_ITEMS,
  'CA':   IA_ITEMS,
  'CR':   IA_ITEMS,
  'StA':  ['Appointment Letter','Signed Report'],
  'RA':   ['Appointment Letter','Signed Audit Report'],
  'TrA':  ['Appointment Letter','Form 10B','Financial Statements','ITR','IT Form']
};

// ── All 30 unique checklist items across all services (for column headers) ──
const ALL_CHECKLIST_ITEMS = [
  // SA / SBA / CSA / LR / GA (15 items)
  'Statutory Audit Appointment Letter',
  'ADT 1',
  'Engagement Letter',
  'MRL',
  'Checklists',
  'Accountant\'s Compilation Report',
  'Statutory Audit Report / Other Audit Reports',
  'LFAR (Bank)',
  'Financial Statements / Results',
  'AGM Notice & Directors Report',
  'Director\'s Declaration (Loan given is not out of Borrowed Funds)',
  'DIR 8',
  'MBP 1',
  'Signed Audit Notes / ML',
  'Tally data',
  // TA unique items (8) — Tally data already listed above
  'Tax Audit Appointment Letter (Signed copy)',
  'Tax Audit MRL (Signed copy)',
  'Computation Sheet',
  'ITR',
  'IT Form',
  'Form 29B',
  'Tax Audit Report (Form 3CD)',
  'Form 10B / 10BB',
  // Cert unique items (2)
  'Certificate Request Letter',
  'Signed Copy of Certificate issued',
  // IA / CA / CR / StA / RA / TrA unique items (3)
  'Appointment Letter',
  'Signed Report',
  'Signed Audit Report',
  // TrA unique items (2) — ITR and IT Form already listed above
  'Form 10B',
  'Financial Statements'
];

// ── Master data columns ──
const MASTER_COLS = ['Sl.No','Partner','FY','UDIN','Date of Signing','Client Name','Document Type','Sub-type','Document Description','Remarks','Turnover','Borrowings','Networth','Peer Review Tagging'];

// ═══════════════════════════════════════════════════════════════
// HANDLE REQUESTS
// ═══════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.type === 'review') {
      writeReviewData(payload.data, payload.audit, payload.udins);
      return ok('Review data saved');
    }
    if (payload.type === 'email') {
      const results = sendEmails(payload.partners, payload.firm, payload.sender);
      return ok(JSON.stringify(results));
    }
    return err('Unknown request type');
  } catch (ex) {
    return err(ex.message);
  }
}

function doGet() {
  return ContentService.createTextOutput('YCRJ Peer Review · Apps Script · Running ✓')
    .setMimeType(ContentService.MimeType.TEXT);
}

function ok(msg)  { return ContentService.createTextOutput(JSON.stringify({ status:'ok',  msg })).setMimeType(ContentService.MimeType.JSON); }
function err(msg) { return ContentService.createTextOutput(JSON.stringify({ status:'err', msg })).setMimeType(ContentService.MimeType.JSON); }

// ═══════════════════════════════════════════════════════════════
// WRITE REVIEW DATA TO SHEETS
// ═══════════════════════════════════════════════════════════════
function writeReviewData(reviewData, auditLog, udinMaster) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Build Review_Data sheet headers ──
  // Fixed columns + all checklist items as columns + trailing status columns
  const fixedCols    = ['Sl.No','Partner','FY','UDIN','Date of Signing','Client Name','Document Type','Sub-type','Document Description','Remarks','Peer Review Tagging','Turnover','Borrowings','Networth','Review Remarks'];
  const trailingCols = ['Status','Completed By','Completed Date','Last Modified By','Last Modified Date'];
  const allHeaders   = [...fixedCols, ...ALL_CHECKLIST_ITEMS, ...trailingCols];

  // ── Ensure Review_Data sheet exists with correct headers ──
  let rdSheet = ss.getSheetByName('Review_Data');
  if (!rdSheet) {
    rdSheet = ss.insertSheet('Review_Data');
  }

  // Always rewrite headers on row 1
  rdSheet.getRange(1, 1, 1, allHeaders.length).setValues([allHeaders]);
  rdSheet.getRange(1, 1, 1, allHeaders.length)
    .setFontWeight('bold')
    .setBackground('#1B4F8A')
    .setFontColor('#FFFFFF');
  rdSheet.setFrozenRows(1);

  // ── Build master data map ──
  const masterMap = {};
  if (udinMaster && udinMaster.length) {
    udinMaster.forEach(r => { if (r['UDIN']) masterMap[String(r['UDIN']).trim()] = r; });
  }

  // ── Write each reviewed UDIN ──
  Object.keys(reviewData).forEach(udin => {
    const rd  = reviewData[udin];
    const rec = masterMap[udin] || {};
    const tag = String(rec['Peer Review Tagging'] || '').trim();
    const checklistItems = CHECKLISTS[tag] || [];

    // Build checklist columns — map item name to Yes/No/NA
    const itemMap = {};
    checklistItems.forEach((item, i) => {
      itemMap[item] = rd.items && rd.items[i] ? rd.items[i] : '';
    });

    // Build the full row
    const row = [
      // Fixed master data columns
      rec['Sl.No']                || '',
      rec['Partner']              || '',
      rec['FY']                   || '',
      udin,
      rec['Date of Signing']      || '',
      rec['Client Name']          || '',
      rec['Document Type']        || '',
      rec['Sub-type']             || '',
      rec['Document Description'] || '',
      rec['Remarks']              || '',
      tag,
      rd.turnover                 || rec['Turnover']    || '',
      rd.borrowings               || rec['Borrowings']  || '',
      rd.networth                 || rec['Networth']    || '',
      rd.reviewRemarks            || '',
      // Checklist item columns (one per item)
      ...ALL_CHECKLIST_ITEMS.map(item => itemMap[item] || ''),
      // Trailing status columns
      rd.completed ? 'Completed' : (Object.keys(rd.items || {}).length > 0 ? 'In Progress' : 'Pending'),
      rd.completedBy              || '',
      rd.completedDate            ? new Date(rd.completedDate).toLocaleString('en-IN') : '',
      rd.savedBy                  || '',
      rd.savedAt                  ? new Date(rd.savedAt).toLocaleString('en-IN') : ''
    ];

    // Find existing row or append
    const existRow = findRow(rdSheet, udin, 4); // UDIN is column 4
    if (existRow > 0) {
      rdSheet.getRange(existRow, 1, 1, row.length).setValues([row]);
    } else {
      rdSheet.appendRow(row);
    }

    // Color Yes/No/NA cells
    colorChecklistCells(rdSheet, existRow > 0 ? existRow : rdSheet.getLastRow(), fixedCols.length + 1, ALL_CHECKLIST_ITEMS.length, itemMap);
  });

  // Auto-resize
  try { rdSheet.autoResizeColumns(1, allHeaders.length); } catch(e) {}

  // ── Audit Log ──
  let alSheet = ss.getSheetByName('Audit_Log');
  if (!alSheet) {
    alSheet = ss.insertSheet('Audit_Log');
    alSheet.appendRow(['Timestamp','Username','Action','UDIN','Detail','Old Value','New Value']);
    alSheet.getRange(1,1,1,7).setFontWeight('bold').setBackground('#1B4F8A').setFontColor('#FFFFFF');
    alSheet.setFrozenRows(1);
  }
  const lastRow = alSheet.getLastRow();
  if (lastRow > 1) alSheet.getRange(2, 1, lastRow - 1, 7).clearContent();
  if (auditLog && auditLog.length) {
    const rows = auditLog.slice(0, 1000).map(e => [
      e.ts     ? new Date(e.ts).toLocaleString('en-IN') : '',
      e.user   || '', e.action || '', e.udin   || '',
      e.detail || '', e.oldVal || '', e.newVal || ''
    ]);
    if (rows.length) alSheet.getRange(2, 1, rows.length, 7).setValues(rows);
  }
  try { alSheet.autoResizeColumns(1, 7); } catch(e) {}
}

// ── Color Yes=green, No=red, NA=grey ──
function colorChecklistCells(sheet, rowNum, startCol, numCols, itemMap) {
  try {
    ALL_CHECKLIST_ITEMS.forEach((item, i) => {
      const val  = itemMap[item] || '';
      const cell = sheet.getRange(rowNum, startCol + i);
      if (val === 'Yes')      { cell.setBackground('#E8F4EB').setFontColor('#276B34'); }
      else if (val === 'No')  { cell.setBackground('#FAE8E8').setFontColor('#A82828'); }
      else if (val === 'NA')  { cell.setBackground('#F0EFE9').setFontColor('#5C5B56'); }
      else                    { cell.setBackground('#FFFFFF').setFontColor('#1A1916'); }
    });
  } catch(e) {}
}

function findRow(sheet, value, col) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][col-1]).trim() === String(value).trim()) return i + 1;
  }
  return -1;
}

// ═══════════════════════════════════════════════════════════════
// DAILY EMAIL TRIGGER — 5:15 PM IST
// ═══════════════════════════════════════════════════════════════
function dailyReport() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const udinSheet = ss.getSheetByName('UDIN_Master');
  const rdSheet   = ss.getSheetByName('Review_Data');
  if (!udinSheet) { Logger.log('UDIN_Master sheet not found'); return; }

  const udinData  = sheetToObjects(udinSheet);
  const reviewMap = {};
  if (rdSheet && rdSheet.getLastRow() > 1) {
    sheetToObjects(rdSheet).forEach(r => {
      if (r['UDIN']) reviewMap[String(r['UDIN']).trim()] = r;
    });
  }

  const today    = new Date();
  const todayStr = Utilities.formatDate(today, 'Asia/Kolkata', 'dd MMM yyyy');
  const todayKey = Utilities.formatDate(today, 'Asia/Kolkata', 'dd MMM yyyy');

  PARTNERS.forEach(partner => {
    const pu         = udinData.filter(r => r['Partner'] === partner.name);
    const total      = pu.length;
    const completed  = pu.filter(r => (reviewMap[r['UDIN']] || {})['Status'] === 'Completed').length;
    const pending    = total - completed;
    const pendingPct = total ? Math.round((pending / total) * 100) : 0;
    const todayDone  = pu.filter(r => {
      const cd = (reviewMap[r['UDIN']] || {})['Completed Date'] || '';
      return String(cd).includes(todayStr.split(' ')[0]);
    }).length;

    const completedList = pu
      .filter(r => (reviewMap[r['UDIN']] || {})['Status'] === 'Completed')
      .map(r => `   ✓  ${r['UDIN']}  —  ${r['Client Name'] || ''}  (${r['Sub-type'] || ''}, ${r['FY'] || ''})`)
      .join('\n');

    const pendingList = pu
      .filter(r => (reviewMap[r['UDIN']] || {})['Status'] !== 'Completed')
      .map(r => `   ○  ${r['UDIN']}  —  ${r['Client Name'] || ''}  (${r['Sub-type'] || ''}, ${r['FY'] || ''})`)
      .join('\n');

    const subject = `Peer Review Status — ${partner.name} — As of ${todayStr}`;
    const body    = buildEmailBody(partner.name, todayStr, total, completed, pending, todayDone, pendingPct, completedList, pendingList);

    try {
      GmailApp.sendEmail(partner.email, subject, body, { name: FIRM, replyTo: SENDER });
      Logger.log('✓ Sent to: ' + partner.email);
    } catch (ex) {
      Logger.log('✗ Failed: ' + partner.email + ' — ' + ex.message);
    }
  });
  Logger.log('Daily report completed: ' + new Date().toLocaleString());
}

// ═══════════════════════════════════════════════════════════════
// MANUAL SEND
// ═══════════════════════════════════════════════════════════════
function sendEmails(partners, firm, sender) {
  const results = [];
  (partners || []).forEach(p => {
    try {
      const completedList = (p.completedList || []).map(u => `   ✓  ${u}`).join('\n');
      const pendingList   = (p.pendingList   || []).map(u => `   ○  ${u}`).join('\n');
      const subject       = `Peer Review Status — ${p.name} — As of ${p.date}`;
      const body          = buildEmailBody(p.name, p.date, p.total, p.completed, p.pending, p.todayDone, p.pendingPct, completedList, pendingList);
      GmailApp.sendEmail(p.email, subject, body, { name: firm || FIRM, replyTo: sender || SENDER });
      results.push({ name: p.name, status: 'sent' });
    } catch (ex) {
      results.push({ name: p.name, status: 'failed', error: ex.message });
    }
  });
  return results;
}

// ═══════════════════════════════════════════════════════════════
// EMAIL BODY
// ═══════════════════════════════════════════════════════════════
function buildEmailBody(name, date, total, completed, pending, todayDone, pendingPct, completedList, pendingList) {
  // Build unique partner link — remove spaces for URL
  const partnerId = name.replace(/\s+/g, '');
  const partnerLink = `https://ajay231.github.io/ycrj-peer-review/partner.html?partner=${partnerId}`;

  return `Dear ${name},

Please find below your Peer Review completion status as of ${date}.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  ${FIRM}
  Peer Review Management System
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  SUMMARY
  ─────────────────────────────────────
  Total UDINs Assigned   :  ${total}
  Completed              :  ${completed}
  Pending                :  ${pending}
  Completed Today        :  ${todayDone}
  Pending %              :  ${pendingPct}%

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  COMPLETED UDINs  (${completed})
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
${completedList || '   None completed yet.'}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  PENDING UDINs  (${pending})
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
${pendingList || '   All UDINs completed — great work!'}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  VIEW YOUR FULL UDIN STATUS ONLINE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  Click the link below to view your complete UDIN list with live status:
  ${partnerLink}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

This is an automated report from the YCRJ Peer Review System.
For queries: ${SENDER}

Regards,
${FIRM}`;
}

// ═══════════════════════════════════════════════════════════════
// HELPER
// ═══════════════════════════════════════════════════════════════
function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  return data.slice(1)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = String(row[i] || '').trim(); });
      return obj;
    })
    .filter(r => r['UDIN'] || r['Partner']);
}
