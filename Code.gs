// ═══════════════════════════════════════════════════════════════
// YCRJ and Associates — Peer Review Tool
// Google Apps Script · Code.gs · v3.0
// ═══════════════════════════════════════════════════════════════

const FIRM     = 'YCRJ and Associates';
const SENDER   = 'info@ycrjca.com';
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

    const subject  = 'Peer Review Status — ' + partner.name + ' — As of ' + todayStr;
    const htmlBody = buildEmailBody(partner.name, todayStr, total, completed, pending, todayDone, pendingPct, [], []);
    try {
      GmailApp.sendEmail(partner.email, subject, '', { name: FIRM, replyTo: SENDER, htmlBody: htmlBody });
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
      const htmlBody = buildEmailBody(p.name, p.date, p.total, p.completed, p.pending, p.todayDone, p.pendingPct, [], []);
      GmailApp.sendEmail(p.email, subject, '', { name: firm || FIRM, replyTo: sender || SENDER, htmlBody: htmlBody });
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
  const partnerId   = name.replace(/\s+/g, '');
  const partnerLink = 'https://ajnonsense231-bot.github.io/ycrj-peer-review/partner.html?partner=' + partnerId;
  const inProgress  = total - completed - pending;

  const htmlBody = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="margin:0;padding:0;background:#F4F3EF;font-family:Arial,sans-serif">' +
    '<table width="100%" cellpadding="0" cellspacing="0" style="background:#F4F3EF;padding:32px 16px">' +
    '<tr><td align="center">' +
    '<table width="560" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:10px;overflow:hidden;border:1px solid #E0DED6">' +

    // Header
    '<tr><td style="background:#1B4F8A;padding:24px 32px">' +
    '<table cellpadding="0" cellspacing="0"><tr>' +
    '<td style="width:40px;height:40px;background:rgba(255,255,255,0.2);border-radius:8px;text-align:center;vertical-align:middle;color:#ffffff;font-weight:700;font-size:14px">YR</td>' +
    '<td style="padding-left:12px"><div style="color:#ffffff;font-size:16px;font-weight:700">' + FIRM + '</div>' +
    '<div style="color:rgba(255,255,255,0.7);font-size:12px;margin-top:2px">Peer Review Management System</div></td>' +
    '</tr></table></td></tr>' +

    // Body
    '<tr><td style="padding:28px 32px">' +
    '<p style="font-size:15px;color:#1A1916;margin:0 0 4px">Dear <strong>' + name + '</strong>,</p>' +
    '<p style="font-size:13px;color:#6B6860;margin:0 0 24px">Here is your peer review completion status as of <strong>' + date + '</strong>.</p>' +

    // Summary Card
    '<div style="background:#F4F3EF;border-radius:10px;padding:20px 24px;margin-bottom:20px">' +
    '<div style="font-size:11px;font-weight:700;color:#9B9990;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:14px">Summary</div>' +

    // Top 3 metrics
    '<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:10px"><tr>' +
    '<td width="32%" style="padding-right:6px"><div style="background:#fff;border-radius:8px;padding:12px;text-align:center">' +
    '<div style="font-size:10px;font-weight:700;color:#9B9990;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:4px">Total</div>' +
    '<div style="font-size:26px;font-weight:700;color:#1B4F8A">' + total + '</div></div></td>' +
    '<td width="32%" style="padding:0 3px"><div style="background:#fff;border-radius:8px;padding:12px;text-align:center">' +
    '<div style="font-size:10px;font-weight:700;color:#9B9990;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:4px">Completed</div>' +
    '<div style="font-size:26px;font-weight:700;color:#276B34">' + completed + '</div></div></td>' +
    '<td width="32%" style="padding-left:6px"><div style="background:#fff;border-radius:8px;padding:12px;text-align:center">' +
    '<div style="font-size:10px;font-weight:700;color:#9B9990;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:4px">Pending</div>' +
    '<div style="font-size:26px;font-weight:700;color:#A82828">' + pending + '</div></div></td>' +
    '</tr></table>' +

    // Bottom 2 metrics
    '<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:16px"><tr>' +
    '<td width="48%" style="padding-right:6px"><div style="background:#fff;border-radius:8px;padding:12px;text-align:center">' +
    '<div style="font-size:10px;font-weight:700;color:#9B9990;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:4px">In Progress</div>' +
    '<div style="font-size:26px;font-weight:700;color:#7A4F00">' + inProgress + '</div></div></td>' +
    '<td width="48%" style="padding-left:6px"><div style="background:#fff;border-radius:8px;padding:12px;text-align:center">' +
    '<div style="font-size:10px;font-weight:700;color:#9B9990;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:4px">Done Today</div>' +
    '<div style="font-size:26px;font-weight:700;color:#276B34">' + todayDone + '</div></div></td>' +
    '</tr></table>' +

    // Progress bar
    '<table width="100%" cellpadding="0" cellspacing="0"><tr>' +
    '<td style="font-size:12px;color:#6B6860;font-weight:500">Completion</td>' +
    '<td align="right" style="font-size:12px;font-weight:700;color:#276B34">' + pendingPct + '% pending</td>' +
    '</tr></table>' +
    '<div style="background:#E0DED6;border-radius:4px;height:8px;margin-top:6px;overflow:hidden">' +
    '<div style="width:' + (100 - pendingPct) + '%;height:100%;background:#276B34;border-radius:4px"></div></div>' +
    '</div>' +

    // CTA Button
    '<div style="text-align:center;margin-bottom:24px">' +
    '<a href="' + partnerLink + '" style="display:inline-block;background:#1B4F8A;color:#ffffff;text-decoration:none;padding:12px 28px;border-radius:8px;font-size:14px;font-weight:700">View Your Full UDIN List &rarr;</a>' +
    '</div>' +

    '<p style="font-size:12px;color:#9B9990;text-align:center;line-height:1.6;margin:0">This is an automated report from the YCRJ Peer Review System.<br>' +
    'For queries: <a href="mailto:' + SENDER + '" style="color:#1B4F8A">' + SENDER + '</a></p>' +
    '</td></tr>' +

    // Footer
    '<tr><td style="background:#F4F3EF;padding:16px 32px;border-top:1px solid #E0DED6;text-align:center">' +
    '<p style="font-size:11px;color:#9B9990;margin:0">YCRJ and Associates &middot; Chartered Accountants &middot; Bangalore</p>' +
    '</td></tr>' +

    '</table></td></tr></table></body></html>';

  return htmlBody;
}
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
