// ═══════════════════════════════════════════════════════════════
// YCRJ and Associates — Peer Review Tool
// Google Apps Script — Web App Bridge + Daily Email Trigger
// ═══════════════════════════════════════════════════════════════
// SETUP INSTRUCTIONS:
// 1. Open your Google Sheet → Extensions → Apps Script
// 2. Paste this entire file into Code.gs
// 3. Click Deploy → New Deployment → Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 4. Copy the Web App URL → paste into the tool's Setup tab
// 5. Add a time trigger: Triggers → Add Trigger
//    - Function: dailyReport
//    - Event: Time-driven → Day timer → 5 PM to 6 PM
// ═══════════════════════════════════════════════════════════════

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SENDER_EMAIL = 'office.ycrj@gmail.com';
const FIRM_NAME = 'YCRJ and Associates';

// ── Handle POST requests from the browser tool ──
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    if (payload.type === 'review') {
      saveReviewData(payload.data, payload.audit);
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (payload.type === 'email') {
      const results = sendPartnerEmails(payload.partners, payload.firm, payload.sender);
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok', results }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'error', msg: 'Unknown type' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', msg: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Handle GET requests (for testing) ──
function doGet(e) {
  return ContentService.createTextOutput('YCRJ Peer Review — Apps Script is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ═══════════════════════════════════════════════════════════════
// SAVE REVIEW DATA TO GOOGLE SHEETS
// ═══════════════════════════════════════════════════════════════
function saveReviewData(reviewData, auditLog) {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // ── Review_Data sheet ──
  let rdSheet = ss.getSheetByName('Review_Data');
  if (!rdSheet) {
    rdSheet = ss.insertSheet('Review_Data');
    rdSheet.appendRow(['UDIN','Turnover','Borrowings','Networth','Completed','Completed By','Completed Date','Saved By','Saved At','Items JSON']);
  }

  const udins = Object.keys(reviewData);
  udins.forEach(udin => {
    const rd = reviewData[udin];
    const existing = findRow(rdSheet, udin, 1);
    const row = [
      udin,
      rd.turnover || '',
      rd.borrowings || '',
      rd.networth || '',
      rd.completed ? 'Yes' : 'No',
      rd.completedBy || '',
      rd.completedDate ? new Date(rd.completedDate).toLocaleString('en-IN') : '',
      rd.savedBy || '',
      rd.savedAt ? new Date(rd.savedAt).toLocaleString('en-IN') : '',
      JSON.stringify(rd.items || {})
    ];
    if (existing > 0) {
      rdSheet.getRange(existing, 1, 1, row.length).setValues([row]);
    } else {
      rdSheet.appendRow(row);
    }
  });

  // ── Audit_Log sheet ──
  let alSheet = ss.getSheetByName('Audit_Log');
  if (!alSheet) {
    alSheet = ss.insertSheet('Audit_Log');
    alSheet.appendRow(['Timestamp','Username','Action','UDIN','Detail','Old Value','New Value']);
    alSheet.setFrozenRows(1);
  }

  // Clear and rewrite audit log (keep last 500)
  const lastRow = alSheet.getLastRow();
  if (lastRow > 1) alSheet.getRange(2, 1, lastRow - 1, 7).clearContent();

  if (auditLog && auditLog.length) {
    const logRows = auditLog.slice(0, 500).map(e => [
      e.ts ? new Date(e.ts).toLocaleString('en-IN') : '',
      e.user || '',
      e.action || '',
      e.udin || '',
      e.detail || '',
      e.oldVal || '',
      e.newVal || ''
    ]);
    if (logRows.length) alSheet.getRange(2, 1, logRows.length, 7).setValues(logRows);
  }
}

function findRow(sheet, value, col) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][col - 1]).trim() === String(value).trim()) return i + 1;
  }
  return -1;
}

// ═══════════════════════════════════════════════════════════════
// DAILY EMAIL TRIGGER — runs at 5:15 PM IST
// ═══════════════════════════════════════════════════════════════
function dailyReport() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const udinSheet = ss.getSheetByName('UDIN_Master');
  const rdSheet = ss.getSheetByName('Review_Data');

  if (!udinSheet) {
    Logger.log('UDIN_Master sheet not found');
    return;
  }

  // Read UDIN data
  const udinData = sheetToObjects(udinSheet);

  // Read review data
  const reviewMap = {};
  if (rdSheet && rdSheet.getLastRow() > 1) {
    const rdData = sheetToObjects(rdSheet);
    rdData.forEach(r => {
      reviewMap[r['UDIN']] = {
        turnover: r['Turnover'],
        borrowings: r['Borrowings'],
        networth: r['Networth'],
        completed: r['Completed'] === 'Yes',
        completedBy: r['Completed By'],
        completedDate: r['Completed Date']
      };
    });
  }

  // Build partner summaries
  const partners = [
    { name: 'Vijayendra R Nayak', email: 'ajay.j.231@gmail.com' },
    { name: 'Yashavanth Khanderi', email: 'navimk18@gmail.com' }
  ];

  const today = new Date();
  const todayStr = Utilities.formatDate(today, 'Asia/Kolkata', 'dd MMM yyyy');
  const todayKey = Utilities.formatDate(today, 'Asia/Kolkata', 'dd/MM/yyyy');

  partners.forEach(partner => {
    const pu = udinData.filter(r => r['Partner Name'] === partner.name);
    const total = pu.length;
    const completed = pu.filter(r => {
      const rd = reviewMap[r['UDIN']];
      return rd && rd.completed;
    }).length;
    const pending = total - completed;
    const pendingPct = total ? Math.round((pending / total) * 100) : 0;

    // Completed today
    const todayDone = pu.filter(r => {
      const rd = reviewMap[r['UDIN']];
      if (!rd || !rd.completedDate) return false;
      return rd.completedDate.includes(todayKey);
    }).length;

    const completedList = pu.filter(r => (reviewMap[r['UDIN']] || {}).completed)
      .map(r => `  ✓  ${r['UDIN']}  —  ${r['Client Name'] || ''}  (${r['Sub-type'] || ''})`).join('\n');

    const pendingList = pu.filter(r => !(reviewMap[r['UDIN']] || {}).completed)
      .map(r => `  ○  ${r['UDIN']}  —  ${r['Client Name'] || ''}  (${r['Sub-type'] || ''})`).join('\n');

    const subject = `Peer Review Status — ${partner.name} — ${todayStr}`;

    const body = `Dear ${partner.name},

Please find below your Peer Review status summary as of ${todayStr}.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
FIRM: ${FIRM_NAME}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

SUMMARY
───────────────────────────────────
  Total UDINs Assigned  :  ${total}
  Completed             :  ${completed}
  Pending               :  ${pending}
  Completed Today       :  ${todayDone}
  Pending %             :  ${pendingPct}%

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
COMPLETED UDINs (${completed})
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
${completedList || '  None completed yet.'}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PENDING UDINs (${pending})
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
${pendingList || '  All UDINs completed! Great work.'}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

This is an automated daily report generated by the YCRJ Peer Review System.
For queries, contact: ${SENDER_EMAIL}

Regards,
${FIRM_NAME}`;

    try {
      GmailApp.sendEmail(partner.email, subject, body, {
        name: FIRM_NAME,
        replyTo: SENDER_EMAIL
      });
      Logger.log('Email sent to: ' + partner.email);
    } catch (err) {
      Logger.log('Error sending to ' + partner.email + ': ' + err.message);
    }
  });

  Logger.log('Daily report completed at ' + new Date().toLocaleString());
}

// ── Send emails triggered manually from the tool ──
function sendPartnerEmails(partners, firm, sender) {
  const results = [];
  partners.forEach(p => {
    try {
      const subject = `Peer Review Status — ${p.name} — ${p.date}`;
      const completedList = (p.completedList || []).map(u => `  ✓  ${u}`).join('\n');
      const pendingList = (p.pendingList || []).map(u => `  ○  ${u}`).join('\n');

      const body = `Dear ${p.name},

Please find below your Peer Review status summary as of ${p.date}.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
FIRM: ${firm}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

SUMMARY
───────────────────────────────────
  Total UDINs Assigned  :  ${p.total}
  Completed             :  ${p.completed}
  Pending               :  ${p.pending}
  Completed Today       :  ${p.todayDone}
  Pending %             :  ${p.pendingPct}%

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
COMPLETED UDINs (${p.completed})
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
${completedList || '  None completed yet.'}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PENDING UDINs (${p.pending})
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
${pendingList || '  All UDINs completed! Great work.'}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

This is an automated report from the YCRJ Peer Review System.
For queries, contact: ${sender}

Regards,
${firm}`;

      GmailApp.sendEmail(p.email, subject, body, { name: firm, replyTo: sender });
      results.push({ name: p.name, status: 'sent' });
    } catch (err) {
      results.push({ name: p.name, status: 'failed', error: err.message });
    }
  });
  return results;
}

// ── Helper: Convert sheet to array of objects ──
function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  }).filter(r => r['UDIN'] || r['Partner Name']);
}
