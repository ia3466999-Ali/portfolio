/**
 * ============================================================================
 * IMRAN ALI PORTFOLIO & LAFZ-E-ALI – LIVE GOOGLE SHEET & EXCEL WEB APP
 * ============================================================================
 * 
 * ATTACHED SPREADSHEET:
 * https://docs.google.com/spreadsheets/d/1YSNR8o1CQ-6GOBOVjayhZ3qRc4tzw8zhD2_7ubQTCZw/edit?gid=0#gid=0
 * 
 * QUICK 1-MINUTE SETUP IN YOUR GOOGLE SHEET:
 * ----------------------------------------------------------------------------
 * 1. Open your Google Sheet:
 *    https://docs.google.com/spreadsheets/d/1YSNR8o1CQ-6GOBOVjayhZ3qRc4tzw8zhD2_7ubQTCZw/edit
 * 2. In the top menu: Click "Extensions" > "Apps Script".
 * 3. Delete everything inside the editor and PASTE THIS ENTIRE FILE.
 * 4. Click the blue "Deploy" button (top right) > "New deployment".
 * 5. Click the gear icon (Select type) > choose "Web app".
 * 6. Set:
 *      - Description: "Live Portfolio Webhook"
 *      - Execute as: "Me" (your email)
 *      - Who has access: "Anyone" (IMPORTANT: allows visitors to submit)
 * 7. Click "Deploy", approve Google permissions (Click Advanced > Go to Untitled project > Allow).
 * 8. Copy the Web App URL (ends with `/exec`).
 * 9. Paste that URL into index.html inside CONFIG.WEBHOOK_URL.
 * 
 * Result: Every form submission instantly appends a new row to your sheet
 * and sends an email alert to ia3466999@gmail.com.
 * ============================================================================
 */

const SPREADSHEET_ID  = '1YSNR8o1CQ-6GOBOVjayhZ3qRc4tzw8zhD2_7ubQTCZw';
const SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1YSNR8o1CQ-6GOBOVjayhZ3qRc4tzw8zhD2_7ubQTCZw/edit';
const OWNER_EMAIL     = 'ia3466999@gmail.com';
const OWNER_SECRET    = 'imran2026';
const SHEET_NAME      = 'Live Submissions';

const HEADERS = [
  'ID',
  'Date & Time (IST)',
  'Full Name',
  'Email Address',
  'Phone / WhatsApp',
  'Category / Purpose',
  'Subject / Title',
  'Message / Verses',
  'Status'
];

/**
 * Returns the active sheet, creating headers and pretty formatting if brand new.
 */
function getSheet_() {
  let ss;
  try {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (err) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheet.setName(SHEET_NAME);
  }

  // If sheet is empty, initialize headers and styling
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    
    // Style the header row
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setBackground('#0F172A'); // Dark slate
    headerRange.setFontColor('#38D5F0'); // Cyan accent text
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(11);
    headerRange.setFontFamily('Arial');
    headerRange.setHorizontalAlignment('center');
    headerRange.setVerticalAlignment('middle');
    sheet.setRowHeight(1, 38);
    sheet.setFrozenRows(1);

    // Auto-fit column widths
    for (let c = 1; c <= HEADERS.length; c++) {
      sheet.autoResizeColumn(c);
    }
  }

  return sheet;
}

/**
 * Sanitizes input to prevent spreadsheet formula injection attacks (=, +, -, @)
 */
function sanitize_(val, maxLen) {
  val = String(val || '').trim().slice(0, maxLen || 2000);
  if (/^[=+\-@]/.test(val)) {
    return "'" + val;
  }
  return val;
}

/**
 * Formats a Date object into IST string: "DD/MM/YYYY, HH:MM:SS"
 */
function formatIST_(d) {
  return Utilities.formatDate(d, 'Asia/Kolkata', 'dd/MM/yyyy, hh:mm:ss a');
}

/**
 * Handles POST requests from the website contact & shayari submission form.
 */
function doPost(e) {
  try {
    let data = {};

    // Support both JSON payload and form-urlencoded
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    // Hidden anti-spam honeypot field
    if (data.website) {
      return jsonResponse_({ ok: true, spam: true });
    }

    const sheet = getSheet_();
    const submissionId = 'INQ-' + new Date().getTime().toString(36).toUpperCase();
    const now = new Date();
    const formattedTime = formatIST_(now);

    const name     = sanitize_(data.name || 'Anonymous', 120);
    const email    = sanitize_(data.email || 'N/A', 150);
    const phone    = sanitize_(data.phone || 'N/A', 40);
    const category = sanitize_(data.category || 'General Message', 100);
    const subject  = sanitize_(data.subject || 'Portfolio Inquiry', 150);
    const message  = sanitize_(data.message || '', 3000);
    const status   = 'New';

    // Append new row to Google Sheet
    sheet.appendRow([
      submissionId,
      formattedTime,
      name,
      email,
      phone,
      category,
      subject,
      message,
      status
    ]);

    // Format new row
    const lastRow = sheet.getLastRow();
    const dataRange = sheet.getRange(lastRow, 1, 1, HEADERS.length);
    dataRange.setFontFamily('Arial');
    dataRange.setFontSize(10);
    dataRange.setVerticalAlignment('top');

    // Send instant email notification to Imran Ali
    try {
      const emailSubject = `[Lafz-e-Ali / Portfolio] ${category} – ${name}`;
      const emailBody = [
        `You received a new submission on your portfolio website!`,
        `-------------------------------------------------------`,
        `ID:        ${submissionId}`,
        `Time:      ${formattedTime}`,
        `Name:      ${name}`,
        `Email:     ${email}`,
        `Phone:     ${phone}`,
        `Category:  ${category}`,
        `Subject:   ${subject}`,
        `-------------------------------------------------------`,
        `Message / Shayari:`,
        `${message}`,
        `-------------------------------------------------------`,
        `View live Google Sheet: ${SPREADSHEET_URL}`
      ].join('\n');

      MailApp.sendEmail({
        to: OWNER_EMAIL,
        subject: emailSubject,
        body: emailBody,
        replyTo: email.includes('@') ? email : OWNER_EMAIL
      });
    } catch (mailErr) {
      Logger.log('Mail send error: ' + mailErr);
    }

    return jsonResponse_({
      ok: true,
      id: submissionId,
      message: 'Submission successfully recorded in live Google Sheet!'
    });

  } catch (error) {
    return jsonResponse_({
      ok: false,
      error: error.toString()
    });
  }
}

/**
 * Handles GET requests:
 * 1. Health check & live connection test: ?ping=1
 * 2. View live messages table: ?key=imran2026
 * 3. Export to CSV: ?key=imran2026&format=csv
 */
function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};

    // 1. Connection Ping Test
    if (params.ping) {
      return jsonResponse_({
        ok: true,
        status: 'online',
        sheetId: SPREADSHEET_ID,
        sheetUrl: SPREADSHEET_URL,
        totalRows: Math.max(0, getSheet_().getLastRow() - 1)
      });
    }

    // 2. Security Check for reading data
    if (!params.key || params.key !== OWNER_SECRET) {
      return jsonResponse_({
        ok: false,
        error: 'Unauthorized: Invalid owner passphrase.'
      });
    }

    const sheet = getSheet_();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return jsonResponse_({ ok: true, rows: [], count: 0 });
    }

    const rows = data.slice(1); // Skip headers

    // 3. CSV format output
    if (params.format === 'csv') {
      let csv = data.map(row => 
        row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')
      ).join('\r\n');
      return ContentService.createTextOutput(csv)
        .setMimeType(ContentService.MimeType.CSV)
        .downloadAsFile('Imran_Ali_Portfolio_Submissions.csv');
    }

    // Return newest entries first
    return jsonResponse_({
      ok: true,
      count: rows.length,
      headers: HEADERS,
      rows: rows.reverse()
    });

  } catch (error) {
    return jsonResponse_({
      ok: false,
      error: error.toString()
    });
  }
}

/**
 * Helper to construct JSON response with CORS headers
 */
function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
