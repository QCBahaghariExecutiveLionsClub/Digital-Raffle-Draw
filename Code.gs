/**
 * ============================================================
 *  LUCKY DRAW RAFFLE — Google Apps Script Backend
 *  File: Code.gs
 * ============================================================
 *  SETUP:
 *  1. Open Google Sheets → Extensions → Apps Script
 *  2. Paste this entire file into Code.gs
 *  3. Update SHEET_ID below with your Google Sheets ID
 *  4. Deploy → New Deployment → Web App
 *     - Execute as: Me
 *     - Who has access: Anyone
 *  5. Copy the Web App URL into your HTML files
 * ============================================================
 */

// ===================== CONFIGURATION ========================
const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID'; // ← Replace this
const TICKET_PRICE = 100; // Price per ticket in PHP

// Sheet tab names (auto-created if missing)
const SHEETS = {
  REGISTRATIONS: 'Registrations',
  WINNERS: 'Winners',
  IMPORTS: 'Imports',
};

// ============================================================
//  ENTRY POINT: doGet — handles all GET requests
// ============================================================
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'ping';

  let result;
  try {
    switch (action) {
      case 'getParticipants': result = getParticipants(); break;
      case 'getAll':          result = getAllRecords();    break;
      case 'getWinners':      result = getWinners();      break;
      case 'getStats':        result = getStats();         break;
      case 'ping':            result = { success: true, message: 'Lucky Draw API is live!' }; break;
      default:                result = { success: false, message: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { success: false, message: err.message };
  }

  return jsonResponse(result);
}

// ============================================================
//  ENTRY POINT: doPost — handles all POST requests
// ============================================================
function doPost(e) {
  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ success: false, message: 'Invalid JSON payload.' });
  }

  const action = payload.action || '';
  let result;

  try {
    switch (action) {
      case 'register':    result = registerParticipant(payload); break;
      case 'saveWinner':  result = saveWinner(payload);          break;
      case 'importExcel': result = importExcelData(payload);     break;
      default:            result = { success: false, message: 'Unknown POST action: ' + action };
    }
  } catch (err) {
    result = { success: false, message: err.message };
  }

  return jsonResponse(result);
}

// ============================================================
//  REGISTER A PARTICIPANT
// ============================================================
function registerParticipant(data) {
  const sheet = getOrCreateSheet(SHEETS.REGISTRATIONS);
  ensureRegistrationsHeader(sheet);

  const qty = parseInt(data.ticketQty) || 1;
  const ticketIds = [];
  const timestamp = new Date();
  const dateStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  // Check for duplicate email per draw (optional — remove if not needed)
  // const existing = findRowsByEmail(sheet, data.email);

  for (let i = 0; i < qty; i++) {
    const ticketId = generateTicketId();
    ticketIds.push(ticketId);

    sheet.appendRow([
      ticketId,                            // A: Ticket ID
      (data.firstName + ' ' + data.lastName).trim(), // B: Full Name
      data.firstName || '',                // C: First Name
      data.lastName || '',                 // D: Last Name
      data.email || '',                    // E: Email
      data.phone || '',                    // F: Phone
      qty,                                 // G: Total Tickets Purchased
      TICKET_PRICE,                        // H: Price per Ticket
      qty * TICKET_PRICE,                  // I: Total Amount
      data.paymentMethod || '',            // J: Payment Method
      data.referenceNo || '',              // K: Reference No
      data.receiptFileName || '',          // L: Receipt File Name
      data.receiptImage ? 'Yes' : 'No',   // M: Has Receipt Image
      dateStr,                             // N: Registered At
      'Active',                            // O: Status
      '',                                  // P: Winner Round (filled later)
      '',                                  // Q: Notes
    ]);

    // Save receipt image to Drive (optional)
    if (data.receiptImage && data.receiptFileName) {
      try { saveReceiptToDrive(ticketId, data.receiptFileName, data.receiptImage); }
      catch(e) { /* Drive saving is optional — don't fail registration */ }
    }
  }

  return {
    success: true,
    ticketIds: ticketIds,
    name: (data.firstName + ' ' + data.lastName).trim(),
    message: 'Registration successful! ' + qty + ' ticket(s) issued.',
  };
}

// ============================================================
//  SAVE RECEIPT IMAGE TO GOOGLE DRIVE
// ============================================================
function saveReceiptToDrive(ticketId, fileName, base64Data) {
  const folder = getDriveFolder('Raffle Receipts');
  const blob = Utilities.newBlob(
    Utilities.base64Decode(base64Data),
    'image/jpeg',
    ticketId + '_' + fileName
  );
  folder.createFile(blob);
}

function getDriveFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

// ============================================================
//  GET PARTICIPANTS (for raffle draw page)
// ============================================================
function getParticipants() {
  const sheet = getOrCreateSheet(SHEETS.REGISTRATIONS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, participants: [] };

  const participants = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[14] === 'Active') { // Status = Active
      participants.push({
        ticket: row[0],
        name:   row[1],
        email:  row[4],
        phone:  row[5],
      });
    }
  }

  return { success: true, participants };
}

// ============================================================
//  GET ALL RECORDS (for admin panel)
// ============================================================
function getAllRecords() {
  const regSheet = getOrCreateSheet(SHEETS.REGISTRATIONS);
  const winSheet = getOrCreateSheet(SHEETS.WINNERS);

  const regData = regSheet.getDataRange().getValues();
  const winData = winSheet.getDataRange().getValues();

  // Build a set of winning ticket IDs
  const winnerMap = {};
  for (let i = 1; i < winData.length; i++) {
    winnerMap[winData[i][1]] = winData[i][3] || 'Winner'; // ticketId → round
  }

  const records = [];
  for (let i = 1; i < regData.length; i++) {
    const row = regData[i];
    const ticketId = row[0];
    records.push({
      ticketId,
      name:          row[1],
      email:         row[4],
      phone:         row[5],
      ticketQty:     row[6],
      paymentMethod: row[9],
      referenceNo:   row[10],
      registeredAt:  row[13] ? String(row[13]).split(' ')[0] : '',
      status:        row[14],
      isWinner:      !!winnerMap[ticketId],
      winnerRound:   winnerMap[ticketId] || null,
    });
  }

  const stats = computeStats(regData, winData);
  return { success: true, records, stats };
}

// ============================================================
//  GET WINNERS
// ============================================================
function getWinners() {
  const sheet = getOrCreateSheet(SHEETS.WINNERS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, winners: [] };

  const winners = data.slice(1).map(row => ({
    drawnAt:   row[0],
    ticketId:  row[1],
    name:      row[2],
    round:     row[3],
    email:     row[4],
    phone:     row[5],
  }));

  return { success: true, winners };
}

// ============================================================
//  SAVE WINNER
// ============================================================
function saveWinner(data) {
  const winSheet = getOrCreateSheet(SHEETS.WINNERS);
  ensureWinnersHeader(winSheet);

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const round = 'Round ' + (winSheet.getLastRow());

  winSheet.appendRow([
    timestamp,
    data.ticket || '',
    data.name   || '',
    round,
    data.email  || '',
    data.phone  || '',
  ]);

  // Mark the ticket in Registrations as winner
  const regSheet = getOrCreateSheet(SHEETS.REGISTRATIONS);
  const regData = regSheet.getDataRange().getValues();
  for (let i = 1; i < regData.length; i++) {
    if (regData[i][0] === data.ticket) {
      regSheet.getRange(i + 1, 16).setValue(round); // Column P: Winner Round
      regSheet.getRange(i + 1, 15).setValue('Winner'); // Column O: Status
      break;
    }
  }

  return { success: true, message: 'Winner saved!', round };
}

// ============================================================
//  IMPORT EXCEL / CSV DATA
// ============================================================
/**
 * This function accepts pre-parsed rows from an uploaded Excel/CSV file.
 * The frontend reads the file and sends rows as JSON.
 * Expected format: array of objects with keys matching registration fields.
 */
function importExcelData(payload) {
  const rows = payload.rows || [];
  if (!rows.length) return { success: false, message: 'No rows to import.' };

  const sheet = getOrCreateSheet(SHEETS.REGISTRATIONS);
  ensureRegistrationsHeader(sheet);

  let imported = 0;
  let skipped = 0;
  const errors = [];

  rows.forEach((row, idx) => {
    try {
      const name   = (row.name || row.Name || row['Full Name'] || '').trim();
      const email  = (row.email || row.Email || '').trim().toLowerCase();
      const phone  = (row.phone || row.Phone || '').trim();
      const method = (row.paymentMethod || row['Payment Method'] || row.payment || 'Import').trim();
      const refNo  = (row.referenceNo || row['Reference No'] || row.reference || 'IMPORTED').trim();
      const qty    = parseInt(row.ticketQty || row.tickets || row.Tickets || 1) || 1;

      if (!name || !email) { skipped++; return; }

      const ticketId = generateTicketId();
      const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      const parts = name.split(' ');
      const firstName = parts[0] || '';
      const lastName  = parts.slice(1).join(' ') || '';

      sheet.appendRow([
        ticketId, name, firstName, lastName, email, phone,
        qty, TICKET_PRICE, qty * TICKET_PRICE,
        method, refNo, '', 'No', dateStr, 'Active', '', 'Imported'
      ]);
      imported++;
    } catch (e) {
      errors.push('Row ' + (idx + 1) + ': ' + e.message);
    }
  });

  return {
    success: true,
    imported, skipped,
    errors: errors.slice(0, 10),
    message: `Import complete. ${imported} records imported, ${skipped} skipped.`,
  };
}

// ============================================================
//  ADMIN: SEARCH TICKETS (server-side, optional)
// ============================================================
function searchTickets(query) {
  const sheet = getOrCreateSheet(SHEETS.REGISTRATIONS);
  const data = sheet.getDataRange().getValues();
  const q = (query || '').toLowerCase();

  const results = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const searchable = [row[0],row[1],row[4],row[5],row[10]].join(' ').toLowerCase();
    if (searchable.includes(q)) {
      results.push({
        ticketId: row[0], name: row[1], email: row[4], phone: row[5],
        ticketQty: row[6], paymentMethod: row[9], referenceNo: row[10],
        registeredAt: String(row[13]).split(' ')[0], status: row[14],
      });
    }
  }
  return { success: true, results };
}

// ============================================================
//  STATS
// ============================================================
function computeStats(regData, winData) {
  const participants = new Set();
  let totalRevenue = 0;
  for (let i = 1; i < regData.length; i++) {
    participants.add(regData[i][4]); // email
    totalRevenue += Number(regData[i][8]) || 0; // total amount
  }
  return {
    total:        Math.max(0, regData.length - 1),
    participants: participants.size,
    winners:      Math.max(0, winData.length - 1),
    revenue:      totalRevenue,
  };
}

function getStats() {
  const regSheet = getOrCreateSheet(SHEETS.REGISTRATIONS);
  const winSheet = getOrCreateSheet(SHEETS.WINNERS);
  const regData = regSheet.getDataRange().getValues();
  const winData = winSheet.getDataRange().getValues();
  return { success: true, stats: computeStats(regData, winData) };
}

// ============================================================
//  UTILITIES
// ============================================================

/** Generate a unique ticket ID: TKT-YYYYMMDD-XXXXXX */
function generateTicketId() {
  const datePart = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  const rand = Math.random().toString(36).substring(2, 8).toUpperCase();
  return 'TKT-' + datePart + '-' + rand;
}

/** Return or create a sheet tab by name */
function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

/** Set up Registrations sheet header if empty */
function ensureRegistrationsHeader(sheet) {
  if (sheet.getLastRow() === 0) {
    const headers = [
      'Ticket ID','Full Name','First Name','Last Name','Email','Phone',
      'Tickets Qty','Price Each','Total Amount','Payment Method','Reference No',
      'Receipt File','Has Receipt','Registered At','Status','Winner Round','Notes'
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#D4A017')
      .setFontColor('#fff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

/** Set up Winners sheet header if empty */
function ensureWinnersHeader(sheet) {
  if (sheet.getLastRow() === 0) {
    const headers = ['Drawn At','Ticket ID','Name','Round','Email','Phone'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#2A9D5C')
      .setFontColor('#fff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

/** JSON response with CORS headers */
function jsonResponse(data) {
  const output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ============================================================
//  MANUAL TRIGGER: Run this function once to initialise sheets
// ============================================================
function initSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const regSheet = getOrCreateSheet(SHEETS.REGISTRATIONS);
  const winSheet = getOrCreateSheet(SHEETS.WINNERS);
  ensureRegistrationsHeader(regSheet);
  ensureWinnersHeader(winSheet);
  Logger.log('Sheets initialised: ' + SHEETS.REGISTRATIONS + ', ' + SHEETS.WINNERS);
}
