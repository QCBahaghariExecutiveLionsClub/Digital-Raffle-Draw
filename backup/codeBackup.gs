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
const SHEET_ID = '1yo-g0xrG-cOSjpU7WM7slBjhAEDSEssw5Sg3cbbrtTg'; // ← Replace this
const TICKET_PRICE = 100; // Price per ticket in PHP

// Shared secret required for sensitive actions (ticket generation, admin data,
// verifying tickets, saving winners). Change this to your own long random string,
// then put the SAME value in generate-ticket.html (and admin.html / raffle-draw.html
// if you add the check there too — see README).
const API_SECRET = 'L10nsClvbQCBELC';

// Actions that require the secret key to run (protects write access + PII)
const PROTECTED_POST_ACTIONS = ['generateTicket', 'saveWinner', 'importExcel', 'verifyGeneratedTicket'];
const PROTECTED_GET_ACTIONS  = ['getAll', 'getGeneratedTickets'];

// Sheet tab names (auto-created if missing)
const SHEETS = {
  REGISTRATIONS: 'Registrations',
  WINNERS: 'Winners',
  IMPORTS: 'Imports',
  GENERATE_TICKETS: 'GenerateTicket',
};

// ============================================================
//  ENTRY POINT: doGet — handles all GET requests
// ============================================================
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'ping';

  if (PROTECTED_GET_ACTIONS.indexOf(action) !== -1) {
    const providedSecret = (e && e.parameter && e.parameter.secret) || '';
    if (providedSecret !== API_SECRET) {
      return jsonResponse({ success: false, message: 'Unauthorized. Missing or invalid access key.' });
    }
  }

  let result;
  try {
    switch (action) {
      case 'getParticipants': result = getParticipants(); break;
      case 'getAll':          result = getAllRecords();    break;
      case 'getWinners':      result = getWinners();      break;
      case 'getStats':        result = getStats();         break;
      case 'getTicketByQR':   result = getTicketByQR(e.parameter.qrCodeId); break;
      case 'getGeneratedTickets': result = getGeneratedTickets(); break;
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

  if (PROTECTED_POST_ACTIONS.indexOf(action) !== -1) {
    if (payload.secret !== API_SECRET) {
      return jsonResponse({ success: false, message: 'Unauthorized. Missing or invalid access key.' });
    }
  }

  let result;

  try {
    switch (action) {
      case 'register':    result = registerParticipant(payload); break;
      case 'saveWinner':  result = saveWinner(payload);          break;
      case 'importExcel': result = importExcelData(payload);     break;
      case 'generateTicket':     result = generateTicket(payload);      break;
      case 'updateTicketHolder': result = updateTicketHolder(payload);  break;
      case 'verifyGeneratedTicket': result = verifyGeneratedTicket(payload); break;
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

    // Save receipt image to Drive first (if provided) so we can store the link on the row
    let receiptLink = '';
    if (data.receiptImage && data.receiptFileName) {
      try { receiptLink = saveReceiptToDrive(ticketId, data.receiptFileName, data.receiptImage); }
      catch(e) { /* Drive saving is optional — don't fail registration */ }
    }

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
      data.referralName || '',             // R: Referral Name
      receiptLink,                         // S: Receipt Link (Google Drive)
    ]);
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
  const file = folder.createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) { /* sharing may already be set by domain policy */ }
  return file.getUrl();
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
      referralName:  row[17] || '',
      receiptLink:   row[18] || '',
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
//  GENERATE TICKET (seller / assistant issued tickets)
// ============================================================
/**
 * Column layout for the GenerateTicket sheet:
 * A: QRCodeID          B: Ticket ID         C: MemberName (Seller/Assistant)
 * D: DonorName         E: CurrentHolderName F: DonorPhone
 * G: TransactionNumberId  H: PaymentMethod   I: TicketQty (batch size)
 * J: PriceEach         K: TotalAmount       L: Status (Not Verified / Verified)
 * M: GeneratedAt        N: VerifiedAt        O: Notes (holder-change / audit log)
 * P: ReceiptLink (Google Drive link to the payment screenshot, if provided)
 */
// Toggle whether a receipt photo is mandatory to generate a ticket.
// Set to false to make the photo optional again (e.g. once the reference-number
// field becomes the primary proof, or vice versa).
const REQUIRE_RECEIPT_PHOTO = false;

function generateTicket(data) {
  const sheet = getOrCreateSheet(SHEETS.GENERATE_TICKETS);
  ensureGenerateTicketsHeader(sheet);

  const memberName = (data.memberName || '').trim();
  const donorName  = (data.donorName || '').trim();
  const transactionNumberId = (data.transactionNumberId || '').trim();
  const hasReceipt = !!(data.receiptImage && data.receiptFileName);

  if (!memberName) return { success: false, message: 'Seller/assistant name is required.' };
  if (!donorName)  return { success: false, message: "Buyer's full name is required." };
  if (REQUIRE_RECEIPT_PHOTO && !hasReceipt) {
    return { success: false, message: 'A photo of the payment receipt/screenshot is required.' };
  }

  const qty = Math.max(1, Math.min(10, parseInt(data.ticketQty) || 1));
  const timestamp = new Date();
  const dateStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  // Save the receipt photo to Drive ONCE per batch (all tickets in this submission
  // share the same proof of payment) rather than re-uploading it per ticket.
  let receiptLink = '';
  if (hasReceipt) {
    try {
      const batchId = generateTicketId(); // reuse as a unique file-naming prefix
      receiptLink = saveReceiptToDrive(batchId, data.receiptFileName, data.receiptImage);
    } catch (e) {
      if (REQUIRE_RECEIPT_PHOTO) {
        return { success: false, message: 'Could not save the receipt photo. Please try again.' };
      }
      // otherwise, non-fatal — continue without a link
    }
  }

  // Policy: tickets generated by a QCBELC member are considered "good as sold" —
  // auto-verified on generation, no separate manual review step required.
  const AUTO_VERIFY_ON_GENERATE = true;
  const initialStatus = AUTO_VERIFY_ON_GENERATE ? 'Verified' : 'Not Verified';
  const initialVerifiedAt = AUTO_VERIFY_ON_GENERATE ? dateStr : '';

  const tickets = [];

  for (let i = 0; i < qty; i++) {
    const qrCodeId = generateQRCodeId();
    const ticketId = generateTicketId();

    sheet.appendRow([
      qrCodeId,                       // A: QRCodeID
      ticketId,                       // B: Ticket ID
      memberName,                     // C: MemberName
      donorName,                      // D: DonorName
      donorName,                      // E: CurrentHolderName (defaults to donor)
      (data.donorPhone || '').trim(), // F: DonorPhone
      transactionNumberId,            // G: TransactionNumberId
      data.paymentMethod || '',       // H: PaymentMethod
      qty,                            // I: TicketQty (batch size)
      TICKET_PRICE,                   // J: PriceEach
      qty * TICKET_PRICE,             // K: TotalAmount
      initialStatus,                  // L: Status
      dateStr,                        // M: GeneratedAt
      initialVerifiedAt,               // N: VerifiedAt
      '',                             // O: Notes
      receiptLink,                    // P: ReceiptLink
    ]);

    tickets.push({ ticketId, qrCodeId });
  }

  return {
    success: true,
    tickets: tickets,
    donorName: donorName,
    message: qty + ' ticket(s) generated for ' + donorName + '.',
  };
}

/** Look up a single ticket by its QRCodeID (used by view-ticket.html) */
function getTicketByQR(qrCodeId) {
  if (!qrCodeId) return { success: false, message: 'No ticket code provided.' };

  const sheet = getOrCreateSheet(SHEETS.GENERATE_TICKETS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] === qrCodeId) {
      // Only return fields safe for public display on view-ticket.html.
      // Phone number and full transaction number are intentionally withheld.
      return {
        success: true,
        ticket: {
          ticketId:          row[1],
          memberName:        row[2],
          donorName:         row[3],
          currentHolderName: row[4],
          status:            row[11],
          generatedAt:       row[12],
        }
      };
    }
  }
  return { success: false, message: 'No ticket found for this QR code.' };
}

/** List all generated tickets (for admin panel use) */
function getGeneratedTickets() {
  const sheet = getOrCreateSheet(SHEETS.GENERATE_TICKETS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, tickets: [] };

  const tickets = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    tickets.push({
      qrCodeId: row[0], ticketId: row[1], memberName: row[2], donorName: row[3],
      currentHolderName: row[4], donorPhone: row[5], transactionNumberId: row[6],
      paymentMethod: row[7], ticketQty: row[8], priceEach: row[9], totalAmount: row[10],
      status: row[11], generatedAt: row[12], verifiedAt: row[13], notes: row[14],
      receiptLink: row[15] || '',
    });
  }
  return { success: true, tickets };
}

/**
 * Update who currently holds a ticket (e.g. it was resold/given away).
 * The original DonorName (column D) is preserved for audit purposes;
 * only CurrentHolderName (column E) changes, with a note logged.
 *
 * To prevent anyone who merely sees/screenshots a QR code from hijacking
 * ownership, the caller must also supply the last 4 characters of the
 * ticket's TransactionNumberId — known only to the seller and buyer.
 */
function updateTicketHolder(data) {
  const qrCodeId = data.qrCodeId;
  const newHolderName = (data.newHolderName || '').trim();
  const verifyCode = (data.verifyCode || '').trim().toUpperCase();

  if (!qrCodeId) return { success: false, message: 'No ticket code provided.' };
  if (newHolderName.length < 2) return { success: false, message: 'Please enter a valid full name.' };
  if (!verifyCode) return { success: false, message: 'Please enter the 4-digit verification code from your seller.' };

  const sheet = getOrCreateSheet(SHEETS.GENERATE_TICKETS);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === qrCodeId) {
      const transactionNumberId = String(values[i][6] || '');
      const expectedCode = transactionNumberId.slice(-4).toUpperCase();

      if (!expectedCode) {
        return { success: false, message: 'This ticket has no transfer code on file. Please contact the raffle organizer to update ownership.' };
      }
      if (verifyCode !== expectedCode) {
        return { success: false, message: 'Incorrect verification code. Please check with the person who gave you this ticket.' };
      }

      const oldHolder = values[i][4];
      const rowNum = i + 1;
      sheet.getRange(rowNum, 5).setValue(newHolderName); // E: CurrentHolderName

      const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      const prevNotes = values[i][14] || '';
      const logLine = `[${ts}] Holder changed from "${oldHolder}" to "${newHolderName}"`;
      const newNotes = prevNotes ? (prevNotes + ' | ' + logLine) : logLine;
      sheet.getRange(rowNum, 15).setValue(newNotes); // O: Notes

      return { success: true, message: 'Ticket holder updated.', newHolderName };
    }
  }
  return { success: false, message: 'No ticket found for this QR code.' };
}

/** Mark a generated ticket as Verified (for admin use once payment is confirmed) */
function verifyGeneratedTicket(data) {
  const qrCodeId = data.qrCodeId;
  if (!qrCodeId) return { success: false, message: 'No ticket code provided.' };

  const sheet = getOrCreateSheet(SHEETS.GENERATE_TICKETS);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === qrCodeId) {
      const rowNum = i + 1;
      const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      sheet.getRange(rowNum, 12).setValue('Verified'); // L: Status
      sheet.getRange(rowNum, 14).setValue(ts);          // N: VerifiedAt
      return { success: true, message: 'Ticket marked as Verified.' };
    }
  }
  return { success: false, message: 'No ticket found for this QR code.' };
}

/** Generate a unique QR code identifier: QR-XXXXXXXX */
function generateQRCodeId() {
  const rand = Utilities.getUuid().replace(/-/g, '').substring(0, 10).toUpperCase();
  return 'QR-' + rand;
}

/** Set up GenerateTicket sheet header if empty */
function ensureGenerateTicketsHeader(sheet) {
  if (sheet.getLastRow() === 0) {
    const headers = [
      'QRCodeID','Ticket ID','MemberName','DonorName','CurrentHolderName','DonorPhone',
      'TransactionNumberId','PaymentMethod','TicketQty','PriceEach','TotalAmount',
      'Status','GeneratedAt','VerifiedAt','Notes','ReceiptLink'
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#E85D04')
      .setFontColor('#fff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
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
      'Receipt File','Has Receipt','Registered At','Status','Winner Round','Notes','Referral Name','Receipt Link'
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
  const genSheet = getOrCreateSheet(SHEETS.GENERATE_TICKETS);
  ensureRegistrationsHeader(regSheet);
  ensureWinnersHeader(winSheet);
  ensureGenerateTicketsHeader(genSheet);
  Logger.log('Sheets initialised: ' + SHEETS.REGISTRATIONS + ', ' + SHEETS.WINNERS + ', ' + SHEETS.GENERATE_TICKETS);
}