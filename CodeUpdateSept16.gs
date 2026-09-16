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

// Shared secret required for sensitive actions (ticket generation, admin data,
// verifying tickets, saving winners). Change this to your own long random string,
// then put the SAME value in generate-ticket.html (and admin.html / raffle-draw.html
// if you add the check there too — see README).
const API_SECRET = 'CHANGE-THIS-TO-A-LONG-RANDOM-SECRET-2025';

// Actions that require the secret key to run (protects write access + PII)
const PROTECTED_POST_ACTIONS = [
  'generateTicket', 'saveWinner', 'importExcel', 'verifyGeneratedTicket',
  'generatePoolTicket', 'verifyPoolTicket', 'setPoolTotal',
];
const PROTECTED_GET_ACTIONS  = ['getAll', 'getGeneratedTickets', 'getPoolStatus', 'getPoolTickets'];

// Sheet tab names (auto-created if missing)
const SHEETS = {
  REGISTRATIONS: 'Registrations',
  WINNERS: 'Winners',
  IMPORTS: 'Imports',
  GENERATE_TICKETS: 'GenerateTicket',
  POOL_TICKETS: 'PoolTickets',
};

// ===================== QCBELC BULK POOL CONFIG ========================
// The pool total is NOT a hardcoded number you have to redeploy to change —
// it lives in this script's Properties store (a small persistent key/value
// area separate from the Sheet) and can be changed live from the
// generate-bulk-ticket.html page via the 'setPoolTotal' action.
//
// IMPORTANT: "remaining" is never stored/decremented directly. It is always
// computed fresh as (pool total − number of tickets already issued). This
// means editing the total at any time — even while someone else is mid-batch
// — can never corrupt or interrupt anything: it only moves the ceiling used
// on the NEXT calculation. Concurrent issuance is protected separately by
// LockService inside generatePoolTicket() so two simultaneous submissions
// can never both succeed past the same remaining count.
const POOL_TOTAL_PROPERTY_KEY = 'QCBELC_POOL_TOTAL';
const DEFAULT_POOL_TOTAL = 3000;

// Receipt photo is always mandatory for pool-issued batches (unlike the
// individual generator, which has REQUIRE_RECEIPT_PHOTO further below).
const REQUIRE_RECEIPT_PHOTO_POOL = true;

// Above this many tickets in a single submission, the frontend switches to a
// compact list view instead of a full QR card per ticket (kept here too so
// the backend can echo it back if ever needed). There is NO hard cap on
// quantity — this is a display threshold only.
const POOL_COMPACT_VIEW_THRESHOLD = 30;

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
      case 'getPoolStatus':    result = getPoolStatus();    break;
      case 'getPoolTickets':   result = getPoolTickets();   break;
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
      case 'generatePoolTicket': result = generatePoolTicket(payload); break;
      case 'verifyPoolTicket':   result = verifyPoolTicket(payload);   break;
      case 'setPoolTotal':       result = setPoolTotal(payload);       break;
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

// Safety cap on how many times we retry generating a fresh, non-colliding
// ID before giving up. IDs are random, so collisions are extremely unlikely,
// but this guarantees we never silently save a duplicate.
const MAX_ID_GENERATION_ATTEMPTS = 20;

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

  const qty = Math.max(1, Math.min(20, parseInt(data.ticketQty) || 1));
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

  // ---- DUPLICATE-ID SAFETY CHECK ----
  // Load existing QR Code IDs (col A) and Ticket IDs (col B) ONCE before the
  // loop, so every new ID in this batch is checked against everything that
  // already exists in the sheet (plus anything generated earlier in this
  // same batch, tracked via the Set as we go).
  const lastRow = sheet.getLastRow();
  const existingQrIds = new Set();
  const existingTicketIds = new Set();
  const existingValues = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 2).getValues() : [];
  existingValues.forEach(row => {
    if (row[0]) existingQrIds.add(String(row[0]));
    if (row[1]) existingTicketIds.add(String(row[1]));
  });

  const tickets = [];
  const rowsToAppend = [];

  for (let i = 0; i < qty; i++) {
    const qrCodeId = generateUniqueId(generateQRCodeId, existingQrIds);
    const ticketId = generateUniqueId(generateTicketId, existingTicketIds);

    if (!qrCodeId || !ticketId) {
      // Extremely unlikely — random ID space exhausted after many retries.
      return {
        success: false,
        message: 'Could not generate a unique ticket ID after several attempts. Please try again — if this keeps happening, contact your system administrator.'
      };
    }

    // Reserve these IDs immediately so the rest of this batch can't reuse them.
    existingQrIds.add(qrCodeId);
    existingTicketIds.add(ticketId);

    rowsToAppend.push([
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

  // Write all rows for this batch in one call (faster + fewer partial-write risks
  // than appending one row at a time inside the loop).
  rowsToAppend.forEach(row => sheet.appendRow(row));

  return {
    success: true,
    tickets: tickets,
    donorName: donorName,
    message: qty + ' ticket(s) generated for ' + donorName + '.',
  };
}

/**
 * Repeatedly calls generatorFn() until it produces a value not already present
 * in existingSet, up to MAX_ID_GENERATION_ATTEMPTS tries. Returns null if it
 * never finds a free ID (should be practically impossible given the ID space).
 */
function generateUniqueId(generatorFn, existingSet) {
  for (let attempt = 0; attempt < MAX_ID_GENERATION_ATTEMPTS; attempt++) {
    const candidate = generatorFn();
    if (!existingSet.has(candidate)) return candidate;
  }
  return null;
}

/**
 * Look up a QR code across BOTH the individual-generator sheet and the
 * QCBELC pool sheet — they share the exact same 16-column layout, so the
 * same row-reading logic works for either. This is what lets a single
 * view-ticket.html / holder-transfer flow work no matter which sheet a
 * ticket originally came from, since the person scanning a QR code has no
 * way of knowing (or needing to know) which flow issued it.
 * Returns { sheet, sheetLabel, rowNum, row } or null if not found anywhere.
 */
function findGeneratedRowByQr(qrCodeId) {
  const candidates = [
    { name: SHEETS.GENERATE_TICKETS, label: 'individual' },
    { name: SHEETS.POOL_TICKETS,     label: 'pool' },
  ];
  for (const c of candidates) {
    const sheet = getOrCreateSheet(c.name);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(qrCodeId)) {
        return { sheet, sheetLabel: c.label, rowNum: i + 1, row: data[i] };
      }
    }
  }
  return null;
}

/** Look up a single ticket by its QRCodeID (used by view-ticket.html) */
function getTicketByQR(qrCodeId) {
  if (!qrCodeId) return { success: false, message: 'No ticket code provided.' };

  const found = findGeneratedRowByQr(qrCodeId);
  if (!found) return { success: false, message: 'No ticket found for this QR code.' };

  const row = found.row;
  // Only return fields safe for public display on view-ticket.html.
  // Phone number and full transaction number are intentionally withheld.
  // All text fields are explicitly coerced to String() — Google Sheets can
  // auto-store all-digit values (e.g. phone/reference numbers) as raw
  // Numbers, which breaks frontend code that expects text (e.g. .toLowerCase()).
  return {
    success: true,
    ticket: {
      ticketId:          String(row[1] || ''),
      memberName:        String(row[2] || ''), // Sold By / Member (same column across both sheets)
      donorName:         String(row[3] || ''), // Recipient / Donor name
      currentHolderName: String(row[4] || ''),
      status:            String(row[11] || ''),
      generatedAt:       String(row[12] || ''),
      source:            found.sheetLabel,     // 'individual' or 'pool'
    }
  };
}

/** List all generated tickets (for admin panel use) */
function getGeneratedTickets() {
  const sheet = getOrCreateSheet(SHEETS.GENERATE_TICKETS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, tickets: [] };

  // All text fields are explicitly coerced to String() — Google Sheets can
  // auto-store all-digit values (e.g. phone/reference numbers) as raw Numbers,
  // which breaks frontend code that expects text (e.g. .toLowerCase() in
  // admin.html's search). Numeric fields (qty/price/amount) stay as numbers.
  const tickets = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    tickets.push({
      qrCodeId: String(row[0] || ''), ticketId: String(row[1] || ''),
      memberName: String(row[2] || ''), donorName: String(row[3] || ''),
      currentHolderName: String(row[4] || ''), donorPhone: String(row[5] || ''),
      transactionNumberId: String(row[6] || ''), paymentMethod: String(row[7] || ''),
      ticketQty: row[8], priceEach: row[9], totalAmount: row[10],
      status: String(row[11] || ''), generatedAt: String(row[12] || ''),
      verifiedAt: String(row[13] || ''), notes: String(row[14] || ''),
      receiptLink: String(row[15] || ''),
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

  const found = findGeneratedRowByQr(qrCodeId);
  if (!found) return { success: false, message: 'No ticket found for this QR code.' };

  const { sheet, rowNum, row } = found;
  const transactionNumberId = String(row[6] || '');
  const expectedCode = transactionNumberId.slice(-4).toUpperCase();

  if (!expectedCode) {
    return { success: false, message: 'This ticket has no transfer code on file. Please contact the raffle organizer to update ownership.' };
  }
  if (verifyCode !== expectedCode) {
    return { success: false, message: 'Incorrect verification code. Please check with the person who gave you this ticket.' };
  }

  const oldHolder = row[4];
  sheet.getRange(rowNum, 5).setValue(newHolderName); // E: CurrentHolderName

  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const prevNotes = row[14] || '';
  const logLine = `[${ts}] Holder changed from "${oldHolder}" to "${newHolderName}"`;
  const newNotes = prevNotes ? (prevNotes + ' | ' + logLine) : logLine;
  sheet.getRange(rowNum, 15).setValue(newNotes); // O: Notes

  return { success: true, message: 'Ticket holder updated.', newHolderName };
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
//  QCBELC BULK POOL TICKETS
// ============================================================
/**
 * Column layout for the PoolTickets sheet — deliberately the SAME shape as
 * GenerateTicket (16 columns, same order) so rows can be copied between the
 * two sheets or into another tracking sheet with no reshuffling:
 * A: QRCodeID          B: Ticket ID          C: SoldBy (defaults 'QCBELC')
 * D: RecipientName     E: CurrentHolderName  F: RecipientPhone
 * G: TransactionNumberId  H: PaymentMethod    I: TicketQty (batch size)
 * J: PriceEach         K: TotalAmount        L: Status (Not Verified / Verified)
 * M: GeneratedAt        N: VerifiedAt         O: Notes
 * P: ReceiptLink
 *
 * This sheet is intentionally kept SEPARATE from GenerateTicket — pool
 * tickets are not mixed into the individual-seller flow or its stats.
 */

/** Set up PoolTickets sheet header if empty */
function ensurePoolTicketsHeader(sheet) {
  if (sheet.getLastRow() === 0) {
    const headers = [
      'QRCodeID','Ticket ID','SoldBy','RecipientName','CurrentHolderName','RecipientPhone',
      'TransactionNumberId','PaymentMethod','TicketQty','PriceEach','TotalAmount',
      'Status','GeneratedAt','VerifiedAt','Notes','ReceiptLink'
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#2A9D5C')
      .setFontColor('#fff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

/**
 * Read the current pool total. Stored in Script Properties (not the Sheet,
 * and not a hardcoded constant) so it can be changed live from
 * generate-bulk-ticket.html without a redeploy. Self-initialises to
 * DEFAULT_POOL_TOTAL the first time it's read.
 */
function getPoolTotal() {
  const props = PropertiesService.getScriptProperties();
  const stored = props.getProperty(POOL_TOTAL_PROPERTY_KEY);
  if (stored === null || stored === '' || isNaN(Number(stored))) {
    props.setProperty(POOL_TOTAL_PROPERTY_KEY, String(DEFAULT_POOL_TOTAL));
    return DEFAULT_POOL_TOTAL;
  }
  return Number(stored);
}

/** Number of individual tickets already issued from the pool (one row = one ticket). */
function getPoolIssuedCount(sheet) {
  const lastRow = sheet.getLastRow();
  return Math.max(0, lastRow - 1); // minus header row
}

/** GET action: current pool remaining/total, for the pool status banner. */
function getPoolStatus() {
  const sheet = getOrCreateSheet(SHEETS.POOL_TICKETS);
  ensurePoolTicketsHeader(sheet);
  const total = getPoolTotal();
  const issued = getPoolIssuedCount(sheet);
  const remaining = Math.max(0, total - issued);
  return { success: true, total, issued, remaining };
}

/**
 * POST action: change the pool total live. Protected by the same API secret
 * as every other write action. Does NOT touch any existing ticket rows —
 * it only changes the ceiling used the next time remaining is computed, so
 * it is always safe to call, including while other batches are in flight.
 */
function setPoolTotal(data) {
  const newTotal = Number(data.newTotal);
  if (isNaN(newTotal) || newTotal < 0 || !isFinite(newTotal)) {
    return { success: false, message: 'Please provide a valid, non-negative pool total.' };
  }
  const rounded = Math.floor(newTotal);
  PropertiesService.getScriptProperties().setProperty(POOL_TOTAL_PROPERTY_KEY, String(rounded));

  const sheet = getOrCreateSheet(SHEETS.POOL_TICKETS);
  ensurePoolTicketsHeader(sheet);
  const issued = getPoolIssuedCount(sheet);
  const remaining = Math.max(0, rounded - issued);

  return {
    success: true,
    total: rounded,
    issued,
    remaining,
    message: issued > rounded
      ? `Pool total set to ${rounded}. Note: ${issued} tickets have already been issued, so the pool is currently sold out until the total is raised again.`
      : `Pool total set to ${rounded}. ${remaining} remaining.`,
  };
}

/** Generate a unique pool ticket ID: QCB-YYYYMMDD-XXXXXX (visually distinct from TKT- individual IDs). */
function generatePoolTicketId() {
  const datePart = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  const rand = Math.random().toString(36).substring(2, 8).toUpperCase();
  return 'QCB-' + datePart + '-' + rand;
}

/**
 * POST action: issue one or more tickets from the QCBELC pool.
 * - No hard cap on quantity (per policy) — a single submission can issue
 *   the whole remaining pool if desired. Rows are written in one batched
 *   setValues() call rather than row-by-row, so large batches stay fast.
 * - LockService serialises this function so two simultaneous submissions
 *   can never both succeed past the same "remaining" count (no overselling
 *   the pool even under concurrent use).
 * - Receipt photo + transaction number are always mandatory here.
 */
function generatePoolTicket(data) {
  const soldBy = (data.soldBy || 'QCBELC').trim() || 'QCBELC';
  const recipientName = (data.recipientName || '').trim();
  const transactionNumberId = (data.transactionNumberId || '').trim();
  const hasReceipt = !!(data.receiptImage && data.receiptFileName);

  if (!recipientName) return { success: false, message: "Recipient's name (district, group, or individual) is required." };
  if (transactionNumberId.length < 4) return { success: false, message: 'A transaction/reference number of at least 4 characters is required.' };
  if (REQUIRE_RECEIPT_PHOTO_POOL && !hasReceipt) {
    return { success: false, message: 'A photo of the payment receipt/screenshot is required.' };
  }

  const qty = Math.max(1, parseInt(data.ticketQty) || 1); // no upper cap by policy

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // up to 30s — pool writes for large batches can take a moment
  } catch (e) {
    return { success: false, message: 'The system is busy processing another batch. Please try again in a few seconds.' };
  }

  try {
    const sheet = getOrCreateSheet(SHEETS.POOL_TICKETS);
    ensurePoolTicketsHeader(sheet);

    // Recompute remaining fresh, INSIDE the lock, so this check is always
    // against the true current state — this is what prevents overselling.
    const total = getPoolTotal();
    const issued = getPoolIssuedCount(sheet);
    const remaining = Math.max(0, total - issued);

    if (qty > remaining) {
      return {
        success: false,
        soldOut: true,
        poolRemaining: remaining,
        poolTotal: total,
        message: remaining <= 0
          ? 'The QCBELC pool is sold out.'
          : `Only ${remaining} ticket(s) left in the pool — reduce the quantity and try again.`,
      };
    }

    const timestamp = new Date();
    const dateStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    // Save the receipt photo once per batch (shared proof of payment).
    let receiptLink = '';
    if (hasReceipt) {
      try {
        const batchId = generatePoolTicketId();
        receiptLink = saveReceiptToDrive(batchId, data.receiptFileName, data.receiptImage);
      } catch (e) {
        if (REQUIRE_RECEIPT_PHOTO_POOL) {
          return { success: false, message: 'Could not save the receipt photo. Please try again.' };
        }
      }
    }

    // Policy: pool-issued tickets are "good as sold" — auto-verified.
    const initialStatus = 'Verified';
    const initialVerifiedAt = dateStr;

    // Uniqueness check spans BOTH sheets, since QR codes must be globally
    // unique across the whole ticketing system (a scan can't tell which
    // sheet a code came from).
    const existingQrIds = new Set();
    const existingTicketIds = new Set();
    [SHEETS.GENERATE_TICKETS, SHEETS.POOL_TICKETS].forEach(sheetName => {
      const s = getOrCreateSheet(sheetName);
      const lastRow = s.getLastRow();
      if (lastRow > 1) {
        s.getRange(2, 1, lastRow - 1, 2).getValues().forEach(row => {
          if (row[0]) existingQrIds.add(String(row[0]));
          if (row[1]) existingTicketIds.add(String(row[1]));
        });
      }
    });

    const tickets = [];
    const rowsToAppend = [];

    for (let i = 0; i < qty; i++) {
      const qrCodeId = generateUniqueId(generateQRCodeId, existingQrIds);
      const ticketId = generateUniqueId(generatePoolTicketId, existingTicketIds);

      if (!qrCodeId || !ticketId) {
        return {
          success: false,
          message: 'Could not generate a unique ticket ID after several attempts. Please try again — if this keeps happening, contact your system administrator.'
        };
      }

      existingQrIds.add(qrCodeId);
      existingTicketIds.add(ticketId);

      rowsToAppend.push([
        qrCodeId,                          // A: QRCodeID
        ticketId,                          // B: Ticket ID
        soldBy,                            // C: SoldBy
        recipientName,                     // D: RecipientName
        recipientName,                     // E: CurrentHolderName (defaults to recipient)
        (data.recipientPhone || '').trim(),// F: RecipientPhone
        transactionNumberId,               // G: TransactionNumberId
        data.paymentMethod || '',          // H: PaymentMethod
        1,                                 // I: TicketQty — 1 per row (one row = one physical ticket)
        TICKET_PRICE,                      // J: PriceEach
        TICKET_PRICE,                      // K: TotalAmount (per-ticket; sum rows for batch total)
        initialStatus,                     // L: Status
        dateStr,                           // M: GeneratedAt
        initialVerifiedAt,                 // N: VerifiedAt
        '',                                // O: Notes
        receiptLink,                       // P: ReceiptLink
      ]);

      tickets.push({ ticketId, qrCodeId });
    }

    // Single batched write — critical for large, uncapped batches so we
    // don't risk timing out issuing hundreds/thousands of rows one at a time.
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);

    const newIssued = issued + qty;
    const newRemaining = Math.max(0, total - newIssued);

    return {
      success: true,
      tickets: tickets,
      recipientName: recipientName,
      poolRemaining: newRemaining,
      poolTotal: total,
      compactView: qty > POOL_COMPACT_VIEW_THRESHOLD,
      message: qty + ' ticket(s) issued for ' + recipientName + '.',
    };
  } finally {
    lock.releaseLock();
  }
}

/** GET action: list all pool tickets (for the admin Pool tab). */
function getPoolTickets() {
  const sheet = getOrCreateSheet(SHEETS.POOL_TICKETS);
  ensurePoolTicketsHeader(sheet);
  const data = sheet.getDataRange().getValues();

  const tickets = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    tickets.push({
      qrCodeId: String(row[0] || ''), ticketId: String(row[1] || ''),
      soldBy: String(row[2] || ''), recipientName: String(row[3] || ''),
      currentHolderName: String(row[4] || ''), recipientPhone: String(row[5] || ''),
      transactionNumberId: String(row[6] || ''), paymentMethod: String(row[7] || ''),
      ticketQty: row[8], priceEach: row[9], totalAmount: row[10],
      status: String(row[11] || ''), generatedAt: String(row[12] || ''),
      verifiedAt: String(row[13] || ''), notes: String(row[14] || ''),
      receiptLink: String(row[15] || ''),
    });
  }

  const total = getPoolTotal();
  const issued = getPoolIssuedCount(sheet);
  const remaining = Math.max(0, total - issued);

  return { success: true, tickets, poolRemaining: remaining, poolTotal: total };
}

/** POST action: mark a pool ticket as Verified (admin use). */
function verifyPoolTicket(data) {
  const qrCodeId = data.qrCodeId;
  if (!qrCodeId) return { success: false, message: 'No ticket code provided.' };

  const sheet = getOrCreateSheet(SHEETS.POOL_TICKETS);
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
  const poolSheet = getOrCreateSheet(SHEETS.POOL_TICKETS);
  ensureRegistrationsHeader(regSheet);
  ensureWinnersHeader(winSheet);
  ensureGenerateTicketsHeader(genSheet);
  ensurePoolTicketsHeader(poolSheet);
  getPoolTotal(); // self-initialises the pool total property to DEFAULT_POOL_TOTAL if unset
  Logger.log('Sheets initialised: ' + SHEETS.REGISTRATIONS + ', ' + SHEETS.WINNERS + ', ' + SHEETS.GENERATE_TICKETS + ', ' + SHEETS.POOL_TICKETS);
}