
/************************************************************
 * UNDER 25 REGISTRATION SYSTEM
 * VERSION: 5.2
 * YEAR: 2026
 * Google Forms + Google Sheets + Google Apps Script
 *
 * FEATURES
 * - Sequential registration IDs
 * - Duplicate email, phone and USN detection
 * - QR code generation
 * - Participant confirmation and admin notification
 * - Email delivery status tracking
 * - Retry failed or pending emails
 * - Input validation
 * - Duplicate-processing protection
 * - Script locking for concurrent submissions
 *
 * COLUMNS
 * A Timestamp
 * B Name
 * C Email
 * D Phone Number
 * E USN
 * F Registration ID
 * G QR Code URL
 * H Status
 * I Participant Email
 * J Admin Email
 ************************************************************/


// ==========================================================
// CONFIGURATION
// ==========================================================

const SHEET_NAME = "Form responses";
const ADMIN_EMAIL = "rahilkalyanker99@gmail.com";

const EVENT_PREFIX = "U25";
const EVENT_YEAR = "2026";

const QR_SIZE = 300;
const LOCK_TIMEOUT = 30000;

const HEADERS = [
  "Timestamp",
  "Name",
  "Email",
  "Phone Number",
  "USN",
  "Registration ID",
  "QR Code",
  "Status",
  "Participant Email",
  "Admin Email"
];

const STATUS = {
  REGISTERED: "🟢 REGISTERED",

  DUPLICATE_EMAIL: "🔴 DUPLICATE - EMAIL",
  DUPLICATE_PHONE: "🔴 DUPLICATE - PHONE",
  DUPLICATE_USN: "🔴 DUPLICATE - USN",

  INVALID_NAME: "🟠 INVALID - NAME",
  INVALID_EMAIL: "🟠 INVALID - EMAIL",
  INVALID_PHONE: "🟠 INVALID - PHONE",
  INVALID_USN: "🟠 INVALID - USN",

  PROCESSING_ERROR: "🔴 PROCESSING ERROR"
};


// ==========================================================
// INSTALL TRIGGER
// Run manually once.
// ==========================================================

function setupTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error("Open the registration spreadsheet first.");
  }

  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error("Sheet not found: " + SHEET_NAME);
  }

  ensureHeaders_(sheet);

  const existingTriggers = ScriptApp.getProjectTriggers();

  existingTriggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "onFormSubmit") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  Logger.log("Version 5.2 trigger installed.");
}


// ==========================================================
// FORM SUBMIT HANDLER
// ==========================================================

function onFormSubmit(e) {
  if (!e || !e.range) {
    throw new Error(
      "Run this function only through the form-submit trigger."
    );
  }

  const sheet = e.range.getSheet();

  if (sheet.getName() !== SHEET_NAME) {
    return;
  }

  const row = e.range.getRow();

  if (row < 2) {
    return;
  }

  ensureHeaders_(sheet);

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(LOCK_TIMEOUT)) {
    throw new Error(
      "Registration lock timed out. Check the row and retry later."
    );
  }

  try {
    processNewSubmission_(sheet, row);
  } finally {
    lock.releaseLock();
  }
}


// ==========================================================
// PROCESS NEW SUBMISSION
// Lock is held during registration and email processing.
// This prevents a retry from overlapping this execution.
// ==========================================================

function processNewSubmission_(sheet, row) {
  const values = sheet.getRange(row, 1, 1, 10).getValues()[0];

  // Existing status or ID means the row has already been evaluated.
  // Use retryPendingEmails() for email retries.
  if (
    String(values[7] || "").trim() ||
    String(values[5] || "").trim()
  ) {
    Logger.log("Already evaluated; skipped row " + row);
    return;
  }

  const registration = {
    name: String(values[1] || "").trim(),
    email: normalizeEmail_(values[2]),
    phone: normalizePhone_(values[3]),
    usn: normalizeUsn_(values[4]),
    row: row
  };

  try {
    const invalid = validateRegistration_(
      registration.name,
      registration.email,
      registration.phone,
      registration.usn
    );

    if (invalid) {
      sheet.getRange(row, 8, 1, 3).setValues([[
        invalid,
        "NOT SENT",
        "NOT SENT"
      ]]);
      return;
    }

    const duplicate = findDuplicate_(
      sheet,
      row,
      registration.email,
      registration.phone,
      registration.usn
    );

    if (duplicate) {
      registration.duplicateStatus = duplicate;

      sheet.getRange(row, 8, 1, 3).setValues([[
        duplicate,
        "PENDING",
        "NOT REQUIRED"
      ]]);

      sendDuplicateNotification_(sheet, registration);
      return;
    }

    registration.registrationId =
      generateRegistrationId_(sheet);

    registration.qrUrl =
      makeQrUrl_(registration.registrationId);

    // Write ID and status before sending any email.
    // This prevents a repeated trigger from creating another ID.
    sheet.getRange(row, 6, 1, 5).setValues([[
      registration.registrationId,
      registration.qrUrl,
      STATUS.REGISTERED,
      "PENDING",
      "PENDING"
    ]]);

    SpreadsheetApp.flush();

    processRegistrationEmails_(sheet, registration);

  } catch (error) {
    Logger.log(
      "Processing error on row " + row + ": " + error
    );

    const current = sheet.getRange(row, 6, 1, 5).getValues()[0];

    // Do not overwrite an established registration.
    if (!current[0] && !current[2]) {
      sheet.getRange(row, 8).setValue(
        STATUS.PROCESSING_ERROR
      );

      sheet.getRange(row, 9, 1, 2).setValues([[
        "NOT SENT",
        "NOT SENT"
      ]]);
    }

    throw error;
  }
}


// ==========================================================
// HEADERS
// Fill blank header cells only.
// Existing header names are not overwritten.
// ==========================================================

function ensureHeaders_(sheet) {
  const range = sheet.getRange(1, 1, 1, HEADERS.length);
  const existing = range.getValues()[0];
  const updated = existing.slice();

  let changed = false;

  HEADERS.forEach((header, index) => {
    if (String(existing[index] || "").trim() === "") {
      updated[index] = header;
      changed = true;
    }
  });

  if (changed) {
    range.setValues([updated]);
  }
}


// ==========================================================
// VALIDATION
// ==========================================================

function validateRegistration_(name, email, phone, usn) {
  if (!name || name.length < 2 || name.length > 100) {
    return STATUS.INVALID_NAME;
  }

  if (!isValidEmail_(email)) {
    return STATUS.INVALID_EMAIL;
  }

  if (!isValidPhone_(phone)) {
    return STATUS.INVALID_PHONE;
  }

  if (!isValidUsn_(usn)) {
    return STATUS.INVALID_USN;
  }

  return null;
}


function isValidEmail_(email) {
  return Boolean(
    email &&
    email.length <= 254 &&
    /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)
  );
}


function isValidPhone_(phone) {
  const digits = String(phone || "").replace(/\D/g, "");

  let national = digits;

  if (digits.length === 12 && digits.startsWith("91")) {
    national = digits.slice(2);
  }

  return (
    national.length === 10 &&
    /^[6-9]\d{9}$/.test(national)
  );
}


function isValidUsn_(usn) {
  return /^[A-Z0-9-]{5,20}$/.test(usn);
}


// ==========================================================
// NORMALIZATION
// ==========================================================

function normalizeEmail_(value) {
  return String(value || "").trim().toLowerCase();
}


function normalizePhone_(value) {
  let digits = String(value || "").replace(/\D/g, "");

  if (digits.length === 12 && digits.startsWith("91")) {
    digits = digits.slice(2);
  }

  return digits;
}


function normalizeUsn_(value) {
  return String(value || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "");
}


// ==========================================================
// DUPLICATE DETECTION
// Only successfully registered rows count.
// ==========================================================

function findDuplicate_(sheet, currentRow, email, phone, usn) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return null;
  }

  const rows = sheet.getRange(
    2, 1, lastRow - 1, 8
  ).getValues();

  for (let i = 0; i < rows.length; i++) {
    const sheetRow = i + 2;

    if (sheetRow === currentRow) {
      continue;
    }

    const existing = rows[i];

    if (
      String(existing[7] || "").trim() !==
      STATUS.REGISTERED
    ) {
      continue;
    }

    if (normalizeEmail_(existing[2]) === email) {
      return STATUS.DUPLICATE_EMAIL;
    }

    if (normalizePhone_(existing[3]) === phone) {
      return STATUS.DUPLICATE_PHONE;
    }

    if (normalizeUsn_(existing[4]) === usn) {
      return STATUS.DUPLICATE_USN;
    }
  }

  return null;
}


// ==========================================================
// REGISTRATION ID
// Format: U25-2026-0001
// ==========================================================

function generateRegistrationId_(sheet) {
  const lastRow = sheet.getLastRow();
  let highest = 0;

  if (lastRow >= 2) {
    const ids = sheet.getRange(
      2, 6, lastRow - 1, 1
    ).getValues().flat();

    const prefix = EVENT_PREFIX + "-" + EVENT_YEAR + "-";

    ids.forEach(value => {
      const id = String(value || "").trim();

      if (!id.startsWith(prefix)) {
        return;
      }

      const suffix = id.slice(prefix.length);

      if (/^\d+$/.test(suffix)) {
        highest = Math.max(highest, Number(suffix));
      }
    });
  }

  return (
    EVENT_PREFIX + "-" +
    EVENT_YEAR + "-" +
    String(highest + 1).padStart(4, "0")
  );
}


// ==========================================================
// QR URL AND FETCH
// ==========================================================

function makeQrUrl_(registrationId) {
  return (
    "https://quickchart.io/qr?text=" +
    encodeURIComponent(registrationId) +
    "&size=" + QR_SIZE
  );
}


function fetchQrBlob_(url, registrationId) {
  const response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true
  });

  const code = response.getResponseCode();

  if (code !== 200) {
    throw new Error("QR service HTTP " + code);
  }

  const blob = response.getBlob();

  if (
    !blob.getContentType() ||
    !blob.getContentType().toLowerCase().startsWith("image/")
  ) {
    throw new Error("QR service did not return an image.");
  }

  if (!blob.getBytes().length) {
    throw new Error("QR image is empty.");
  }

  return blob.setName(registrationId + ".png");
}


// ==========================================================
// REGISTRATION EMAILS
// ==========================================================

function processRegistrationEmails_(sheet, registration) {
  const row = registration.row;

  const current = sheet.getRange(
    row, 1, 1, 10
  ).getValues()[0];

  registration.registrationId =
    String(current[5] || registration.registrationId || "").trim();

  registration.qrUrl =
    String(current[6] || registration.qrUrl || "").trim();

  const participantStatus = String(current[8] || "").trim();
  const adminStatus = String(current[9] || "").trim();

  let participantSent = participantStatus === "SENT";

  // Send participant email if not already marked SENT.
  if (!participantSent && participantStatus !== "NOT SENT") {
    try {
      const qrBlob = fetchQrBlob_(
        registration.qrUrl,
        registration.registrationId
      );

      sendParticipantEmail_(registration, qrBlob);

      participantSent = true;
      sheet.getRange(row, 9).setValue("SENT");

    } catch (error) {
      Logger.log(
        "Participant email failed on row " + row + ": " + error
      );

      sheet.getRange(row, 9).setValue(
        "FAILED - " + shortError_(error)
      );
    }
  }

  // Admin notification is independent of participant email.
  if (adminStatus !== "SENT" && adminStatus !== "NOT REQUIRED") {
    try {
      sendAdminEmail_(registration, participantSent);
      sheet.getRange(row, 10).setValue("SENT");

    } catch (error) {
      Logger.log(
        "Admin email failed on row " + row + ": " + error
      );

      sheet.getRange(row, 10).setValue(
        "FAILED - " + shortError_(error)
      );
    }
  }
}


// ==========================================================
// PARTICIPANT EMAIL
// ==========================================================

function sendParticipantEmail_(registration, qrBlob) {
  const safeName = escapeHtml_(registration.name);
  const safeId = escapeHtml_(registration.registrationId);

  const htmlBody = `
    <div style="font-family:Arial,sans-serif;line-height:1.6">
      <h2>UNDER 25 REGISTRATION</h2>
      <p>Hello ${safeName},</p>
      <p>Your registration has been completed successfully.</p>
      <p><b>Registration ID:</b> ${safeId}</p>
      <p>Please keep your registration ID and QR code safe.</p>
      <p><b>Your QR Code:</b></p>
      <img src="cid:registrationQR"
           alt="Registration QR Code"
           width="250"
           height="250">
      <p>Thank you for registering with Under 25!</p>
    </div>
  `;

  const textBody =
    "Hello " + registration.name + ",\n\n" +
    "Your Under 25 registration is confirmed.\n\n" +
    "Registration ID: " + registration.registrationId + "\n\n" +
    "Please keep your QR code and registration ID safe.\n\n" +
    "Thank you!";

  MailApp.sendEmail({
    to: registration.email,
    subject: "Under 25 Registration Confirmed",
    body: textBody,
    htmlBody: htmlBody,
    inlineImages: {
      registrationQR: qrBlob
    },
    name: "Under 25"
  });
}


// ==========================================================
// ADMIN EMAIL
// ==========================================================

function sendAdminEmail_(registration, participantEmailSent) {
  const htmlBody = `
    <div style="font-family:Arial,sans-serif;line-height:1.6">
      <h2>New Registration Received</h2>
      <p><b>Name:</b> ${escapeHtml_(registration.name)}</p>
      <p><b>Email:</b> ${escapeHtml_(registration.email)}</p>
      <p><b>Phone:</b> ${escapeHtml_(registration.phone)}</p>
      <p><b>USN:</b> ${escapeHtml_(registration.usn)}</p>
      <p><b>Registration ID:</b>
        ${escapeHtml_(registration.registrationId)}
      </p>
      <p><b>Participant confirmation:</b>
        ${participantEmailSent ? "SENT" : "FAILED"}
      </p>
    </div>
  `;

  const textBody =
    "New Under 25 registration\n\n" +
    "Name: " + registration.name + "\n" +
    "Email: " + registration.email + "\n" +
    "Phone: " + registration.phone + "\n" +
    "USN: " + registration.usn + "\n" +
    "Registration ID: " + registration.registrationId + "\n" +
    "Participant confirmation: " +
    (participantEmailSent ? "SENT" : "FAILED");

  MailApp.sendEmail({
    to: ADMIN_EMAIL,
    subject: "New Under 25 Registration - " +
      registration.registrationId,
    body: textBody,
    htmlBody: htmlBody,
    name: "Under 25 Registration System"
  });
}


// ==========================================================
// DUPLICATE NOTIFICATION
// ==========================================================

function sendDuplicateNotification_(sheet, registration) {
  const row = registration.row;

  const current = String(
    sheet.getRange(row, 9).getValue() || ""
  ).trim();

  if (current === "SENT") {
    return;
  }

  try {
    MailApp.sendEmail({
      to: registration.email,
      subject: "Under 25 Registration - Duplicate Submission",
      body:
        "Hello " + registration.name + ",\n\n" +
        "Our system detected a duplicate registration.\n" +
        "Reason: " + registration.duplicateStatus + "\n\n" +
        "If you believe this is a mistake, please contact the organizers.",
      name: "Under 25"
    });

    sheet.getRange(row, 9).setValue("SENT");

  } catch (error) {
    Logger.log(
      "Duplicate notification failed on row " + row + ": " + error
    );

    sheet.getRange(row, 9).setValue(
      "FAILED - " + shortError_(error)
    );
  }
}


// ==========================================================
// RETRY PENDING OR FAILED EMAILS
// Run manually from Apps Script when required.
// This function shares the same script lock as submissions.
// ==========================================================

function retryPendingEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error("Open the registration spreadsheet first.");
  }

  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error("Sheet not found: " + SHEET_NAME);
  }

  ensureHeaders_(sheet);

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    Logger.log("No registration rows to retry.");
    return;
  }

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(LOCK_TIMEOUT)) {
    throw new Error(
      "Registration system is busy. Try again later."
    );
  }

  try {
    for (let row = 2; row <= lastRow; row++) {
      const values = sheet.getRange(
        row, 1, 1, 10
      ).getValues()[0];

      const status = String(values[7] || "").trim();
      const participantStatus = String(values[8] || "").trim();
      const adminStatus = String(values[9] || "").trim();

      if (
        status === STATUS.REGISTERED &&
        String(values[5] || "").trim()
      ) {
        const participantNeedsRetry =
          participantStatus !== "SENT" &&
          participantStatus !== "NOT SENT";

        const adminNeedsRetry =
          adminStatus !== "SENT" &&
          adminStatus !== "NOT REQUIRED";

        if (!participantNeedsRetry && !adminNeedsRetry) {
          continue;
        }

        const registration = {
          row: row,
          name: String(values[1] || "").trim(),
          email: normalizeEmail_(values[2]),
          phone: normalizePhone_(values[3]),
          usn: normalizeUsn_(values[4]),
          registrationId: String(values[5]).trim(),
          qrUrl: String(values[6] || "").trim()
        };

        processRegistrationEmails_(sheet, registration);

      } else if (
        status.indexOf("DUPLICATE - ") !== -1 &&
        participantStatus !== "SENT"
      ) {
        sendDuplicateNotification_(sheet, {
          row: row,
          name: String(values[1] || "").trim(),
          email: normalizeEmail_(values[2]),
          duplicateStatus: status
        });
      }
    }
  } finally {
    lock.releaseLock();
  }

  Logger.log("Email retry process completed.");
}


// ==========================================================
// ERROR TEXT
// ==========================================================

function shortError_(error) {
  return String(
    error && error.message ? error.message : error
  )
    .replace(/[\r\n]+/g, " ")
    .slice(0, 120);
}


// ==========================================================
// HTML ESCAPING
// ==========================================================

function escapeHtml_(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

// ==========================================================
// UNDER 25 - PHASE 2: QR ATTENDANCE BACKEND
// ==========================================================

const ATTENDANCE_SHEET_NAME = "Attendance";

// Scan a registration ID and record attendance.
// Called by the attendance scanner in the next step.
function recordAttendance(registrationId, checkedInBy) {
  registrationId = String(registrationId || "").trim();
  checkedInBy = String(checkedInBy || "").trim();

  if (!registrationId) {
    return { success: false, message: "Registration ID is required." };
  }

  if (!checkedInBy) {
    return { success: false, message: "Please enter the check-in operator's name." };
  }

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(LOCK_TIMEOUT)) {
    return { success: false, message: "System busy. Please try again." };
  }

  try {
    const ss = SpreadsheetApp.openById("1CYTmgL5AraFSJQIGBlcHSFyA2j8nE8HJc_Y9f7kD2iw");
    const registrationSheet = ss.getSheetByName(SHEET_NAME);
    const attendanceSheet = ss.getSheetByName(ATTENDANCE_SHEET_NAME);

    if (!registrationSheet || !attendanceSheet) {
      return { success: false, message: "Required sheet not found." };
    }

    // Find the participant in the registration sheet.
    const lastRow = registrationSheet.getLastRow();

    if (lastRow < 2) {
      return { success: false, message: "No registrations found." };
    }

    const registrations = registrationSheet
      .getRange(2, 1, lastRow - 1, 8)
      .getValues();

    let participant = null;

    for (const row of registrations) {
      const id = String(row[5] || "").trim();
      const status = String(row[7] || "").trim();

      if (id === registrationId && status === STATUS.REGISTERED) {
        participant = {
          id: id,
          name: String(row[1] || "").trim()
        };
        break;
      }
    }

    if (!participant) {
      return {
        success: false,
        message: "Invalid registration ID or registration not confirmed."
      };
    }

    // Prevent duplicate check-ins.
    const attendanceLastRow = attendanceSheet.getLastRow();

    if (attendanceLastRow >= 2) {
      const existingIds = attendanceSheet
        .getRange(2, 1, attendanceLastRow - 1, 1)
        .getValues()
        .flat()
        .map(value => String(value || "").trim());

      if (existingIds.includes(registrationId)) {
        return {
          success: false,
          message: participant.name + " has already checked in."
        };
      }
    }

    // Record attendance.
    const checkInTime = new Date();

    attendanceSheet.appendRow([
      participant.id,
      participant.name,
      "PRESENT",
      checkInTime,
      checkedInBy
    ]);

    SpreadsheetApp.flush();

    return {
      success: true,
      message: "Attendance recorded successfully.",
      name: participant.name,
      registrationId: participant.id,
      checkInTime: checkInTime.toISOString()
    };

  } catch (error) {
    Logger.log("Attendance error: " + error);
    return {
      success: false,
      message: "Attendance could not be recorded. Please try again."
    };
  } finally {
    lock.releaseLock();
  }
}


function doGet(e) {
  const data = (e && e.parameter) || {};

  // GitHub Pages cannot read a normal cross-origin Apps Script response.
  // JSONP lets the scanner receive the actual attendance result.
  if (data.action === "recordAttendance") {
    const callback = String(data.callback || "");

    if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
      return ContentService
        .createTextOutput("/* Invalid callback */")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    const result = recordAttendance(
      data.registrationId,
      data.checkedInBy
    );

    return ContentService
      .createTextOutput(callback + "(" + JSON.stringify(result) + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return HtmlService
    .createHtmlOutputFromFile("Scanner")
    .setTitle("Under 25 Attendance Scanner");
}


function doPost(e) {
  try {
    Logger.log("doPost started");

    const data = e.parameter || {};
    Logger.log("Received data: " + JSON.stringify(data));

    if (data.action !== "recordAttendance") {
      throw new Error("Invalid action: " + data.action);
    }

    const result = recordAttendance(
      data.registrationId,
      data.checkedInBy
    );

    Logger.log("recordAttendance result: " + JSON.stringify(result));

    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        result: result
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log("ERROR: " + error.message);
    Logger.log(error.stack || "No stack trace");

    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: error.message
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}