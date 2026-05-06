
const SECURITY_PEPPER = "v9_PEF_SYS_2026_!@#"; 

/**
 * Modern 2026 Secure Hash
 * Uses HmacSha256 + Salt (Email) + Pepper (Secret)
 */
function generateSHA256(input, salt) {
  if (!input || !salt) throw new Error("Security Error: Missing Hash Inputs.");
  const secret = salt.trim().toLowerCase() + SECURITY_PEPPER;
  const signature = Utilities.computeHmacSha256Signature(input.trim(), secret);
  return signature.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

/**
 * Mandatory Helper: Santize strings to prevent XSS and Template Injection
 */
function safeValue(text) {
  if (typeof text !== 'string') return text || "";
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

/**
 * UI INITIALIZATION
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Payment Endorsement Platform")
    // Use SAMEORIGIN to prevent malicious iframing (2026 standard)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT) //Set to Default Always for Security (requires user authentication, or performs actions in a user's Google account.)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

const USER_DB_ID = "1dBO8ThI7FEKb24D9sPVWokfXLuWUx5aCQvisrT9wBvI";
const USER_TAB = "PEP";

/**
 * ==========================================
 * REGISTRATION SYSTEM (ZERO-TRUST)
 * ==========================================
 */

function getActiveUserEmail() {
  const activeEmail = Session.getActiveUser().getEmail();
  const effectiveEmail = Session.getEffectiveUser().getEmail();

  // LOGGING: Check your Apps Script Logs to see what is happening
  console.log("Active User: " + activeEmail);
  console.log("Effective User: " + effectiveEmail);

  const finalEmail = activeEmail || effectiveEmail;

  if (!finalEmail) {
    throw new Error(
      "Google Identity not found. Please ensure you are logged into your browser with your company email."
    );
  }

  return finalEmail;
}

/**
 * Stage 1: Security Check & OTP
 * Logic: Checks if account already exists. Only sends code if new user.
 */
function sendVerificationCode(email) {
  const sessionEmail = Session.getActiveUser().getEmail();
  if (!email || email !== sessionEmail) throw new Error("Security Violation: Identity mismatch.");

  const ss = SpreadsheetApp.openById(USER_DB_ID);
  const sheet = ss.getSheetByName(USER_TAB);
  const data = sheet.getDataRange().getValues();

  // Index 2 is USERNAME/EMAIL (Column C)
  const alreadyExists = data.some((row) => row[2] === email);
  if (alreadyExists) throw new Error("This email is already registered. Please proceed to Login.");

  const otp = Math.floor(100000 + Math.random() * 900000).toString();
  const cache = CacheService.getScriptCache();
  cache.put(email, otp, 600); // 10 min expiry

  try {
    MailApp.sendEmail(email, "Account Verification Code", "Your verification code is: " + otp);
    return { success: true };
  } catch (e) {
    throw new Error("SMTP Error: Failed to deliver verification email.");
  }
}

/**
 * Stage 2: OTP Validation
 */
function verifyRegistrationCode(email, userCode) {
  const sessionEmail = Session.getActiveUser().getEmail();
  if (email !== sessionEmail) throw new Error("Security Violation: Session hijacked.");

  const cache = CacheService.getScriptCache();
  const storedCode = cache.get(email);

  if (!storedCode) throw new Error("Verification code expired.");
  if (storedCode !== userCode) throw new Error("Incorrect verification code.");

  return { success: true };
}

/**
 * Stage 3: Write Record with Gap-Filling
 * AUTOMATIC VALUES: USERNAME = Email, ACCESS LEVEL = REQUESTOR, STATUS = PENDING
 */

function finalizeRegistration(formData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const activeUser = Session.getActiveUser().getEmail(); // Server-side truth
    const ss = SpreadsheetApp.openById(USER_DB_ID);
    let sheet = ss.getSheetByName(USER_TAB);

    const timestamp = new Date();
    const organicStatus = activeUser.toLowerCase().includes("@megaworld-lifestyle.com") ? "ORGANIC" : "NON ORGANIC";

    // SECURITY UPDATE: Passing activeUser as the SALT
    const encryptedPassword = generateSHA256(formData.password, activeUser);

    // [Maintain your existing gap-filling logic...]
    const nameCol = sheet.getRange("B:B").getValues();
    let targetRow = -1;
    for (let i = 1; i < nameCol.length; i++) {
      if (nameCol[i][0] === "" || nameCol[i][0] === null) {
        targetRow = i + 1;
        break;
      }
    }
    if (targetRow === -1) targetRow = sheet.getLastRow() + 1;

    const rowPayload = [
      timestamp, safeValue(formData.fullName), activeUser, encryptedPassword,
      "REQUESTOR", organicStatus, "Pending", "INACTIVE",
    ];

    sheet.getRange(targetRow, 1, 1, rowPayload.length).setValues([rowPayload]);
    return { success: true, message: "Registered under " + activeUser };
  } finally {
    lock.releaseLock();
  }
}

/**
 * SHA-256 Hashing Helper Function
 * Ensures registration and login encryption methods match.
 */
function generateSHA256(input) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input);
  let txtHash = "";
  for (let i = 0; i < rawHash.length; i++) {
    let hashVal = rawHash[i];
    if (hashVal < 0) hashVal += 256; 
    if (hashVal.toString(16).length === 1) txtHash += "0";
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}


/**
 * Main authentication function
 */
function authenticateUser(credentials) {
  SpreadsheetApp.flush();
  const ss = SpreadsheetApp.openById(USER_DB_ID);
  const sheet = ss.getSheetByName(USER_TAB);
  const data = sheet.getDataRange().getDisplayValues();
  data.shift(); 

  const inputEmail = String(credentials.email || "").trim().toLowerCase();
  // SALT comparison must use the specific inputEmail
  const hashedInputPass = generateSHA256(credentials.password, inputEmail); 

  const userRow = data.find((row) => {
    return String(row[2]).trim().toLowerCase() === inputEmail && String(row[3]).trim() === hashedInputPass;
  });

  if (!userRow) throw new Error("Invalid credentials.");

  const permission = String(userRow[6]).toUpperCase();
  const status = String(userRow[7]).toUpperCase();

  if (permission !== "APPROVED") throw new Error("Access Denied: Pending admin approval.");
  if (status === "RESIGNED") throw new Error("Access Denied: Account deactivated.");

  return {
    success: true,
    user: { fullName: userRow[1], email: userRow[2], role: userRow[4] }
  };
}



/**
 * ==========================================
 * RFP TRANSACTION & GENERATION (Spreadsheet)
 * ==========================================
 */
var REGISTRY_SS_ID = "1YAvZmCdWXbjOcJA-uUY40e6qVqzyiHcB06NpiPcz6y4";
var TAB_NAME = "autogenrfp";

function generateRfpNumber() {
  var lock = LockService.getPublicLock();
  try {
    // 1. Critical Section: Prevent concurrent executions
    lock.waitLock(30000);

    var ss = SpreadsheetApp.openById(REGISTRY_SS_ID);
    var sheet = ss.getSheetByName(TAB_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(TAB_NAME);
      sheet.appendRow(["RFP Number", "Timestamp"]);
    }

    var data = sheet.getRange("A:A").getValues();
    var lastRow = 0;
    for (var i = data.length - 1; i >= 0; i--) {
      if (data[i][0] != "") {
        lastRow = i + 1;
        break;
      }
    }
    var nextSuffix = 1;

    // 2. Determine Next Suffix (Continuous Sequence)
    if (lastRow > 1) {
      var lastRfpValue = sheet.getRange(lastRow, 1).getValue().toString();

      // Use regex to find the last group of digits at the end of the string
      // This ensures we get the counter even if the YYYY-MM prefix changes
      var match = lastRfpValue.match(/(\d+)$/);
      if (match) {
        nextSuffix = parseInt(match[1], 10) + 1;
      }
    }

    // 3. Get Current Date in Philippine Time (GMT+8)
    var now = new Date();
    var yearMonth = Utilities.formatDate(now, "GMT+8", "yyyy-MM");

    // Generate RFP ID: MALL-YYYY-MM-000000X
    var newRfp = "MALL-" + yearMonth + "-" + ("0000000" + nextSuffix).slice(-7);

    // 4. Deduplication Check: Search Column A for the new ID
    var range = sheet.getRange("A:A");
    var duplicate = range.createTextFinder(newRfp).matchEntireCell(true).findNext();

    if (duplicate) {
      console.warn("Duplicate detected for " + newRfp + ". Incrementing sequence...");
      // Increment suffix and try again to avoid recursion depth issues
      nextSuffix++;
      newRfp = "MALL-" + yearMonth + "-" + ("0000000" + nextSuffix).slice(-7);
    }

    // 5. Persistence
    sheet.appendRow([newRfp, now]);

    // 6. Force write to DB before releasing lock
    SpreadsheetApp.flush();

    return newRfp;
  } catch (e) {
    console.error("Error generating RFP: " + e.toString());
    throw new Error("Failed to generate RFP Number: " + e.message);
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

/** 
 * UPDATED: Fetches Filename, Tab, and URL
 */
function getSourceFileNames() {
  const MASTER_ID = "10p-nv_qAN0GzZHAVInyb9_bprk2_sLf08190qdMf8Mc";
  const ss = SpreadsheetApp.openById(MASTER_ID);
  const sheet = ss.getSheetByName("SOURCEFILES");

  if (!sheet) throw new Error("Sheet 'SOURCEFILES' not found.");

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data.shift().map((h) => h.toString().trim().toUpperCase());

  // Indices
  let nameColIdx = headers.indexOf("FILE NAME");
  let tabColIdx = headers.indexOf("TAB NAME");
  let urlColIdx = headers.indexOf("FILE URL"); // Retrieve URL Column

  const fileData = data
    .map((row) => {
      return {
        fileName: row[nameColIdx] ? row[nameColIdx].toString().trim() : "",
        tabName: row[tabColIdx] ? row[tabColIdx].toString().trim() : "",
        fileUrl: row[urlColIdx] ? row[urlColIdx].toString().trim() : "" // Captured here
      };
    })
    .filter((item) => item.fileName !== "");

  return fileData;
}

/**
 * Unified Advanced Search used by the "Search" button in UI
 */
function getRfpDataAdvanced(sourceFile, rfpInput, invoiceInput) {
  try {
    const result = performStrictSearch(sourceFile, rfpInput, invoiceInput);

    if (result.length > 0) {
      // 1. Check if the FIRST match is already PAID
      // (Assumption: if one matches and is PAID, the record is locked)
      const currentStatus = String(result[0]["GENERAL STATUS"] || "")
        .trim()
        .toUpperCase();

      if (currentStatus === "PAID") {
        return {
          status: "ALREADY_PAID",
          message: "The record found (" + (rfpInput || invoiceInput) + ") is already marked as PAID.",
        };
      }

      // 2. Return data if allowed
      return {
        status: "FOUND",
        data: result,
      };
    }

    return { status: "NOT_FOUND" };
  } catch (e) {
    throw new Error("Search logic failed: " + e.message);
  }
}

function performStrictSearch(selectedSource, rfpInput, invoiceInput) {
  const rfpClean = String(rfpInput || "")
    .trim()
    .toUpperCase();
  const invClean = String(invoiceInput || "")
    .trim()
    .toUpperCase();

  const MASTER_ID = "10p-nv_qAN0GzZHAVInyb9_bprk2_sLf08190qdMf8Mc";
  const indexSheet = SpreadsheetApp.openById(MASTER_ID).getSheetByName("SOURCEFILES");
  const indexData = indexSheet.getDataRange().getValues();
  const indexHeaders = indexData.shift().map((h) => h.toString().trim().toUpperCase());

  const urlColIdx = indexHeaders.indexOf("FILE URL");
  const tabColIdx = indexHeaders.indexOf("TAB NAME");
  let nameColIdx = indexHeaders.indexOf("FILE NAME");
  if (nameColIdx === -1) nameColIdx = 0;

  const sourceRow = indexData.find(
    (row) => row[nameColIdx].toString().trim().toUpperCase() === selectedSource.toUpperCase()
  );

  if (!sourceRow) return [];

  const targetSs = SpreadsheetApp.openByUrl(sourceRow[urlColIdx]);
  const tz = targetSs.getSpreadsheetTimeZone();
  const tabs = sourceRow[tabColIdx]
    .toString()
    .split(",")
    .map((t) => t.trim());

  let matches = [];

  tabs.forEach((tabName) => {
    const sheet = targetSs.getSheetByName(tabName);
    if (!sheet) return;

    const data = sheet.getRange(5, 1, sheet.getLastRow() - 4, sheet.getLastColumn()).getValues();
    const headers = data[0].map((h) => h.toString().trim().toUpperCase());

    const rfpIdx = headers.indexOf("RFP|PEF NO.");
    const invIdx = headers.indexOf("INVOICE NO.");

    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      const rfpVal = String(row[rfpIdx] || "")
        .trim()
        .toUpperCase();
      const invVal = String(row[invIdx] || "")
        .trim()
        .toUpperCase();

      let isMatch = false;
      if (rfpClean && rfpVal === rfpClean) isMatch = true;
      else if (invClean && invVal === invClean) isMatch = true;

      if (isMatch) {
        let record = {};
        headers.forEach((h, idx) => {
          let val = row[idx];
          record[h || "COL_" + idx] = val instanceof Date ? Utilities.formatDate(val, tz, "yyyy-MM-dd") : val;
        });
        matches.push(record);
      }
    }
  });

  return matches;
}

/**
 * FORGOT PASSWORD - STAGE 1: Check account and send OTP
 * Triggered by: google.script.run.requestPasswordResetCode(email)
 */
function requestPasswordResetCode(email) {
  const sessionEmail = Session.getActiveUser().getEmail();
  if (!email || email !== sessionEmail)
    throw new Error("Security Violation: You can only reset the password for your own authenticated account.");

  const ss = SpreadsheetApp.openById(USER_DB_ID);
  const sheet = ss.getSheetByName(USER_TAB);
  const data = sheet.getDataRange().getValues();

  // Find index 2 (Column C - USERNAME/EMAIL)
  const accountExists = data.some((row) => row[2] === email);
  if (!accountExists) throw new Error("No registered account found for this email address.");

  const otp = Math.floor(100000 + Math.random() * 900000).toString();
  const cache = CacheService.getScriptCache();
  cache.put("RESET_" + email, otp, 600); // Unique cache key for resets

  try {
    MailApp.sendEmail(email, "Password Reset Code", "Your password reset verification code is: " + otp);
    return { success: true, message: "Verification code sent." };
  } catch (e) {
    throw new Error("SMTP Error: Failed to send reset code.");
  }
}

/**
 * FORGOT PASSWORD - STAGE 2: Validate Reset OTP
 * Triggered by: google.script.run.verifyResetCode(email, otp)
 */
function verifyResetCode(email, userCode) {
  const cache = CacheService.getScriptCache();
  const storedCode = cache.get("RESET_" + email);

  if (!storedCode) throw new Error("Reset code expired. Please request a new one.");
  if (storedCode !== userCode) throw new Error("The reset code is incorrect.");

  return { success: true, message: "Identity verified for reset." };
}

/**
 * FORGOT PASSWORD - STAGE 3: Overwrite Old Password
 * Includes fix for numeric passwords and data latency.
 */
function submitPasswordReset(email, newPassword) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const ss = SpreadsheetApp.openById(USER_DB_ID);
    const sheet = ss.getSheetByName(USER_TAB);
    const data = sheet.getDataRange().getDisplayValues();
    
    let rowIndex = -1;
    const cleanEmail = email.toString().trim().toLowerCase();
    
    // Salted with user email
    const hashedNewPassword = generateSHA256(newPassword, cleanEmail);

    for (let i = 0; i < data.length; i++) {
      if (data[i][2].toString().trim().toLowerCase() === cleanEmail) {
        rowIndex = i + 1;
        if (data[i][3] === hashedNewPassword) throw new Error("Must be a different password.");
        break;
      }
    }

    if (rowIndex === -1) throw new Error("Record not found.");
    sheet.getRange(rowIndex, 4).setValue(hashedNewPassword);
    SpreadsheetApp.flush();
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * PROCESS SUBMISSION - Optimized for Data Integrity & Specific Error Messaging
 */
function processSubmission(p) {
  const lock = LockService.getPublicLock();
  const SS_ID = "1YAvZmCdWXbjOcJA-uUY40e6qVqzyiHcB06NpiPcz6y4";
  const FOLDER_ID = "1eFcLGXPEnUSi14aPvvA9m2ATV8Racsuf";
  const MAX_CELL_LIMIT = 9800000;

  try {
    if (!lock.tryLock(30000)) throw new Error("Server Timeout: The database is busy. Please wait a moment and try again.");

    // 1. Identity Verification (Server-Side Truth)
    const authEmail = Session.getActiveUser().getEmail().toLowerCase().trim();
    if (!authEmail) throw new Error("Security Alert: We could not identify your Google account. Please reload the app.");

    // 2. Data Validation & Formatting
    const rfpNo = p.header.rfpNo ? p.header.rfpNo.toString().trim() : "";
    const invNo = p.header.invoiceNo ? p.header.invoiceNo.toString().trim() : "";
    const identifier = rfpNo || invNo;

    if (!invNo) throw new Error("Input Required: An Invoice Number is mandatory for recording.");
    if (!p.attachment || !p.attachment.base64) throw new Error("Attachment Missing: You must upload a supporting PDF.");

    const ss = SpreadsheetApp.openById(SS_ID);
    const sh = ss.getSheetByName("SUBMISSIONS");

    // 3. System Limits Check
    const totalCells = ss.getSheets().reduce((sum, s) => sum + (s.getMaxRows() * s.getMaxColumns()), 0);
    if (totalCells >= MAX_CELL_LIMIT) throw new Error("Database Full: System cell limit (10M) approaching. Please archive old transactions.");

    // 4. Primary/Secondary Conflict Logic
    const primaryEmails = p.participants.filter(x => x.tag === "Primary").map(x => x.email.toLowerCase().trim());
    const secondaryRaw = p.participants.filter(x => x.tag === "Secondary").map(x => x.email.toLowerCase().trim());
    
    // Server-side guard against overlaps
    const overlapEmails = secondaryRaw.filter(email => primaryEmails.includes(email));
    if (overlapEmails.length > 0) throw new Error("Duplicate Roles: The email(s) " + overlapEmails.join(", ") + " are set as both Primary AND Secondary. This is not allowed.");

    // 5. RFP Duplicate Check
    if (rfpNo !== "") {
      const existingRfp = sh.getRange("F:F").createTextFinder(rfpNo).matchEntireCell(true).findNext();
      if (existingRfp) throw new Error("Transaction Already Exists: RFP No '" + rfpNo + "' was previously submitted.");
    }

    // 6. Drive & Folder Scoping
    const folder = DriveApp.getFolderById(FOLDER_ID);

    // Secure Support Upload
    let supportUrl = "";
    try {
      const blob = Utilities.newBlob(Utilities.base64Decode(p.attachment.base64), "application/pdf", safeValue(identifier) + "_Support.pdf");
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.NONE); 
      supportUrl = file.getUrl();
    } catch (driveErr) {
      throw new Error("File Upload Failed: System couldn't save to the designated folder. Please check your Drive space.");
    }

    // Secure RFP Auto-Gen
    const rfpCopyUrl = createRfpPdf(p, folder);

    // 7. Field Mapping with Sanitization
    let parts = {};
    p.tableFields.forEach(f => {
      let v = p.particulars[f] || "";
      if (f.toUpperCase().includes("AMOUNT")) v = parseFloat(v.toString().replace(/[^0-9.-]/g, "")) || 0;
      else v = safeValue(v); 
      parts[f] = v;
    });

    const finalRow = [
      "TXN-" + Utilities.getUuid().split("-")[0].toUpperCase(), 
      new Date(), 
      authEmail, 
      primaryEmails.join(", "), 
      secondaryRaw.join(", "), // Safe as it was de-duplicated from primary above
      safeValue(rfpNo || "N/A"), 
      safeValue(p.header.dueDate || "N/A"),
      parts["YEAR"], parts["MONTH"], parts["PAYOR NAME"], parts["PAYEE NAME"],
      parts["PROPERTY"], parts["LOCATION"], parts["SECTOR"], parts["KINDS OF SERVICE"],
      parts["CONTRACT NO"], parts["CONTRACT AMOUNT"],
      safeValue(invNo), 
      parts["BILLING PERIOD"], parts["SOA AMOUNT"], parts["GENERAL STATUS"],
      supportUrl, 
      rfpCopyUrl, 
      safeValue(p.header.sourceFileName || "No Reference File"), 
      safeValue(p.header.sourceTabName || "N/A")
    ];

    sh.appendRow(finalRow);
    SpreadsheetApp.flush();

    return { success: true, message: identifier };

  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

/**
 * Generates a Secure PDF with Logo and Sanity Checks
 */
function createRfpPdf(p, folder) {
  const LOGO_FILE_ID = "1QZ3XkRk1x-p_GFSjKhQO8i4dmIVzE1dG";

  // --- LOGO FETCHING LOGIC ---
  let logoBase64 = "";
  try {
    const logoFile = DriveApp.getFileById(LOGO_FILE_ID);
    logoBase64 = "data:" + logoFile.getMimeType() + ";base64," + Utilities.base64Encode(logoFile.getBlob().getBytes());
  } catch (e) {
    console.warn("Logo could not be loaded: " + e.message);
  }

  // --- SECURE PARTICULARS TABLE GENERATION ---
  let tableRowsHtml = "";
  for (let i = 0; i < p.tableFields.length; i += 3) {
    tableRowsHtml += "<tr>";
    for (let j = 0; j < 3; j++) {
      const field = p.tableFields[i + j];
      if (field) {
        tableRowsHtml += `
          <td class="grid-cell" style="width: 33.33%;">
             <div class="field-label">${safeValue(field.replace(/_/g, " "))}</div>
             <div class="field-value">${safeValue(String(p.particulars[field] || "—"))}</div>
          </td>`;
      } else {
        tableRowsHtml += `<td class="grid-cell" style="width: 33.33%;"></td>`;
      }
    }
    tableRowsHtml += "</tr>";
  }

  // --- STATIC NOTE LINES ---
  let noteLinesHtml = "";
  for (let n = 0; n < 5; n++) {
    noteLinesHtml += '<div class="note-line"></div>';
  }

  // --- COMPLETE SECURE HTML TEMPLATE ---
  const html = `
  <html>
  <head>
    <style>
      @page { size: letter portrait; margin: 0.35in; }
      body { font-family: 'Arial', sans-serif; font-size: 10pt; color: #000; line-height: 1.2; margin: 0; }
      .brand-wrapper { text-align: center; margin-bottom: 8px; width: 100%; }
      .logo-img { width: 170px; height: auto; }
      .doc-title { 
        text-align: center; font-size: 17pt; font-weight: bold; 
        text-transform: uppercase; border-bottom: 3px solid #000;
        padding-bottom: 8px; margin-top: 2px; margin-bottom: 15px;
      }
      .meta-table { width: 100%; margin-bottom: 15px; border-collapse: collapse; }
      .meta-label { font-size: 7.5pt; color: #555; text-transform: uppercase; font-weight: bold; }
      .meta-value { font-size: 11pt; font-weight: bold; border-bottom: 1px solid #ccc; padding: 2px 0; }
      .particulars-grid { width: 100%; border-collapse: collapse; table-layout: fixed; border: 1.5px solid #000; }
      .grid-cell { border: 1px solid #000; padding: 6px 8px; vertical-align: top; word-wrap: break-word; min-height: 42px; }
      .field-label { font-size: 7pt; color: #333; text-transform: uppercase; font-weight: bold; margin-bottom: 3px; }
      .field-value { font-size: 9.5pt; font-weight: bold; color: #000; }
      .notes-title { font-size: 9pt; font-weight: bold; text-transform: uppercase; margin-top: 15px; margin-bottom: 5px; }
      .note-line { width: 100%; height: 23px; border-bottom: 1px solid #000; }
      .sig-table { margin-top: 35px; width: 100%; border-collapse: collapse; }
      .sig-line { border-top: 1.5px solid #000; width: 85%; margin: 35px auto 0 auto; padding-top: 6px; font-size: 8pt; font-weight: bold; text-align: center; text-transform: uppercase; }
    </style>
  </head>
  <body>
    <div class="brand-wrapper">${logoBase64 ? `<img src="${logoBase64}" class="logo-img">` : ""}</div>
    <div class="doc-title">Request for Payment</div>

    <table class="meta-table">
      <tr>
        <td style="width: 45%;">
          <div class="meta-label">RFP Number</div>
          <div class="meta-value" style="font-size: 14pt;">${safeValue(p.header.rfpNo)}</div>
        </td>
        <td style="width: 27.5%; vertical-align: bottom;">
          <div class="meta-label">Date Requested</div>
          <div class="meta-value">${safeValue(p.header.date)}</div>
        </td>
        <td style="width: 27.5%; vertical-align: bottom;">
          <div class="meta-label">Due Date</div>
          <div class="meta-value">${safeValue(p.header.dueDate)}</div>
        </td>
      </tr>
    </table>

    <table class="particulars-grid">
      ${tableRowsHtml}
    </table>

    <div class="notes-title">NOTES / REMARKS:</div>
    ${noteLinesHtml}

    <table class="sig-table">
      <tr>
        <td><div class="sig-line">Requested By</div></td>
        <td><div class="sig-line">Verified By</div></td>
        <td><div class="sig-line">Approved By</div></td>
      </tr>
    </table>
  </body>
  </html>`;

  // Create PDF
  const pdfBlob = Utilities.newBlob(html, "text/html").getAs("application/pdf").setName(`RFP_${p.header.rfpNo}.pdf`);
  const file = folder.createFile(pdfBlob);
  
  // Set explicit security permissions
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.NONE);
  
  return file.getUrl();
}

/**
 * Backend: Retrieves suggestions for the email search list.
 */
function getPreviousParticipantEmails() {
  const SS_ID = "1YAvZmCdWXbjOcJA-uUY40e6qVqzyiHcB06NpiPcz6y4";
  try {
    const sh = SpreadsheetApp.openById(SS_ID).getSheetByName("SUBMISSIONS");
    if (!sh) throw new Error("Data retrieval source (SUBMISSIONS) not found.");

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return [];

    const data = sh.getRange(2, 4, lastRow - 1, 2).getValues(); 
    const results = new Set();
    
    data.forEach(row => {
      row.forEach(cell => {
        if (cell && cell.toString().includes("@")) {
          cell.toString().split(",").forEach(email => {
            const clean = email.trim().toLowerCase();
            if (clean) results.add(clean);
          });
        }
      });
    });

    return Array.from(results).sort();
  } catch (e) {
    throw new Error("System Suggestion Failure: " + e.message);
  }
}
/**
 * Frontend: addSecondaryRecipient
 * Blocks addition if already in Primary list or invalid.
 */
function addSecondaryRecipient(emailToVerify) {
  const cleanEmail = emailToVerify.trim().toLowerCase();
  
  if (!cleanEmail.includes("@") || !cleanEmail.includes(".")) {
    Swal.fire("Invalid Format", "The entry '" + cleanEmail + "' does not appear to be a valid email address.", "error");
    return;
  }

  // Identify overlaps
  const isPrimary = formState.participants.some(p => p.email.toLowerCase() === cleanEmail && p.tag === "Primary");

  if (isPrimary) {
    Swal.fire({
      title: "Duplicate Found",
      text: "The person " + cleanEmail + " is already assigned as a Primary Recipient.",
      icon: "warning"
    });
    return;
  }

  // Prevent UI double-entry in the same category
  if (formState.participants.some(p => p.email.toLowerCase() === cleanEmail && p.tag === "Secondary")) return;

  // Function from your existing logic to add pill to UI
  addParticipant(cleanEmail, "Secondary"); 
}
