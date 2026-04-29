/**
 * UI INITIALIZATION
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Payment Endorsement Platform")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
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

    const activeUser = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.openById(USER_DB_ID);
    let sheet = ss.getSheetByName(USER_TAB);

    // Auto-setup if sheet missing (Headers A-H)
    if (!sheet) {
      sheet = ss.insertSheet(USER_TAB);
      sheet.appendRow([
        "TIMESTAMP",
        "FULL NAME",
        "USERNAME",
        "PASSWORD",
        "ACCESS LEVEL",
        "ORGANIC OR NON ORGANIC",
        "PERMISSION STATUS",
        "EMPLOYEE STATUS", // Col H
      ]);
    }

    const timestamp = new Date();
    const username = activeUser;
    const accessLevel = "REQUESTOR";
    const permissionStatus = "PENDING";
    const organicStatus = activeUser.toLowerCase().includes("@megaworld-lifestyle.com") ? "ORGANIC" : "NON ORGANIC";

    // GAP FILLING: Find first empty Row (checks Col B - Full Name)
    const nameCol = sheet.getRange("B:B").getValues();
    let targetRow = -1;
    for (let i = 1; i < nameCol.length; i++) {
      if (nameCol[i][0] === "" || nameCol[i][0] === null) {
        targetRow = i + 1;
        break;
      }
    }
    if (targetRow === -1) targetRow = sheet.getLastRow() + 1;

    /**
     * UPDATED LOGIC:
     * We only prepare data for Columns A through G (7 columns).
     * Column H (Index 8) is ignored entirely.
     */
    const rowPayload = [
      timestamp, // A
      formData.fullName, // B
      username, // C
      formData.password, // D
      accessLevel, // E
      organicStatus, // F
      permissionStatus, // G
      "INACTIVE", // H (Better to set a default for safety)
    ];

    // Update range to cover 8 columns
    sheet.getRange(targetRow, 1, 1, rowPayload.length).setValues([rowPayload]);

    return { success: true, message: "Registered under " + activeUser };
  } catch (e) {
    throw new Error("Persistence Error: " + e.message);
  } finally {
    lock.releaseLock();
  }
}
/**
 * Main authentication function
 */
function authenticateUser(credentials) {
  // 1. Force Google to commit all pending changes before reading
  SpreadsheetApp.flush();

  const ss = SpreadsheetApp.openById(USER_DB_ID);
  const sheet = ss.getSheetByName(USER_TAB);

  // 2. Read spreadsheet data
  const data = sheet.getDataRange().getDisplayValues();
  data.shift(); // Remove headers

  // 3. Normalize inputs
  const inputEmail = String(credentials.email || "").trim().toLowerCase();
  
  // --- PASSWORD HASHING ---
  const rawInputPass = String(credentials.password || "").trim();
  const hashedInputPass = generateSHA256(rawInputPass); 
  // -------------------------

  const userRow = data.find((row) => {
    const storedEmail = String(row[2] || "").trim().toLowerCase();
    const storedPass = String(row[3] || "").trim(); // Matches against your 64-character hash
    return storedEmail === inputEmail && storedPass === hashedInputPass;
  });

  if (!userRow) throw new Error("Invalid credentials.");

  // --- NEW STATUS LOGIC ---
  const permission = String(userRow[6] || "").trim().toUpperCase(); // Col G
  const employeeStatus = String(userRow[7] || "").trim().toUpperCase(); // Col H

  // 1. If not APPROVED, login is rejected regardless of Employee Status
  if (permission !== "APPROVED") {
    throw new Error(`Access Denied: Account is ${permission}. Contact admin for approval.`);
  }

  // 2. If APPROVED, check if they are RESIGNED
  if (employeeStatus === "RESIGNED") {
    throw new Error("Access Denied: Your account is locked because employee status is RESIGNED.");
  }

  // Logic outcome: (Approved AND not Resigned) will pass through here
  return {
    success: true,
    user: {
      fullName: userRow[1],
      email: userRow[2],
      role: userRow[4],
      classification: userRow[5],
    },
  };
}

/**
 * Helper to generate SHA-256 hash
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
    // Wait up to 30 seconds for other processes to finish
    lock.waitLock(30000);

    const ss = SpreadsheetApp.openById(USER_DB_ID);
    const sheet = ss.getSheetByName(USER_TAB);

    // getDisplayValues() is critical: it reads "123" as a string, not a number
    const data = sheet.getDataRange().getDisplayValues();

    let rowIndex = -1;
    let currentStoredPassword = "";

    // Normalize email for comparison
    const cleanEmail = email.toString().trim().toLowerCase();

    for (let i = 0; i < data.length; i++) {
      if (data[i][2].toString().trim().toLowerCase() === cleanEmail) {
        rowIndex = i + 1;
        currentStoredPassword = data[i][3]; // Column D
        break;
      }
    }

    if (rowIndex === -1) throw new Error("Account record not found.");

    // Security: Block reusing the same password
    if (newPassword.toString() === currentStoredPassword) {
      throw new Error("The new password must be different from your current password.");
    }

    /**
     * Update Column D
     * We convert to string explicitly to prevent Google Sheets from
     * treating "123" as a math-ready number.
     */
    sheet.getRange(rowIndex, 4).setValue(newPassword.toString());

    // CRITICAL: Forces Google to commit the save before the script finishes
    SpreadsheetApp.flush();

    return {
      success: true,
      message: "Password updated successfully. You may now log in.",
    };
  } catch (e) {
    throw new Error(e.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * High-Performance Submission Logic
 * Optimized for datasets exceeding 1 million rows/cells.
 */
/**
 * Processes the payment request submission
 * @param {Object} p - The payload from the frontend
 */
/**
 * PROCESS SUBMISSION - The Single Source of Truth for Validation
 */
function processSubmission(p) {
  const lock = LockService.getPublicLock();
  const SS_ID = "1YAvZmCdWXbjOcJA-uUY40e6qVqzyiHcB06NpiPcz6y4";
  const FOLDER_ID = "1eFcLGXPEnUSi14aPvvA9m2ATV8Racsuf";
  const MAX_CELL_LIMIT = 9800000; // Limit safety (Google's max is 10M)

  try {
    if (!lock.tryLock(30000)) throw new Error("Server busy. Please try again.");

    const ss = SpreadsheetApp.openById(SS_ID);

    // 1. CAPACITY CHECK: Check if Sheet is almost full
    const totalCells = ss.getSheets().reduce((sum, s) => sum + (s.getMaxRows() * s.getMaxColumns()), 0);
    if (totalCells >= MAX_CELL_LIMIT) {
      throw new Error("System Error: Spreadsheet capacity reached. Please archive old data.");
    }

    const sh = ss.getSheetByName("SUBMISSIONS");

    // 2. DATA EXTRACTION & VALIDATION
    const rfpNo = p.header.rfpNo ? p.header.rfpNo.toString().trim() : "";
    const invNo = p.header.invoiceNo ? p.header.invoiceNo.toString().trim() : "";
    const hasRfp = rfpNo !== "";

    if (invNo === "") throw new Error("Validation Error: Invoice Number is required.");

    // RFP Duplicate Check (Only if RFP is filled)
    if (hasRfp) {
      const duplicate = sh.getRange("F:F").createTextFinder(rfpNo).matchEntireCell(true).findNext();
      if (duplicate) throw new Error("RFP Number " + rfpNo + " has already been submitted.");
    }

    // 3. FILE ASSETS
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const identifier = hasRfp ? rfpNo : invNo;
    const blob = Utilities.newBlob(Utilities.base64Decode(p.attachment.base64), "application/pdf", identifier + "_Support.pdf");
    const uploadedFileUrl = folder.createFile(blob).getUrl();
    const rfpCopyUrl = (typeof createRfpPdf === 'function') ? createRfpPdf(p, folder) : "N/A";

    // 4. THE ROW MAP: Define strictly where each data point lands
    // This allows you to easily move rows or see column indices at a glance.
    
    // Prep Particulars list
    let partsMap = {};
    p.tableFields.forEach(f => {
      let val = p.particulars[f] || "";
      if (f.toUpperCase().includes("AMOUNT")) val = parseFloat(val.toString().replace(/[^0-9.-]/g, "")) || 0;
      partsMap[f] = val;
    });

    const finalRowData = [
      /* Col A: SUBMISSION ID         */ "TXN-" + Utilities.getUuid().split("-")[0].toUpperCase(),
      /* Col B: SUBMISSION DATE       */ new Date(),
      /* Col C: USER EMAIL            */ Session.getActiveUser().getEmail(),
      /* Col D: PRIMARY RECIPIENT      */ p.participants.filter(x => x.tag === "Primary").map(x => x.email).join(", "),
      /* Col E: SECONDARY RECIPIENT    */ p.participants.filter(x => x.tag === "Secondary").map(x => x.email).join(", "),
      /* Col F: RFP|PEF NO.           */ hasRfp ? rfpNo : "N/A",
      /* Col G: DUE DATE              */ p.header.dueDate || "N/A",
      
      /* --- START OF PARTICULARS --- */
      /* Col H: YEAR                  */ partsMap["YEAR"],
      /* Col I: MONTH                 */ partsMap["MONTH"],
      /* Col J: PAYOR NAME            */ partsMap["PAYOR NAME"],
      /* Col K: PAYEE NAME            */ partsMap["PAYEE NAME"],
      /* Col L: PROPERTY              */ partsMap["PROPERTY"],
      /* Col M: LOCATION              */ partsMap["LOCATION"],
      /* Col N: SECTOR                */ partsMap["SECTOR"],
      /* Col O: KINDS OF SERVICE      */ partsMap["KINDS OF SERVICE"],
      /* Col P: CONTRACT NO           */ partsMap["CONTRACT NO"],
      /* Col Q: CONTRACT AMOUNT       */ partsMap["CONTRACT AMOUNT"],

      /* Col R: INVOICE NO            */ invNo, // <--- INVOICE PLACED AFTER CONTRACT AMOUNT

      /* Col S: BILLING PERIOD        */ partsMap["BILLING PERIOD"],
      /* Col T: SOA AMOUNT            */ partsMap["SOA AMOUNT"],
      /* Col U: GENERAL STATUS        */ partsMap["GENERAL STATUS"],
      /* --- END OF PARTICULARS ---   */

      /* Col V: UPLOADED FILE         */ uploadedFileUrl,
      /* Col W: RFP COPY              */ rfpCopyUrl,
      /* Col X: SOURCEFILE            */ p.header.sourceFileUrl || p.header.sourceFileName,
      /* Col Y: SOURCEFILE TABS       */ p.header.sourceTabName
    ];

    // 5. SUBMIT TO SHEET
    sh.appendRow(finalRowData);
    SpreadsheetApp.flush();

    // 6. RETURN SUCCESS
    return {
      success: true,
      message: `${hasRfp ? 'RFP Number' : 'Invoice Number'}: ${identifier}`
    };

  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}
/**
 * Generates a PDF.
 */
function createRfpPdf(p, folder) {
  const LOGO_FILE_ID = "1QZ3XkRk1x-p_GFSjKhQO8i4dmIVzE1dG";

  const getLogoDataUri = (fileId) => {
    try {
      const file = DriveApp.getFileById(fileId);
      return "data:" + file.getMimeType() + ";base64," + Utilities.base64Encode(file.getBlob().getBytes());
    } catch (e) {
      return "";
    }
  };

  const logoBase64 = getLogoDataUri(LOGO_FILE_ID);

  // --- REFACTOR 1: PARTICULARS LOOP ---
  let tableRowsHtml = "";
  for (let i = 0; i < p.tableFields.length; i += 3) {
    tableRowsHtml += "<tr>";
    for (let j = 0; j < 3; j++) {
      const field = p.tableFields[i + j];
      tableRowsHtml += field
        ? `<td class="grid-cell" style="width: 33.33%;">
             <div class="field-label">${field.replace(/_/g, " ")}</div>
             <div class="field-value">${p.particulars[field] || "—"}</div>
           </td>`
        : `<td class="grid-cell" style="width: 33.33%;"></td>`;
    }
    tableRowsHtml += "</tr>";
  }

  // --- REFACTOR 2: GENERATE 5 LINES ---
  let noteLinesHtml = "";
  for (let n = 0; n < 5; n++) {
    noteLinesHtml += '<div class="note-line"></div>';
  }

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
      .particulars-grid { 
        width: 100%; 
        border-collapse: collapse; 
        table-layout: fixed; /* ESSENTIAL: Keeps columns at exactly 33.3% */
        border: 1.5px solid #000;
      }
       .grid-cell { 
        border: 1px solid #000; 
        padding: 6px 8px; 
        vertical-align: top; 
        
        /* THE OVERLAP FIX: */
        word-wrap: break-word;       /* Standard */
        overflow-wrap: break-word;  /* Modern fallback */
        word-break: break-word;     /* Support for long continuous strings */
        white-space: normal;        /* Allows the line to break */
        overflow: hidden;           /* Safety: clipped if it tries to invade neighbor */
        
        /* Allow height to grow based on content */
        height: auto; 
        min-height: 42px;           /* Minimum visual height for short data */
      }

      .field-label { 
        font-size: 7pt; 
        color: #333; 
        text-transform: uppercase; 
        font-weight: bold; 
        margin-bottom: 3px; 
      }
      .field-value { 
        font-size: 9.5pt;           /* Slightly larger for clarity */
        font-weight: bold; 
        display: block;             /* Ensure it fills the cell container */
        color: #000;
      }
      
      .notes-title { font-size: 9pt; font-weight: bold; text-transform: uppercase; margin-top: 15px; margin-bottom: 5px; }
      /* Spacing fix for 5 lines */
      .note-line { width: 100%; height: 23px; border-bottom: 1px solid #000; }

      .sig-table { margin-top: 35px; width: 100%; border-collapse: collapse; }
      .sig-line { 
        border-top: 1.5px solid #000; width: 85%; margin: 35px auto 0 auto; 
        padding-top: 6px; font-size: 8pt; font-weight: bold; text-align: center; text-transform: uppercase;
      }
    </style>
  </head>
  <body>
    <div class="brand-wrapper">${logoBase64 ? `<img src="${logoBase64}" class="logo-img">` : ""}</div>
    <div class="doc-title">Request for Payment</div>

    <table class="meta-table">
      <tr>
        <td style="width: 45%;">
          <div class="meta-label">RFP Number</div>
          <div class="meta-value" style="font-size: 14pt;">${p.header.rfpNo}</div>
        </td>
        <td style="width: 27.5%; vertical-align: bottom;">
          <div class="meta-label">Date Requested</div>
          <div class="meta-value">${p.header.date}</div>
        </td>
        <td style="width: 27.5%; vertical-align: bottom;">
          <div class="meta-label">Due Date</div>
          <div class="meta-value">${p.header.dueDate}</div>
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

  const pdfBlob = Utilities.newBlob(html, "text/html").getAs("application/pdf").setName(`RFP_${p.header.rfpNo}.pdf`);
  return folder.createFile(pdfBlob).getUrl();
}
/**
 * Retrieves a unique list of emails used in previous transactions for suggestions.
 */
function getPreviousParticipantEmails() {
  const SS_ID = "1YAvZmCdWXbjOcJA-uUY40e6qVqzyiHcB06NpiPcz6y4";
  const sh = SpreadsheetApp.openById(SS_ID).getSheetByName("SUBMISSIONS");
  const lastRow = sh.getLastRow();

  if (lastRow < 2) return [];

  // We are grabbing Columns D (Primary) and E (Secondary)
  // Which are index 4 and 5 in the row
  const data = sh.getRange(2, 4, lastRow - 1, 2).getValues();

  let emailSet = new Set();

  data.forEach((row) => {
    // Column D (Primary Emails string)
    if (row[0]) {
      row[0].split(",").forEach((e) => {
        let clean = e.trim().toLowerCase();
        if (clean.includes("@")) emailSet.add(clean);
      });
    }
    // Column E (Secondary Emails string)
    if (row[1]) {
      row[1].split(",").forEach((e) => {
        let clean = e.trim().toLowerCase();
        if (clean.includes("@")) emailSet.add(clean);
      });
    }
  });

  // Convert Set back to an Array and return
  return Array.from(emailSet).sort();
}
