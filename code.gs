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
      "Google Identity not found. Please ensure you are logged into your browser with your company email.",
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
  if (!email || email !== sessionEmail)
    throw new Error("Security Violation: Identity mismatch.");

  const ss = SpreadsheetApp.openById(USER_DB_ID);
  const sheet = ss.getSheetByName(USER_TAB);
  const data = sheet.getDataRange().getValues();

  // Index 2 is USERNAME/EMAIL (Column C)
  const alreadyExists = data.some((row) => row[2] === email);
  if (alreadyExists)
    throw new Error(
      "This email is already registered. Please proceed to Login.",
    );

  const otp = Math.floor(100000 + Math.random() * 900000).toString();
  const cache = CacheService.getScriptCache();
  cache.put(email, otp, 600); // 10 min expiry

  try {
    MailApp.sendEmail(
      email,
      "Account Verification Code",
      "Your verification code is: " + otp,
    );
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
  if (email !== sessionEmail)
    throw new Error("Security Violation: Session hijacked.");

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
    const organicStatus = activeUser
      .toLowerCase()
      .includes("@megaworld-lifestyle.com")
      ? "ORGANIC"
      : "NON ORGANIC";

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
      timestamp, // A: TIMESTAMP
      formData.fullName, // B: FULL NAME
      username, // C: USERNAME
      formData.password, // D: PASSWORD
      accessLevel, // E: ACCESS LEVEL
      organicStatus, // F: ORGANIC OR NON ORGANIC
      permissionStatus, // G: PERMISSION STATUS
    ];

    // Target columns 1 through 7 (A-G) only. Column 8 (H) is not touched.
    sheet.getRange(targetRow, 1, 1, rowPayload.length).setValues([rowPayload]);

    return { success: true, message: "Registered under " + activeUser };
  } catch (e) {
    throw new Error("Persistence Error: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * AUTHENTICATION (Login Logic)
 * Enforcement: Employee must be ACTIVE & Permission must be ACTIVE
 * FIX: Uses getDisplayValues and Flush to handle numeric passwords like '123'
 */
function authenticateUser(credentials) {
  // 1. Force Google to commit all pending changes before reading
  SpreadsheetApp.flush();

  const ss = SpreadsheetApp.openById(USER_DB_ID);
  const sheet = ss.getSheetByName(USER_TAB);

  // 2. Use getDisplayValues() to ensure "123" is read as a string "123"
  const data = sheet.getDataRange().getDisplayValues();
  data.shift(); // Remove headers

  // 3. Normalize inputs for comparison
  const inputEmail = String(credentials.email || "")
    .trim()
    .toLowerCase();
  const inputPass = String(credentials.password || "").trim();

  const userRow = data.find((row) => {
    const storedEmail = String(row[2] || "")
      .trim()
      .toLowerCase();
    const storedPass = String(row[3] || "").trim();
    return storedEmail === inputEmail && storedPass === inputPass;
  });

  if (!userRow) throw new Error("Invalid credentials.");

  // Validation 1: EMPLOYEE STATUS (Col H / Index 7)
  const employeeStatus = String(userRow[7] || "")
    .trim()
    .toUpperCase();
  if (employeeStatus !== "ACTIVE") {
    throw new Error(
      "Access Denied: Your employee status is " + employeeStatus + ".",
    );
  }

  // Validation 2: PERMISSION STATUS (Col G / Index 6)
  const permission = String(userRow[6] || "")
    .trim()
    .toUpperCase();
  if (permission === "PENDING") {
    throw new Error("Account Pending: Awaiting admin activation.");
  }
  if (permission !== "APPROVED") {
    throw new Error("Access Denied: Account status is " + permission + ".");
  }

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

    var lastRow = sheet.getLastRow();
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
    var duplicate = range
      .createTextFinder(newRfp)
      .matchEntireCell(true)
      .findNext();

    if (duplicate) {
      console.warn(
        "Duplicate detected for " + newRfp + ". Incrementing sequence...",
      );
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
 * Fetches RFP data based on an RFP number, with a block on records 
 * marked with "GENERAL STATUS" as "PAID".
 */
function getRfpData(rfpNumber, offset = 0) {
  // Call the high-performance search function
  const result = performCrossFileSearch(rfpNumber, offset);

  // CASE 1: A matching record was found
  if (result.records && result.records.length > 0) {
    const raw = result.records[0];

    // --- SECURITY/STATUS CHECK: BLOCK IF PAID ---
    // Note: 'raw' uses the keys directly from your sheet headers
    const currentStatus = String(raw["GENERAL STATUS"] || "").trim().toUpperCase();

    if (currentStatus === "PAID") {
      return {
        status: "ALREADY_PAID",
        message: "This RFP (" + rfpNumber + ") is already marked as PAID and cannot be retrieved for a new submission."
      };
    }
    // --------------------------------------------

    // List of specific fields required for your UI form
    const tableFields = [
      "YEAR",
      "MONTH",
      "PAYOR NAME",
      "PAYEE NAME",
      "PROPERTY",
      "LOCATION",
      "SECTOR",
      "KINDS OF SERVICE",
      "CONTRACT NO",
      "CONTRACT AMOUNT",
      "INVOICE NO.",
      "BILLING PERIOD",
      "SOA AMOUNT",
      "GENERAL STATUS",
    ];

    // Format the particulars object based on the matched record
    const formattedParticulars = {};
    tableFields.forEach((f) => {
      // If the field exists in the record, use it; otherwise, return empty string
      formattedParticulars[f] = (raw[f] !== undefined && raw[f] !== null) ? raw[f] : "";
    });

    return {
      status: "FOUND",
      // Handles both header naming conventions "DUE DATE" or "DATE DUE"
      dueDate: raw["DUE DATE"] || raw["DATE DUE"] || "",
      particulars: formattedParticulars,
    };
  }

  // CASE 2: The script reached the time limit (210s) and needs to run again
  if (result.nextOffset !== null) {
    return {
      status: "CONTINUE",
      nextOffset: result.nextOffset,
    };
  }

  // CASE 3: Entire index was searched and no record was found
  return { status: "NOT_FOUND" };
}

/**
 * Optimized Cross-File Search
 * Searches for rfpNumber across multiple spreadsheets listed in a Master file.
 */
function performCrossFileSearch(rfpNumber, offset = 0) {
  const START_TIME = Date.now();
  const TIME_LIMIT = 210000; // 3.5 minutes (Safe buffer)
  const cleanInput = rfpNumber.toString().trim().toUpperCase();

  const MASTER_ID = "10p-nv_qAN0GzZHAVInyb9_bprk2_sLf08190qdMf8Mc";

  // 1. Get the list of files to search
  const indexSheet = SpreadsheetApp.openById(MASTER_ID).getSheetByName("SOURCEFILES");
  const indexData = indexSheet.getDataRange().getValues();
  const indexHeaders = indexData.shift().map(h => h.toString().trim().toUpperCase());

  const urlColIdx = indexHeaders.indexOf("FILE URL");
  const tabColIdx = indexHeaders.indexOf("TAB NAME");

  let allMatchedRecords = [];

  for (let i = offset; i < indexData.length; i++) {
    // Check execution time remaining
    if (Date.now() - START_TIME > TIME_LIMIT) {
      return { records: allMatchedRecords, nextOffset: i };
    }

    let row = indexData[i];
    let targetUrl = row[urlColIdx];
    let rawTabNames = row[tabColIdx] ? row[tabColIdx].toString() : "";

    if (!targetUrl || targetUrl.trim() === "") continue;

    try {
      let targetSs = SpreadsheetApp.openByUrl(targetUrl);
      let tz = targetSs.getSpreadsheetTimeZone(); // Dynamic timezone
      let tabNamesArray = rawTabNames.split(",").map(n => n.trim()).filter(n => n !== "");

      for (let tabName of tabNamesArray) {
        let targetSheet = targetSs.getSheetByName(tabName);
        if (!targetSheet) continue;

        let lastRow = targetSheet.getLastRow();
        let lastCol = targetSheet.getLastColumn();
        if (lastRow < 6) continue;

        // FETCH ENTIRE DATA RANGE FOR THIS SHEET ONCE (Minimize API Calls)
        // Optimization: Instead of row-by-row range calls, get everything starting row 5
        let sheetData = targetSheet.getRange(5, 1, lastRow - 4, lastCol).getValues();

        let fileHeaders = sheetData[0].map(h => h.toString().trim().toUpperCase());
        let rfpColIdx = fileHeaders.indexOf("RFP|PEF NO.");
        if (rfpColIdx === -1) continue;

        // Search the data (skipping the header row 0 of sheetData)
        for (let r = 1; r < sheetData.length; r++) {
          let cellValue = sheetData[r][rfpColIdx];

          if (cellValue && cellValue.toString().trim().toUpperCase() === cleanInput) {
            let record = {};
            let matchRow = sheetData[r];

            fileHeaders.forEach((h, idx) => {
              let val = matchRow[idx];
              // Safe check for blank headers
              let key = h || "COLUMN_" + idx;

              record[key] = (val instanceof Date)
                ? Utilities.formatDate(val, tz, "yyyy-MM-dd")
                : val;
            });

            allMatchedRecords.push(record);

            // If you only want the FIRST found match globally:
            // return { records: allMatchedRecords, nextOffset: null };
          }
        }
      }
    } catch (e) {
      console.warn("Skipping " + targetUrl + " due to error: " + e.message);
    }
  }

  return { records: allMatchedRecords, nextOffset: null };
}
/**
 * FORGOT PASSWORD - STAGE 1: Check account and send OTP
 * Triggered by: google.script.run.requestPasswordResetCode(email)
 */
function requestPasswordResetCode(email) {
  const sessionEmail = Session.getActiveUser().getEmail();
  if (!email || email !== sessionEmail)
    throw new Error(
      "Security Violation: You can only reset the password for your own authenticated account.",
    );

  const ss = SpreadsheetApp.openById(USER_DB_ID);
  const sheet = ss.getSheetByName(USER_TAB);
  const data = sheet.getDataRange().getValues();

  // Find index 2 (Column C - USERNAME/EMAIL)
  const accountExists = data.some((row) => row[2] === email);
  if (!accountExists)
    throw new Error("No registered account found for this email address.");

  const otp = Math.floor(100000 + Math.random() * 900000).toString();
  const cache = CacheService.getScriptCache();
  cache.put("RESET_" + email, otp, 600); // Unique cache key for resets

  try {
    MailApp.sendEmail(
      email,
      "Password Reset Code",
      "Your password reset verification code is: " + otp,
    );
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

  if (!storedCode)
    throw new Error("Reset code expired. Please request a new one.");
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
      throw new Error(
        "The new password must be different from your current password.",
      );
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
function processSubmission(p) {
  const lock = LockService.getPublicLock();
  const FOLDER_ID = "1eFcLGXPEnUSi14aPvvA9m2ATV8Racsuf";
  const SS_ID = "1YAvZmCdWXbjOcJA-uUY40e6qVqzyiHcB06NpiPcz6y4";

  try {
    // 1. CONCURRENCY LOCK (30s timeout)
    if (!lock.tryLock(30000)) {
      throw new Error(
        "Server is congested. Please wait 30 seconds and try again.",
      );
    }

    const ss = SpreadsheetApp.openById(SS_ID);
    const sh = ss.getSheetByName("SUBMISSIONS");

    // 2. SCALABILITY & CAPACITY GUARD
    const MAX_CELLS_LIMIT = 9500000;
    const currentMaxRows = sh.getMaxRows();
    const currentMaxCols = sh.getMaxColumns();
    const currentTotalCells = currentMaxRows * currentMaxCols;

    if (currentTotalCells >= MAX_CELLS_LIMIT) {
      throw new Error(
        "Storage Limit Alert: 9.5M cell capacity reached. Please archive data before next submission.",
      );
    }

    // 3. VALIDATIONS (Backend Safety)
    if (!p.header.rfpNo || !p.header.dueDate)
      throw new Error("RFP Number and Due Date are required.");
    if (!p.participants || p.participants.length === 0)
      throw new Error("At least one Transaction Participant is required.");
    if (!p.participants.some((item) => item.tag === "Primary"))
      throw new Error("At least one 'Primary' participant is required.");

    // 4. HIGH-SPEED DUPLICATE CHECK
    // TextFinder is significantly faster than loading the column into an array for large sheets
    const duplicate = sh
      .getRange("F:F")
      .createTextFinder(p.header.rfpNo)
      .matchEntireCell(true)
      .findNext();
    if (duplicate)
      throw new Error("Duplicate RFP: " + p.header.rfpNo + " already exists.");

    // 5. ASSET GENERATION
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const blob = Utilities.newBlob(
      Utilities.base64Decode(p.attachment.base64),
      "application/pdf",
      p.header.rfpNo + "_Support.pdf",
    );
    const attachmentUrl = folder.createFile(blob).getUrl();
    const formPdfUrl = createRfpPdf(p, folder);

    // 6. DATA PREPARATION
    const txnId = "TXN-" + Utilities.getUuid().split("-")[0].toUpperCase();
    const primaryEmails = p.participants
      .filter((i) => i.tag === "Primary")
      .map((i) => i.email)
      .join(", ");
    const secondaryEmails = p.participants
      .filter((i) => i.tag === "Secondary")
      .map((i) => i.email)
      .join(", ");

    const rowPrefix = [
      txnId,
      new Date(),
      Session.getActiveUser().getEmail(),
      primaryEmails,
      secondaryEmails,
      p.header.rfpNo,
      p.header.dueDate,
    ];

    const rowData = p.tableFields.map((f) => {
      const val = p.particulars[f];
      // Strict numeric parsing for Amount fields to ensure spreadsheet math works at scale
      return f.includes("AMOUNT")
        ? parseFloat(String(val).replace(/,/g, ""))
        : val;
    });

    // Construct the final flattened array for the row
    const finalRowData = [
      rowPrefix.concat(rowData).concat([attachmentUrl, formPdfUrl]),
    ];

    // 7. OPTIMIZED DIRECT-WRITE (Handling Millions of Rows)
    const lastRow = sh.getLastRow();

    // Auto-expand sheet if we are at the very bottom to prevent insertion errors
    if (lastRow === currentMaxRows) {
      sh.insertRowsAfter(currentMaxRows, 100); // Pre-emptively add 100 rows
    }

    // Write directly to the range (Faster than appendRow for massive sheets)
    const targetRange = sh.getRange(lastRow + 1, 1, 1, finalRowData[0].length);
    targetRange.setValues(finalRowData);

    // 8. FINAL COMMIT
    SpreadsheetApp.flush();

    return { success: true, message: p.header.rfpNo };
  } catch (e) {
    console.error("Critical Submission Error: " + e.message);
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

  data.forEach(row => {
    // Column D (Primary Emails string)
    if (row[0]) {
      row[0].split(",").forEach(e => {
        let clean = e.trim().toLowerCase();
        if (clean.includes("@")) emailSet.add(clean);
      });
    }
    // Column E (Secondary Emails string)
    if (row[1]) {
      row[1].split(",").forEach(e => {
        let clean = e.trim().toLowerCase();
        if (clean.includes("@")) emailSet.add(clean);
      });
    }
  });

  // Convert Set back to an Array and return
  return Array.from(emailSet).sort();
}
