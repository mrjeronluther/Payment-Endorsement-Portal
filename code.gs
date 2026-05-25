/**
 * GLOBAL CONFIGURATION
 */
const STORAGE_CONFIG = {
    PRIMARY_DB_ID: "1YAvZmCdWXbjOcJA-uUY40e6qVqzyiHcB06NpiPcz6y4",
    MASTER_LOG_ID: "1O8zbQ85U_Z4eskB0TYWNY1jZWRr6nFnVJTTy-d_EDF0",
    REGISTRY_ID: "1TI12SzvfLUtVXa1WbiklpY3KFOZOMvmSXfN3Zo18O7g",
    USER_DB_ID: "1dBO8ThI7FEKb24D9sPVWokfXLuWUx5aCQvisrT9wBvI",
    ARCHIVE_FOLDER_ID: "1hIMae05PAHR1bwE5VYE3heUQO71tey3Z",
    ATTACHMENT_FOLDER_ID: "1eFcLGXPEnUSi14aPvvA9m2ATV8Racsuf",
    DB_CELL_LIMIT: 8500000,
    MASTER_CELL_LIMIT: 8500000,
    REGISTRY_CELL_LIMIT: 8000000,
    MAX_PDF_SIZE_MB: 15,
    SECURITY_PEPPER: "v9_PEF_SYS_2027_!@#",
    USER_TAB: "PEP",
};
/**
 * UI INITIALIZATION
 */
function doGet(e) {
    return HtmlService.createTemplateFromFile("Index")
        .evaluate()
        .setTitle("Payment Endorsement Platform")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
        .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * REPLACEMENT FOR: generateSHA256
 * Single, secure salted hashing function used for Login, Reg, and Reset.
 */
function generateSHA256(input, salt) {
    const safeSalt = salt ? salt.trim().toLowerCase() : "SYSTEM_DEFAULT";
    const signature = Utilities.computeHmacSha256Signature(input.trim(), safeSalt + STORAGE_CONFIG.SECURITY_PEPPER);
    return signature.map((b) => ("0" + (b & 0xff).toString(16)).slice(-2)).join("");
}

/**
 * REPLACEMENT: replicateMasterLog
 * Handles cloning of Master system with relay tracking.
 */
function replicateMasterLog(fullId) {
    const anchorId = STORAGE_CONFIG.REGISTRY_ID;
    const anchorSs = SpreadsheetApp.openById(anchorId);
    const relaySheet = anchorSs.getSheetByName("MASTER") || anchorSs.insertSheet("MASTER");

    const folder = DriveApp.getFolderById(STORAGE_CONFIG.ARCHIVE_FOLDER_ID);
    const name = "MASTER_SHARD_" + Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd_HHmm");

    // Clone current master
    const clone = DriveApp.getFileById(fullId).makeCopy(name, folder);
    const newSs = SpreadsheetApp.openById(clone.getId());

    // Wipe data from tabs except GLOBAL_COUNTER
    newSs.getSheets().forEach((sh) => {
        if (sh.getName() === "GLOBAL_COUNTER") return;
        if (sh.getMaxRows() > 1) sh.getRange(2, 1, sh.getMaxRows() - 1, sh.getMaxColumns()).clearContent();
        if (sh.getMaxRows() > 100) sh.deleteRows(101, sh.getMaxRows() - 100);
    });

    // Record URL in the Registry Anchor
    relaySheet.appendRow([newSs.getUrl(), "LIVE", new Date()]);
    return newSs.getId();
}
/**
 * Mandatory Helper: Santize strings to prevent XSS and Template Injection
 */
function safeValue(t) {
    if (typeof t !== "string") return t || "";
    return t.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/**
 * PRIVATE HELPER: Safe ID Extractor
 * Prevents "Reading 0 of null" errors.
 */
function extractIdSafely(text, fallback) {
    if (!text || typeof text !== "string") return fallback;
    const match = text.match(/[-\w]{25,}/);
    return match ? match[0] : fallback;
}

/**
 * REPLACEMENT: getLiveMasterId
 * Looks in File 3 (Registry) for the Live Master URL.
 */
function getLiveMasterId() {
    const ss = SpreadsheetApp.openById(STORAGE_CONFIG.REGISTRY_ID);
    const sh = ss.getSheetByName("MASTER") || ss.insertSheet("MASTER").appendRow(["URL", "Status", "Date"]);

    const data = sh.getRange("A:A").getValues().filter(String).flat();

    // If data length is 1 or less, only the header exists. Use Anchor.
    if (data.length <= 1) return STORAGE_CONFIG.MASTER_LOG_ID;

    const latestUrl = data[data.length - 1];
    return extractIdSafely(latestUrl, STORAGE_CONFIG.MASTER_LOG_ID);
}

/**
 * REPLACEMENT: getLiveRegistryId
 * Looks in File 3 (Registry) for sharded versions of itself.
 */
function getLiveRegistryId() {
    const ss = SpreadsheetApp.openById(STORAGE_CONFIG.REGISTRY_ID);
    const sh =
        ss.getSheetByName("REGISTRY_RELAY") || ss.insertSheet("REGISTRY_RELAY").appendRow(["URL", "Status", "Date"]);

    const data = sh.getRange("A:A").getValues().filter(String).flat();

    if (data.length <= 1) return STORAGE_CONFIG.REGISTRY_ID;

    const latestUrl = data[data.length - 1];
    return extractIdSafely(latestUrl, STORAGE_CONFIG.REGISTRY_ID);
}

/**
 * REPLACEMENT: getActiveDatabaseId
 * Orchestrates Master capacity check and resolves Database ID.
 */
function getActiveDatabaseId() {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(30000);

        let masterId = getLiveMasterId();
        let masterSs = SpreadsheetApp.openById(masterId);

        // --- 1. Master Limit Check ---
        const mCells = masterSs.getSheets().reduce((sum, s) => sum + s.getMaxRows() * s.getMaxColumns(), 0);
        if (mCells >= STORAGE_CONFIG.MASTER_CELL_LIMIT) {
            masterId = replicateMasterLog(masterId);
            masterSs = SpreadsheetApp.openById(masterId);
        }

        // --- 2. Database Resolve ---
        const dbTab =
            masterSs.getSheetByName("DATABASES_ARCHIVES") ||
            masterSs.insertSheet("DATABASES_ARCHIVES").appendRow(["FILE URL"]);
        const urls = dbTab.getRange("A:A").getValues().filter(String).flat();

        // If only header, use the default primary ID
        if (urls.length <= 1) return STORAGE_CONFIG.PRIMARY_DB_ID;

        const latestDbUrl = urls[urls.length - 1];
        return extractIdSafely(latestDbUrl, STORAGE_CONFIG.PRIMARY_DB_ID);
    } finally {
        if (lock.hasLock()) lock.releaseLock();
    }
}

/**
 * ==========================================
 * REGISTRATION SYSTEM (ZERO-TRUST)
 * ==========================================
 */

function getActiveUserEmail() {
    return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
}

function replicateMaster(fullId) {
    const anchor = SpreadsheetApp.openById(STORAGE_CONFIG.MASTER_LOG_ID);
    const relay = anchor.getSheetByName("MASTER_RELAY_LOG");
    const folder = DriveApp.getFolderById(STORAGE_CONFIG.ARCHIVE_FOLDER_ID);

    const clone = DriveApp.getFileById(fullId).makeCopy(
        "PEP_MASTER_LIVE_" + Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd"),
        folder
    );
    const cloneSs = SpreadsheetApp.openById(clone.getId());

    // Wipe database links in the clone so it starts "light"
    const logTab = cloneSs.getSheetByName("DATABASES_ARCHIVES");
    if (logTab && logTab.getMaxRows() > 1)
        logTab.getRange(2, 1, logTab.getMaxRows() - 1, logTab.getMaxColumns()).clearContent();

    relay.appendRow([clone.getUrl(), "LIVE", new Date()]);
    return clone.getId();
}

/**
 * Standard Cloner: Clones, Wipes data, Updates Parent Log.
 */
function autoHealFile(oldId, masterLogSheet, type) {
    const folder = DriveApp.getFolderById(STORAGE_CONFIG.ARCHIVE_FOLDER_ID);
    const copy = DriveApp.getFileById(oldId).makeCopy(
        `PEP_${type}_SHARD_${Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd_HHmm")}`,
        folder
    );
    const ss = SpreadsheetApp.openById(copy.getId());

    ss.getSheets().forEach((sh) => {
        if (sh.getMaxRows() > 1) sh.getRange(2, 1, sh.getMaxRows() - 1, sh.getMaxColumns()).clearContent();
        if (sh.getMaxRows() > 100) sh.deleteRows(101, sh.getMaxRows() - 100); // Shave physically to keep fast
    });

    masterLogSheet.appendRow([copy.getUrl()]);
    return copy.getId();
}

/**
 * High-Volume Optimized Sharding
 */
function maintenanceEngine(targetId, parentId, parentTab, fixedCol, cellLimit) {
    const ss = SpreadsheetApp.openById(targetId);
    const totalCells = ss.getSheets().reduce((sum, s) => sum + s.getMaxRows() * s.getMaxColumns(), 0);

    if (totalCells >= cellLimit) {
        const folder = DriveApp.getFolderById(STORAGE_CONFIG.ARCHIVE_FOLDER_ID);
        const ts = Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd_HHmmss");

        // 1. Copy Data to Archive
        const backup = DriveApp.getFileById(targetId).makeCopy(`BACKUP_${ts}_${ss.getName()}`, folder);
        const url = backup.getUrl();

        // 2. Clear current file (preserving headers)
        ss.getSheets().forEach((sh) => {
            const lastRow = sh.getLastRow();
            if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, sh.getMaxColumns()).clearContent();

            // SHAVE FILE: Force shrink rows to speed up future writes
            if (sh.getMaxRows() > 100) sh.deleteRows(101, sh.getMaxRows() - 100);
        });

        // 3. Log to Parent
        const parentSs = SpreadsheetApp.openById(parentId);
        let parentSh = parentSs.getSheetByName(parentTab) || parentSs.insertSheet(parentTab);

        // Column logic: if File 1 archive, use next empty A/B. If File 2 archive, fixed Col 1.
        let col = fixedCol;
        if (!col) {
            const lastA = parentSh.getRange("A:A").getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
            const lastB = parentSh.getRange("B:B").getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
            col = lastA <= lastB ? 1 : 2;
        }

        parentSh.getRange(parentSh.getLastRow() + 1, col).setValue(url);
        SpreadsheetApp.flush();
    }
}

function maintainSystem(targetId, parentId, parentTab, fixedCol, cellLimit) {
    const ss = SpreadsheetApp.openById(targetId);
    const sheets = ss.getSheets();
    const totalCells = sheets.reduce((sum, s) => sum + s.getMaxRows() * s.getMaxColumns(), 0);

    if (totalCells >= cellLimit) {
        const folder = DriveApp.getFolderById(STORAGE_CONFIG.ARCHIVE_FOLDER_ID);
        const ts = Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd_HHmm");

        // 1. Snapshot the data
        const backup = DriveApp.getFileById(targetId).makeCopy(`BACKUP_${ts}_${ss.getName()}`, folder);
        const backupUrl = backup.getUrl();

        // 2. Self-Heal Target File (Empty rows but KEEP IDs and HEADERS)
        sheets.forEach((sh) => {
            const lastRow = sh.getLastRow();
            if (lastRow > 1) {
                sh.getRange(2, 1, lastRow - 1, sh.getMaxColumns()).clearContent();
            }
            // Speed Boost: Physically delete thousands of phantom rows to speed up writing.
            if (sh.getMaxRows() > 101) sh.deleteRows(102, sh.getMaxRows() - 101);
        });

        // 3. Log into Parent File
        const pSs = SpreadsheetApp.openById(parentId);
        let pSh = pSs.getSheetByName(parentTab) || pSs.insertSheet(parentTab);

        let targetCol = fixedCol;
        if (!targetCol) {
            // Determine Col A or Col B for File 1 sharding
            const lenA = pSh.getRange("A:A").getValues().filter(String).length;
            const lenB = pSh.getRange("B:B").getValues().filter(String).length;
            targetCol = lenA <= lenB ? 1 : 2;
        }

        pSh.getRange(pSh.getLastRow() + 1, targetCol).setValue(backupUrl);
        SpreadsheetApp.flush();
    }
}

/**
 * Internal Helper for Sharding
 */
function maintainTier(targetId, parentId, parentTab, colIndex, cellLimit) {
    const ss = SpreadsheetApp.openById(targetId);

    // Logic: Max Rows * Max Columns = Physical footprint in Google Servers
    const totalCells = ss.getSheets().reduce((sum, s) => sum + s.getMaxRows() * s.getMaxColumns(), 0);

    if (totalCells >= cellLimit) {
        console.log("Healing initiated for ID: " + targetId);

        // 1. MAKE ARCHIVE COPY
        const folder = DriveApp.getFolderById(STORAGE_CONFIG.ARCHIVE_FOLDER_ID);
        const timeStamp = Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd_HHmmss");
        const copy = DriveApp.getFileById(targetId).makeCopy("ARCHIVE_" + timeStamp + "_" + ss.getName(), folder);
        const archiveUrl = copy.getUrl();

        // 2. CLEAR ALL TABS IN CURRENT FILE (No Data remains, but Headers stay)
        ss.getSheets().forEach((sh) => {
            const lastRow = sh.getLastRow();
            if (lastRow > 1) {
                // Clears from Row 2 down to protect headers
                sh.getRange(2, 1, lastRow - 1, sh.getMaxColumns()).clearContent();
            }
            // CRITICAL: Shrink the sheet to 100 rows to reset the cell count officially
            if (sh.getMaxRows() > 100) {
                sh.deleteRows(101, sh.getMaxRows() - 100);
            }
        });

        // 3. LOG TO PARENT
        const parentSs = SpreadsheetApp.openById(parentId);
        let parentSh = parentSs.getSheetByName(parentTab);
        if (!parentSh) parentSh = parentSs.insertSheet(parentTab);

        // Handle Col A or Col B requirement
        const targetRow = parentSh.getLastRow() + 1;
        parentSh.getRange(targetRow, colIndex).setValue(archiveUrl);

        SpreadsheetApp.flush();
    }
}

/**
 * Main authentication function
 */
function authenticateUser(c) {
    const sh = SpreadsheetApp.openById(STORAGE_CONFIG.USER_DB_ID).getSheetByName(STORAGE_CONFIG.USER_TAB);
    const data = sh.getDataRange().getDisplayValues();
    const email = c.email.trim().toLowerCase();
    const hash = generateSHA256(c.password, email);
    const row = data.find((r) => r[2].toLowerCase() === email && r[3] === hash);

    if (!row) throw new Error("Invalid credentials.");
    if (row[6].toUpperCase() !== "APPROVED") throw new Error("Approval pending.");
    if (row[7].toUpperCase() === "RESIGNED") throw new Error("Account deactivated.");

    return { success: true, user: { fullName: row[1], email: row[2], role: row[4] } };
}

/**
 * Stage 1: Security Check & OTP
 * Logic: Checks if the account already exists before sending the code.
 */
function sendVerificationCode(email) {
    const sh = SpreadsheetApp.openById(STORAGE_CONFIG.USER_DB_ID).getSheetByName(STORAGE_CONFIG.USER_TAB);
    const emailExists = sh
        .getRange("C:C")
        .getValues()
        .some((r) => r[0].toString().toLowerCase() === email.toLowerCase());
    if (emailExists) throw new Error("Email already exists.");

    const otp = Math.floor(100000 + Math.random() * 900000).toString();
    CacheService.getScriptCache().put(email.trim().toLowerCase(), otp, 600);
    MailApp.sendEmail(email, "Registration OTP", "Code: " + otp);
    return { success: true };
}

/**
 * Stage 2: OTP Validation
 */
function verifyRegistrationCode(e, c) {
    const v = CacheService.getScriptCache().get(e.trim().toLowerCase());
    if (!v || v !== c) throw new Error("Invalid OTP.");
    return { success: true };
}

/**
 * Stage 3: Write Record with Gap-Filling
 * AUTOMATIC VALUES: USERNAME = Email, ACCESS LEVEL = Dynamic, STATUS = PENDING
 */
function finalizeRegistration(f) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(15000);
        const ss = SpreadsheetApp.openById(STORAGE_CONFIG.USER_DB_ID);
        const sh = ss.getSheetByName(STORAGE_CONFIG.USER_TAB);
        const activeUser = f.email.trim().toLowerCase();

        const rowPayload = [
            new Date(),
            safeValue(f.fullName),
            activeUser,
            generateSHA256(f.password, activeUser),
            f.department === "PCU" ? "REQUESTOR" : "APPROVER",
            activeUser.includes("@megaworld-lifestyle.com") ? "ORGANIC" : "NON ORGANIC",
            "Pending",
            "",
            "",
            safeValue(f.department),
            safeValue(f.jobDesignation),
        ];

        sh.appendRow(rowPayload);
        return { success: true };
    } finally {
        lock.releaseLock();
    }
}

/**
 * FORGOT PASSWORD - STAGE 1: Check account and send OTP
 * Triggered by: google.script.run.requestPasswordResetCode(email)
 */
function requestPasswordResetCode(email) {
    // Use STORAGE_CONFIG prefix!
    const ss = SpreadsheetApp.openById(STORAGE_CONFIG.USER_DB_ID);
    const sheet = ss.getSheetByName(STORAGE_CONFIG.USER_TAB);
    const data = sheet.getDataRange().getDisplayValues();

    // Clean email input
    const targetEmail = (email || "").toString().trim().toLowerCase();

    // Find index 2 (Column C - USERNAME/EMAIL)
    const accountExists = data.some((row) => row[2].toString().toLowerCase() === targetEmail);

    if (!accountExists) {
        throw new Error("No registered account found for: " + targetEmail);
    }

    const otp = Math.floor(100000 + Math.random() * 900000).toString();
    const cache = CacheService.getScriptCache();
    // Store code with "RESET_" prefix for security
    cache.put("RESET_" + targetEmail, otp, 600);

    try {
        MailApp.sendEmail(targetEmail, "Password Reset Code", "Your verification code is: " + otp);
        return { success: true, message: "Verification code sent." };
    } catch (e) {
        throw new Error("Failed to send email. Check app permissions.");
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
        const ss = SpreadsheetApp.openById(STORAGE_CONFIG.USER_DB_ID); // Fix prefix
        const sheet = ss.getSheetByName(STORAGE_CONFIG.USER_TAB);
        const data = sheet.getDataRange().getDisplayValues();

        let rowIndex = -1;
        const cleanEmail = email.toString().trim().toLowerCase();
        const hashedNewPassword = generateSHA256(newPassword, cleanEmail);

        for (let i = 0; i < data.length; i++) {
            if (data[i][2].toString().trim().toLowerCase() === cleanEmail) {
                rowIndex = i + 1;
                break;
            }
        }

        if (rowIndex === -1) throw new Error("Record no longer exists.");
        sheet.getRange(rowIndex, 4).setValue(hashedNewPassword);
        SpreadsheetApp.flush();
        return { success: true };
    } finally {
        lock.releaseLock();
    }
}

/**
 * REPLACEMENT: isRfpUnique
 * Checks the UNIQUE_REGISTRY tab inside the current Master Shard.
 */
function isRfpUnique(rfpNo) {
    if (!rfpNo || rfpNo === "" || rfpNo === "N/A") return true;
    const val = rfpNo.toString().trim().toUpperCase();

    // Tier 1: Anti-collision lock (Script cache)
    if (CacheService.getScriptCache().get("processing_" + val)) return false;

    // Tier 2: Check ONLY inside the Master Shard resolved from the relay
    const liveMasterId = getLiveMasterId();
    const masterSs = SpreadsheetApp.openById(liveMasterId);
    const sh = masterSs.getSheetByName("UNIQUE_REGISTRY");

    if (!sh) return true; // If no registry yet, it's unique

    // High speed search in the sharded file
    const finder = sh.getRange("A:A").createTextFinder(val).matchEntireCell(true).findNext();
    return finder ? false : true;
}

/**
 * REPLACEMENT: commitRfpToRegistry
 * Writes success identifiers ONLY into the sharded Master file,
 * NOT into the File 3 Registry Anchor.
 */
function commitRfpToRegistry(rfpNo) {
    if (!rfpNo || rfpNo === "" || rfpNo === "N/A") return;
    const idValue = rfpNo.toString().trim().toUpperCase();

    // Find where the Master currently is (from File 3 relay)
    const liveMasterId = getLiveMasterId();
    const masterSs = SpreadsheetApp.openById(liveMasterId);

    // Resolve UNIQUE_REGISTRY inside the SHARDED Master file
    let sh = masterSs.getSheetByName("UNIQUE_REGISTRY");
    if (!sh) {
        sh = masterSs.insertSheet("UNIQUE_REGISTRY").appendRow(["Identifier", "Timestamp"]);
    }

    sh.appendRow([idValue, new Date()]);

    // Clear the in-memory processing lock
    CacheService.getScriptCache().remove("processing_" + idValue);
    SpreadsheetApp.flush();
}
/**
 * ==========================================
 * RFP TRANSACTION & GENERATION (Spreadsheet)
 * ==========================================
 */
//var REGISTRY_SS_ID = "1YAvZmCdWXbjOcJA-uUY40e6qVqzyiHcB06NpiPcz6y4";
//var TAB_NAME = "autogenrfp";
function generateRfpNumber() {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(30000);
        const ss = SpreadsheetApp.openById(STORAGE_CONFIG.MASTER_LOG_ID);
        let counterSh = ss.getSheetByName("GLOBAL_COUNTER") || ss.insertSheet("GLOBAL_COUNTER");
        if (counterSh.getLastRow() === 0) counterSh.appendRow(["Sequence", 0]);
        const nextSeq = (parseInt(counterSh.getRange("B1").getValue()) || 0) + 1;
        counterSh.getRange("B1").setValue(nextSeq);
        return "MALL-" + Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM") + "-" + ("0000000" + nextSeq).slice(-7);
    } finally {
        lock.releaseLock();
    }
}

/**
 * REPLACEMENT FOR: processSubmission (Lines 492-563)
 * Reordered for Transaction Integrity: Heavy logic first, Data Write last.
 */
/**
 * PROCESS SUBMISSION - FULL VERSION
 * Implements:
 * 1. App User Identity Tracking (Payload-based)
 * 2. Mandatory Attachment Notes Validation
 * 3. 500-character Remarks Logic
 * 4. Multi-layered Transaction Security (Locks, Cache, Sharding)
 */
function processSubmission(p) {
    const lock = LockService.getScriptLock();

    // SECURE LOCK: Prevent concurrent write collisions (Wait up to 45 seconds)
    if (!lock.tryLock(45000)) {
        throw new Error(
            "SYSTEM_BUSY: The database is currently being updated by another user. Please try again in a moment."
        );
    }

    try {
        // 1. DATA INITIALIZATION & IDENTITY SECURITY
        const rfpNo = p.header.rfpNo ? p.header.rfpNo.toString().trim().toUpperCase() : "";

        // IDENTITY: Use the email from the App Login state, not the Google account session
        const submissionEmail = p.userAppEmail ? p.userAppEmail.toLowerCase().trim() : "anonymous_app_user";

        // REMARKS: Truncate to 500 characters max for database safety
        let remarks = p.header.remarks ? p.header.remarks.toString().trim() : "";
        if (remarks.length > 500) remarks = remarks.substring(0, 500);

        // If user didn't pick a source file (manual generation), we force "N/A"
        const sourceFileName =
            p.header.sourceFileName && p.header.sourceFileName.trim() !== "" ? p.header.sourceFileName.trim() : "N/A";

        const sourceTabName =
            p.header.sourceTabName && p.header.sourceTabName.trim() !== "" ? p.header.sourceTabName.trim() : "N/A";

        // 2. TRANSACTIONAL SECURITY (Debouncing & Duplicates)
        if (rfpNo !== "" && !isRfpUnique(rfpNo)) {
            throw new Error("ALREADY RECORDED: This RFP No. (" + rfpNo + ") is already in the registry.");
        }

        if (rfpNo !== "") {
            const cache = CacheService.getScriptCache();
            if (cache.get("processing_" + rfpNo)) {
                throw new Error("SUBMISSION_IN_PROGRESS: This transaction is already being processed.");
            }
            cache.put("processing_" + rfpNo, "true", 120); // Debounce for 2 minutes
        }

        const activeDbId = getActiveDatabaseId(); // Retrieve ID from configuration or sharding logic

        // 3. FILE ATTACHMENTS (Enforcing Mandatory Notes)
        const folder = DriveApp.getFolderById(STORAGE_CONFIG.ATTACHMENT_FOLDER_ID);
        let attachmentSummary = [];

        if (!p.attachments || p.attachments.length === 0) {
            throw new Error("ATTACHMENT_REQUIRED: At least one PDF supporting document is required.");
        }

        p.attachments.forEach((f, i) => {
            // MANDATORY NOTE CHECK (Backend enforcement)
            const fileNote = f.notes ? f.notes.trim() : "";
            if (!fileNote) {
                throw new Error(`MISSING_NOTE: File #${i + 1} (${f.name}) is missing the required description/note.`);
            }

            // Uploading and setting permissions
            const blob = Utilities.newBlob(
                Utilities.base64Decode(f.base64),
                "application/pdf",
                `${rfpNo}_FILE_${i + 1}.pdf`
            );
            const file = folder.createFile(blob);
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.NONE);

            // Formatting the cell content: [NOTE] : URL
            attachmentSummary.push(`[${fileNote.toUpperCase()}] : ${file.getUrl()}`);
        });

        // 4. GENERATE SYSTEM PDF (The Official RFP Snapshot)
        const snapshotUrl = createRfpPdf(p, folder);

        // 5. DATA PREPARATION (Particulars Mapping & Type Conversion)
        const sh =
            SpreadsheetApp.openById(activeDbId).getSheetByName("SUBMISSIONS") ||
            SpreadsheetApp.openById(activeDbId).insertSheet("SUBMISSIONS");

        let parts = {};
        p.tableFields.forEach((f) => {
            let v = p.particulars[f] || "";
            // Financial Sanitization: Remove formatting for clean numbers in DB
            parts[f] = f.toUpperCase().includes("AMOUNT")
                ? parseFloat(v.toString().replace(/[^0-9.-]/g, "")) || 0
                : v === null || v === undefined
                  ? ""
                  : v.toString().trim();
        });

        const primaryRecipients = p.participants
            .filter((x) => x.tag === "Primary")
            .map((x) => x.email)
            .join(", ");
        const secondaryRecipients = p.participants
            .filter((x) => x.tag === "Secondary")
            .map((x) => x.email)
            .join(", ");

        // 6. FINAL ROW MAPPING (Matches your database column architecture)
        const rowData = [
            "TXN-" + Utilities.getUuid().split("-")[0].toUpperCase(), // Col 1: System TXN ID
            new Date(), // Col 2: Processing Date
            submissionEmail, // Col 3: APP LOGGED-IN USER (Updated)
            primaryRecipients, // Col 4: Primary To
            secondaryRecipients, // Col 5: CC Emails
            rfpNo || "N/A", // Col 6: RFP Number
            p.header.dueDate, // Col 7: Deadline
            parts["YEAR"], // Col 8
            parts["MONTH"], // Col 9
            parts["PAYOR NAME"], // Col 10
            parts["PAYEE NAME"], // Col 11
            parts["PROPERTY"], // Col 12
            parts["LOCATION"], // Col 13
            parts["SECTOR"], // Col 14
            parts["KINDS OF SERVICE"], // Col 15
            parts["CONTRACT NO"], // Col 16
            parts["CONTRACT AMOUNT"], // Col 17 (Clean Number)
            p.header.invoiceNo, // Col 18
            parts["BILLING PERIOD"], // Col 19
            parts["SOA AMOUNT"], // Col 20 (Clean Number)
            parts["GENERAL STATUS"], // Col 21
            attachmentSummary.join("\n"), // Col 22: Linked PDF Array with Mandatory Notes
            snapshotUrl, // Col 23: PDF Receipt Link
            remarks, // Col 24: PURPOSE / REMARKS (New 500 Char)
            sourceFileName, // Uses the sanitized variable (Col 25)
            sourceTabName, // Uses the sanitized variable (Col 26)
        ];

        // Commit to Google Sheets
        sh.appendRow(rowData);

        // Trigger downstream processes (Registry update & caching)
        commitRfpToRegistry(rfpNo);

        // Ensure all data is written before releasing the script lock
        SpreadsheetApp.flush();

        return { success: true, message: rfpNo || "Submission Successful" };
    } catch (err) {
        console.error("Critical ProcessSubmission Error: " + err.message);
        return { success: false, message: err.message || err.toString() };
    } finally {
        // Always release the lock, even if the script failed
        if (lock.hasLock()) lock.releaseLock();
    }
}
/**
 * UPDATED: Fetches Filename, Tab, and URL
 */
function getSourceFileNames() {
    const data = SpreadsheetApp.openById("10p-nv_qAN0GzZHAVInyb9_bprk2_sLf08190qdMf8Mc")
        .getSheetByName("SOURCEFILES")
        .getDataRange()
        .getValues();
    data.shift();
    return data.map((r) => ({ fileName: r[0], tabName: r[1], fileUrl: r[2] })).filter((x) => x.fileName);
}

/**
 * Unified Advanced Search used by the "Search" button in UI
 */

function getRfpDataAdvanced(source, rfp, inv) {
    // Use our search engine to find the record
    const results = performStrictSearch(source, rfp, inv);

    if (!results || results.length === 0) {
        return { status: "NOT_FOUND" };
    }

    // Identify the match
    const firstMatch = results[0];

    // LOGIC FIX: Normalize the key name in case of spacing issues in the spreadsheet header
    const statusKey = Object.keys(firstMatch).find((key) => key.includes("GENERAL STATUS"));
    const statusValue = statusKey ? String(firstMatch[statusKey]).trim().toUpperCase() : "";

    // If Status is PAID, block the fetch immediately
    if (statusValue === "PAID") {
        const searchId = (rfp || inv).trim();
        return {
            status: "ALREADY_PAID",
            message: `Action Blocked: Record ${searchId} is already tagged as PAID in the source file.`,
        };
    }

    // If not paid, allow the fetch to proceed
    return { status: "FOUND", data: results };
}

function performStrictSearch(selectedSource, rfpInput, invoiceInput) {
    const isRfpSearch = !!rfpInput;
    const term = (rfpInput || invoiceInput || "").toString().trim();
    if (!term) return [];

    const indexSs = SpreadsheetApp.openById("10p-nv_qAN0GzZHAVInyb9_bprk2_sLf08190qdMf8Mc");
    const indexData = indexSs.getSheetByName("SOURCEFILES").getDataRange().getValues();
    const meta = indexData.find((r) => r[0].toString().trim().toUpperCase() === selectedSource.toUpperCase());
    if (!meta) return [];

    const targetSs = SpreadsheetApp.openByUrl(meta[2]);
    const tabs = meta[1]
        .toString()
        .split(",")
        .map((t) => t.trim());
    let matches = [];

    // Identify which header text we are looking for based on user input
    const targetHeaderName = isRfpSearch ? "RFP|PEF NO." : "INVOICE NO.";

    tabs.forEach((tabName) => {
        const sh = targetSs.getSheetByName(tabName);
        if (!sh) return;

        // 1. Map headers from Row 5 to find the target column index
        const rawHeaders = sh.getRange(5, 1, 1, sh.getLastColumn()).getValues()[0];
        const headers = rawHeaders.map((h) => h.toString().trim().toUpperCase());

        // Find index of the column we want to search (1-based index)
        const colIndex = headers.indexOf(targetHeaderName.toUpperCase()) + 1;

        // Safety check: If header is not found in Row 5, try a fuzzy match without the period
        let finalColIndex = colIndex;
        if (finalColIndex === 0) {
            finalColIndex = headers.findIndex((h) => h.includes(isRfpSearch ? "RFP" : "INVOICE")) + 1;
        }

        if (finalColIndex === 0) {
            console.warn(`Could not find ${targetHeaderName} in tab ${tabName}`);
            return;
        }

        // 2. SEARCH ONLY IN THAT SPECIFIC COLUMN (Row 6 down)
        const searchRange = sh.getRange(6, finalColIndex, sh.getLastRow() > 5 ? sh.getLastRow() - 5 : 1, 1);
        const finder = searchRange.createTextFinder(term).matchCase(false).matchEntireCell(true).findAll();

        if (!finder.length) return;

        // 3. Process results for matches found ONLY in that column
        finder.forEach((res) => {
            const row = res.getRow();
            const rowVal = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];

            let record = {};
            headers.forEach((h, i) => {
                let v = rowVal[i];
                const key = h || "COLUMN_" + (i + 1);
                if (v instanceof Date) {
                    record[key] = Utilities.formatDate(v, "GMT+8", "yyyy-MM-dd");
                } else {
                    record[key] = v;
                }
            });

            record["META_TAB_NAME"] = tabName;
            matches.push(record);
        });
    });

    return matches;
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
        logoBase64 =
            "data:" + logoFile.getMimeType() + ";base64," + Utilities.base64Encode(logoFile.getBlob().getBytes());
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
 * REPLACEMENT: getPreviousParticipantEmails
 * Loophole Fix: Scans archives listed in MASTER_LOG so suggestions don't "reset."
 */
function getPreviousParticipantEmails() {
    const masterLog = SpreadsheetApp.openById(STORAGE_CONFIG.MASTER_LOG_ID).getSheetByName("DATABASES_ARCHIVES");
    const results = new Set();

    let scanList = [STORAGE_CONFIG.PRIMARY_DB_ID];
    if (masterLog) {
        const archives = masterLog.getRange("A:A").getValues().filter(String).flat().slice(-3); // Get latest 3 archive shards
        archives.forEach((url) => {
            const m = url.match(/[-\w]{25,}/);
            if (m) scanList.push(m[0]);
        });
    }

    scanList.forEach((id) => {
        try {
            const sh = SpreadsheetApp.openById(id).getSheetByName("SUBMISSIONS");
            if (!sh) return;
            const data = sh.getRange(2, 4, Math.min(sh.getLastRow() - 1, 300), 2).getValues();
            data.forEach((r) =>
                r.forEach((cell) => {
                    if (cell && cell.includes("@"))
                        cell.split(",").forEach((em) => results.add(em.trim().toLowerCase()));
                })
            );
        } catch (e) {}
    });

    return Array.from(results).sort();
}

/**
 * Frontend: addSecondaryRecipient
 * Blocks addition if already in Primary list or invalid.
 */
function addSecondaryRecipient(emailToVerify) {
    const cleanEmail = emailToVerify.trim().toLowerCase();

    if (!cleanEmail.includes("@") || !cleanEmail.includes(".")) {
        Swal.fire(
            "Invalid Format",
            "The entry '" + cleanEmail + "' does not appear to be a valid email address.",
            "error"
        );
        return;
    }

    // Identify overlaps
    const isPrimary = formState.participants.some((p) => p.email.toLowerCase() === cleanEmail && p.tag === "Primary");

    if (isPrimary) {
        Swal.fire({
            title: "Duplicate Found",
            text: "The person " + cleanEmail + " is already assigned as a Primary Recipient.",
            icon: "warning",
        });
        return;
    }

    // Prevent UI double-entry in the same category
    if (formState.participants.some((p) => p.email.toLowerCase() === cleanEmail && p.tag === "Secondary")) return;

    // Function from your existing logic to add pill to UI
    addParticipant(cleanEmail, "Secondary");
}

/**
 * Fetches standardized file types. 
 * If tab is missing, returns empty to trigger frontend blocking.
 */
function getFileTypeLabels() {
    try {
        const targetSsId = "1FIfBuwWvdj4M1UkcAw9jtnRdNPkk5tKgIotD7K_MHxE";
        const ss = SpreadsheetApp.openById(targetSsId);
        const sh = ss.getSheetByName("FILE NOTES");

        if (!sh) {
            console.error("Critical: 'FILE NOTES' tab not found.");
            return null; // Return null so frontend knows configuration is broken
        }

        const lastRow = sh.getLastRow();
        if (lastRow < 2) return [];

        return sh.getRange(2, 1, lastRow - 1, 1).getValues()
                 .flat()
                 .filter(String)
                 .map(label => label.toString().trim());
    } catch (e) {
        return null;
    }
}

function emergencyClearCache() {
    const cache = CacheService.getScriptCache();
    // Add the IDs of the files that keep sharding here
    cache.remove("shard_full_1YAvZmCdWXbjOcJA-uUY40e6qVqzyiHcB06NpiPcz6y4");
    console.log("Cache cleared. Try submitting now.");
}
