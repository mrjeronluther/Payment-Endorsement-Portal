/**
 * ============================================================================
 * HIGH-PERFORMANCE GLOBAL CONFIGURATION
 * ============================================================================
 */
const VAR = {
     SOURCE_MAPPING_ID: "1O4kM0mBBMBjSzgwNJqb0eMiSUYshvPDYLrBPPWq42s8",
     SOURCE_MAPPING_TAB: "SourceFile",

     FILE_1_ID: "1cUtvzb_pDvtppKfU7SnAra6IOv2V1CqAsA6TBflmaVM", // Backup DB
     FILE_1_TAB: "ActionLogs",
     FILE_2_ID: "1eSWAWG6Rgc_etdLyEeb9u4yDMVGonp6RjWITTwj5ho8", // Master DB
     FILE_2_TAB: "Sample",

     COL_LIMIT: 16,
     LOCK_TIMEOUT: 30000,
     MAPPING: [
          "YEAR",
          "MONTH",
          "PAYOR NAME",
          "PAYEE NAME",
          "PROPERTY",
          "LOCATION",
          "SECTOR",
          "KINDS OF SERVICE",
          "RFP NUMBER",
          "CONTRACT NO",
          "CONTRACT AMOUNT",
          "INVOICE NO.",
          "BILLING PERIOD",
          "SOA AMOUNT",
          "GENERAL STATUS",
          "REMARKS",
     ],
     REVISED_COL_NAME: "Action Logs", // Column for history
};

const Security = {
     sanitize(val) {
          if (typeof val !== "string") return val;
          const sanitized = val.trim();
          return /^[=\+\-\@\t\r]/.test(sanitized) ? `'${sanitized}` : sanitized;
     },
};

function doGet() {
     return HtmlService.createTemplateFromFile("Index")
          .evaluate()
          .setTitle("PEP System")
          .addMetaTag("viewport", "width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0")
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function getSourceFiles() {
     try {
          const ss = SpreadsheetApp.openById(VAR.SOURCE_MAPPING_ID);
          const sheet = ss.getSheetByName(VAR.SOURCE_MAPPING_TAB);
          if (!sheet) return [];
          const data = sheet.getDataRange().getValues();
          return data
               .slice(1)
               .filter((r) => r[0] && r[1])
               .map((r) => ({
                    name: String(r[0]).trim(),
                    id: r[1].match(/[-\w]{25,}/) ? r[1].match(/[-\w]{25,}/)[0] : r[1],
                    tab: String(r[2] || "").trim(),
               }));
     } catch (e) {
          return [];
     }
}

// 1. CLEANS DATA COMING FROM CLIENT BEFORE SAVING TO SHEET
// Converts "1,500.00" -> 1500 (Number) so formulas work
function cleanForDB(val) {
     if (typeof val !== "string") return val;
     if (val.trim() === "") return "";

     // Check if it looks like a number with commas (e.g. "1,234.56" or "-500")
     // Regex: Optional minus, digits, optional commas, optional decimals
     if (/^-?[\d,]+(\.\d+)?$/.test(val.trim())) {
          const raw = val.replace(/,/g, ""); // Remove commas
          if (!isNaN(raw) && isFinite(raw)) {
               return Number(raw); // Return actual Number type
          }
     }
     return val.trim(); // Return text otherwise
}

// 2. FORMATS DATA GOING TO CLIENT (LOOKUP)
// Converts Date Objects -> "YYYY-MM-DD" string
function formatForClient(val) {
     if (val === null || val === undefined) return "";
     if (val instanceof Date) {
          return Utilities.formatDate(val, "GMT+8", "yyyy-MM-dd");
     }
     return String(val).trim();
}
/**
 * UTILITY: Normalizes any value into a clean, searchable String.
 * Prevents 12345 vs "12345" mismatch.
 */
function normalize(val) {
     if (val === null || val === undefined) return "";
     return String(val).trim();
}

/**
 * 1. IMPROVED LOOKUP: Optimized Sniper Mode
 * This avoids loading all rows. It only touches the specific row found.
 */
/**
 * SNIPER LOOKUP: Fetches contract details (including Source Status)
 * and performs background Master DB security checks.
 */
function lookupContractData(contractNo, sourceId, sourceTabString) {
     try {
          if (!sourceId || !sourceTabString) throw new Error("Selection Required.");

          const searchVal = String(contractNo).trim();
          if (searchVal.length < 3) return { success: false };

          // --- PART 1: SOURCE FILE LOOKUP (Contract Details) ---
          const ssSource = SpreadsheetApp.openById(sourceId);
          const tabNames = sourceTabString.split(",").map((name) => name.trim());

          let foundCell = null;
          for (const tabName of tabNames) {
               const sheet = ssSource.getSheetByName(tabName);
               if (!sheet) continue;
               const finder = sheet.createTextFinder(searchVal).matchEntireCell(true).findNext();
               if (finder) {
                    foundCell = finder;
                    break;
               }
          }

          if (!foundCell) {
               return { success: false, error: "NOT_FOUND", message: `RFP '${searchVal}' not found in registry.` };
          }

          const sheetFound = foundCell.getSheet();
          const rowIdx = foundCell.getRow();
          const lastCol = sheetFound.getLastColumn();
          const headers = sheetFound.getRange(1, 1, 1, lastCol).getValues()[0];
          const rowValues = sheetFound.getRange(rowIdx, 1, 1, lastCol).getValues()[0];
          const normalizedHeaders = headers.map((h) => String(h).trim().toUpperCase());

          // Map indices 0 to 14 from the Source File (Includes Source General Status)
          const mappedData = VAR.MAPPING.map((header) => {
               const target = header.toUpperCase();
               let idx = normalizedHeaders.indexOf(target);
               if (idx === -1) idx = normalizedHeaders.findIndex((h) => target.includes(h) || h.includes(target));
               return idx > -1 ? formatForClient(rowValues[idx]) : "";
          });

          // --- PART 2: BACKGROUND SECURITY CHECK (Master DB Bottom-Up) ---
          let validation = { type: "VALID", message: "" };
          const ssMaster = SpreadsheetApp.openById(VAR.FILE_2_ID);
          const shMaster = ssMaster.getSheetByName(VAR.FILE_2_TAB);

          // Find ALL instances of this RFP to check history
          const results = shMaster.createTextFinder(searchVal).matchEntireCell(true).findAll();

          if (results.length > 0) {
               const latestMatch = results[results.length - 1];
               const mRowIdx = latestMatch.getRow();
               const mHeaders = shMaster
                    .getRange(1, 1, 1, shMaster.getLastColumn())
                    .getValues()[0]
                    .map((h) => h.toString().toUpperCase());
               const statusColIdx = mHeaders.indexOf("RBG STATUS (LATEST)");

               if (statusColIdx !== -1) {
                    const latestStatus = String(shMaster.getRange(mRowIdx, statusColIdx + 1).getValue())
                         .trim()
                         .toUpperCase();

                    if (latestStatus === "PAID") {
                         validation.type = "BLOCK";
                         // MESSAGE REMOVED: No longer shows the status string
                         validation.message = "SECURITY BLOCK: This RFP is locked for re-endorsement.";
                    } else if (latestStatus.includes("RETURN")) {
                         validation.type = "MODAL_WARN";
                         // MESSAGE REMOVED: No longer shows the status string
                         validation.message =
                              "ATTENTION: This RFP was previously returned. Action Taken Note required on submission.";
                    } else {
                         validation.type = "TOAST_INFO";
                         validation.message = "RFP record found!";
                    }
               }
          }

          return { success: true, data: mappedData, validation: validation };
     } catch (e) {
          return { success: false, error: "SYS_ERR", message: e.toString() };
     }
}

/**
 * STRICT NUMBER ENFORCER
 * Forces a value to be a Javascript Number.
 * "1,500.00" -> 1500
 * "PHP 500"  -> 500
 * "Free"     -> 0
 */
function forceNumber(val) {
     if (val === null || val === undefined) return 0;

     // Convert to string, strip everything except digits, dots, and minus
     let cleanString = String(val).replace(/[^0-9.-]/g, "");

     // Parse
     let number = parseFloat(cleanString);

     // If result is NaN (Not a Number), default to 0.00
     return isNaN(number) ? 0 : number;
}

/**
 * Backend Submission Logic
 */
function submitData(gridData, rowCommentsMap, confirmedOverwrites = [], sourceId, sourceTab) {
     const lock = LockService.getScriptLock();
     const finishedRFPs = [];
     const heldRFPs = [];

     try {
          lock.waitLock(30000);

          // Forces exact format and string type (Jan 29, 2026)
          const ts = "'" + Utilities.formatDate(new Date(), "GMT+8", "MMM d, yyyy");

          const validRows = gridData
               .map((r, index) => ({ data: r, originalIndex: index }))
               .filter((item) => String(item.data[8] || "").trim() !== "");

          if (validRows.length === 0) {
               return { success: false, message: "No valid data found." };
          }

          const ssMaster = SpreadsheetApp.openById(VAR.FILE_2_ID);
          const shMaster = ssMaster.getSheetByName(VAR.FILE_2_TAB);
          const ssLogs = SpreadsheetApp.openById(VAR.FILE_1_ID);
          let shLogs = ssLogs.getSheetByName("ActionLogs");

          if (!shLogs) {
               shLogs = ssLogs.insertSheet("ActionLogs");
               const masterHeaders = shMaster.getRange(1, 1, 1, shMaster.getLastColumn()).getValues();
               shLogs.getRange(1, 1, 1, masterHeaders[0].length).setValues(masterHeaders);
          }

          const mHeaders = shMaster
               .getRange(1, 1, 1, shMaster.getLastColumn())
               .getValues()[0]
               .map((h) => String(h).trim().toUpperCase());
          const statusIdxM = mHeaders.indexOf("RBG STATUS (LATEST)");
          const logIdxM = mHeaders.indexOf(VAR.REVISED_COL_NAME.toUpperCase());
          const rbgCmtIdxM = mHeaders.indexOf("LATEST COMMENT OF RBG");
          const dateStatIdxM = mHeaders.indexOf("DATE OF LATEST STATUS");

          const sourceUrl = sourceId ? "https://docs.google.com/spreadsheets/d/" + sourceId : "N/A";
          const sourceTabName = sourceTab || "N/A";

          for (let i = 0; i < validRows.length; i++) {
               const item = validRows[i];
               const row = item.data;
               const rfp = String(row[8]).trim();
               const userNoteFromModal = rowCommentsMap[rfp];

               let cleanData = row.map((cell, idx) => {
                    if (idx === 10 || idx === 13) return forceNumber(cell);
                    return cleanForDB(cell);
               });

               cleanData[16] = ts;
               cleanData[17] = sourceUrl;
               cleanData[18] = sourceTabName;

               const allMatches = shMaster.createTextFinder(rfp).matchEntireCell(true).findAll();
               let rowIdxInMaster = -1;
               let existingRowData = [];
               let currentStatus = "NEW";

               if (allMatches.length > 0) {
                    const latestMatchCell = allMatches[allMatches.length - 1];
                    rowIdxInMaster = latestMatchCell.getRow();
                    existingRowData = shMaster.getRange(rowIdxInMaster, 1, 1, mHeaders.length).getValues()[0];

                    currentStatus = String(existingRowData[statusIdxM] || "").trim();
                    const statusUpper = currentStatus.toUpperCase();

                    if (statusUpper === "PAID") {
                         heldRFPs.push(rfp);
                         continue;
                    }

                    // Guard: Cancelled Logic
                    if (statusUpper === "CANCELLED") {
                         if (!confirmedOverwrites.includes(rfp)) {
                              // Stop and ask user if they want to proceed as NEW
                              return {
                                   success: false,
                                   actionRequired: "CANCELLED_NEW",
                                   contractNo: rfp,
                                   finishedRFPs,
                                   heldRFPs,
                                   message: `RFP ${rfp} was previously CANCELLED. Proceed as a NEW submission?`,
                              };
                         } else {
                              // USER CONFIRMED: Reset index to -1 so it triggers "CASE: NEW SUBMISSION" below
                              rowIdxInMaster = -1;
                         }
                    }

                    if (statusUpper.includes("RETURN") && !userNoteFromModal) {
                         return {
                              success: false,
                              actionRequired: "NOTE",
                              contractNo: rfp,
                              finishedRFPs,
                              heldRFPs,
                              message: `RFP ${rfp} is ${currentStatus}. Note required.`,
                         };
                    }

                    if (
                         ["WITH ACCTG", "PENDING ORIGINAL DOCS/ HARD COPY"].includes(statusUpper) &&
                         !confirmedOverwrites.includes(rfp)
                    ) {
                         return {
                              success: false,
                              actionRequired: "OVERWRITE",
                              contractNo: rfp,
                              finishedRFPs,
                              heldRFPs,
                              message: `RFP ${rfp} is currently [${currentStatus}]. Overwrite?`,
                         };
                    }
               }

               // EXECUTE WRITE
               if (rowIdxInMaster !== -1) {
                    // Slice(0,17) saves column index 0 through 16 (The Timestamp)
                    shMaster.getRange(rowIdxInMaster, 1, 1, 17).setValues([cleanData.slice(0, 17)]);

                    if (logIdxM !== -1) {
                         const prevLogs = String(existingRowData[logIdxM] || "");
                         const statusBefore = String(existingRowData[statusIdxM] || "NEW").trim();

                         const dateBeforeRaw = existingRowData[dateStatIdxM];
                         const dateBeforeStr =
                              dateBeforeRaw instanceof Date && !isNaN(dateBeforeRaw.getTime())
                                   ? Utilities.formatDate(dateBeforeRaw, "GMT+8", "MMM d, yyyy")
                                   : dateBeforeRaw
                                     ? String(dateBeforeRaw)
                                     : "";

                         const rawCmt = existingRowData[rbgCmtIdxM];
                         let commentBefore =
                              rawCmt instanceof Date
                                   ? Utilities.formatDate(rawCmt, "GMT+8", "MMM d, yyyy")
                                   : String(rawCmt || "").trim();

                         let logEntry = `${ts} | ${statusBefore}`;
                         if (dateBeforeStr) logEntry += ` (${dateBeforeStr})`;
                         if (commentBefore && commentBefore.toUpperCase() !== "N/A") logEntry += ` / ${commentBefore}`;
                         logEntry += ` -> WITH ACCTG`;
                         if (userNoteFromModal) logEntry += ` | Note: ${userNoteFromModal}`;

                         shMaster
                              .getRange(rowIdxInMaster, logIdxM + 1)
                              .setValue(prevLogs ? logEntry + "\n" + prevLogs : logEntry);
                    }
                    if (statusIdxM !== -1) shMaster.getRange(rowIdxInMaster, statusIdxM + 1).setValue("WITH ACCTG");
               } else {
                    shMaster.appendRow(cleanData.slice(0, 19));
                    const newIdx = shMaster.getLastRow();
                    if (statusIdxM !== -1) shMaster.getRange(newIdx, statusIdxM + 1).setValue("WITH ACCTG");
                    if (logIdxM !== -1) shMaster.getRange(newIdx, logIdxM + 1).setValue(`${ts} | Initial Submission`);
               }

               shLogs.appendRow(cleanData.slice(0, 19));
               finishedRFPs.push(rfp);
          }

          // Cleanup Source Tab logic remains same
          const ssSource = SpreadsheetApp.getActiveSpreadsheet();
          const shSource = ssSource.getSheetByName(sourceTab);
          if (shSource && shSource.getLastRow() > 1) {
               shSource.getRange(2, 1, shSource.getLastRow() - 1, shSource.getLastColumn()).clearContent();
          }

          return { success: true, finishedRFPs, heldRFPs, message: "Submission Successful." };
     } catch (e) {
          console.error(e);
          return { success: false, message: e.toString(), finishedRFPs, heldRFPs };
     } finally {
          lock.releaseLock();
     }
}

/**
 * Helper to ensure numeric columns don't break
 */
function forceNumber(val) {
     if (val === "" || val === null || val === undefined) return "";
     const num = Number(val.toString().replace(/[^0-9.-]+/g, ""));
     return isNaN(num) ? val : num;
}

/**
 * Helper to clean strings for DB entry
 */
function cleanForDB(val) {
     if (val === null || val === undefined) return "";
     if (val instanceof Date) return Utilities.formatDate(val, "GMT+8", "MMM d, yyyy");
     return String(val).trim();
}

function formatForClient(val) {
     if (val instanceof Date) return Utilities.formatDate(val, "GMT+8", "yyyy-MM-dd");
     return val === null || val === undefined ? "" : String(val).trim();
}
