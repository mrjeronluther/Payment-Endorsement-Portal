# Payment Endorsement Portal

Table of contents
- Project Overview
- Repo Structure
- Quick Start
- Prerequisites
- Configuration (spreadsheet IDs, folder IDs, etc.)
- Deploying to Google Apps Script (Web App)
- Frontend / Backend integration
- Exposed backend functions (API surface)
- UI walkthrough (major views & flows)
- Security considerations and recommendations
- Troubleshooting & common errors
- Development & deployment tips
- Contributing & license

---

## Project Overview
The Payment Endorsement Portal (PEP) is a web application for preparing, generating, and submitting Request For Payment (RFP) / PEF entries, built as a Google Apps Script web app (backend) with a single-file Vue.js frontend contained in an HTML file. The backend uses Google Sheets as a persistence layer, Drive for file storage, and MailApp for OTP and notification emails.

Key capabilities:
- User registration (OTP via email) and zero-trust registration flow.
- Authentication against a users spreadsheet.
- Generate sequential RFP numbers (thread-safe using LockService).
- Advanced RFP/Invoice lookup across source spreadsheets.
- File upload (PDF attachment) and persistence to Drive.
- RFP PDF generation using an HTML template.
- Frontend built with Vue 3 embedded into the HTML served by Apps Script.

---

## Repo structure
- `code.gs` — Google Apps Script server code (business logic, data access, email, Drive).
- `index.html` — Single-file frontend (Vue 3) with all UI components and client-side calls to `google.script.run`.
- `README.md` — (this file will replace the repository README with consolidated docs).

Notes:
- The Apps Script `doGet` in `code.gs` calls `HtmlService.createTemplateFromFile("Index")`. Apps Script file names are case-sensitive. If you deploy the repository files directly to Apps Script, ensure the HTML file is named `Index.html` (capital "I") or change the `doGet` call to `createTemplateFromFile("index")`. See "Deployment" below.

---

## Quick Start (developer)
1. Clone the repository:
   git clone https://github.com/mrjeronluther/Payment-Endorsement-Portal.git

2. Open the Google Apps Script editor or use clasp to push files.

3. Configure spreadsheet/folder IDs (see next section).

4. Deploy as a Web App (see Deploying to Google Apps Script).

---

## Prerequisites
- Google account with appropriate access to target Google Sheets/Drive files.
- The Google Sheets referenced by the script (IDs inside `code.gs`) must exist and be accessible by the account running the script.
- If you want to edit/deploy from local machine: Node + clasp (optional).
- A G Suite / Google Workspace domain is recommended if you rely on Session.getActiveUser() for domain-restricted auth.

---

## Configuration — where to edit
Open `code.gs` and locate the top-level constants and variables. Update these to match your environment or store them in ScriptProperties/PropertiesService if preferred.

Important IDs found in the code:
- USER_DB_ID (users spreadsheet): `1dBO8ThI7FEKb24D9sPVWokfXLuWUx5aCQvisrT9wBvI`
- USER_TAB (users sheet/tab): `PEP`
- REGISTRY_SS_ID (RFP registry spreadsheet): `1YAvZmCdWXbjOcJA-uUY40e6qVqzyiHcB06NpiPcz6y4`
- TAB_NAME for registry: `autogenrfp`
- MASTER_ID (index of source files): `10p-nv_qAN0GzZHAVInyb9_bprk2_sLf08190qdMf8Mc`
- FOLDER_ID (Drive folder for attachments): `1eFcLGXPEnUSi14aPvvA9m2ATV8Racsuf`
- LOGO_FILE_ID (logo used when generating PDF): `1QZ3XkRk1x-p_GFSjKhQO8i4dmIVzE1dG`

Action items:
- Replace these IDs with your own spreadsheets and folder IDs.
- Consider storing IDs in `PropertiesService.getScriptProperties()` for easier config and to avoid editing code.

---

## Deploying to Google Apps Script (Web App)
1. Open Google Apps Script (https://script.google.com) and create a new project.
2. Add `code.gs` and create an HTML file with the same content as `index.html`. Important: the file name used in `doGet(e)` must match the HTML filename (without `.html`):
   - If `doGet` uses `createTemplateFromFile("Index")`, name the file `Index.html`.
   - If you prefer `index.html` (lowercase), change `createTemplateFromFile("Index")` to `createTemplateFromFile("index")` in `code.gs`.
3. In Project Settings, enable required scopes when prompted (the script will request authorization).
4. Test the web app by running `doGet`/preview in the editor or using the deployed URL.
5. Deploy > New deployment > Select "Web app":
   - Who executes the app: choose "Me" or "User accessing the web app" depending on your flow.
   - Who has access: choose "Only myself", "Anyone", or "Anyone within <domain>" — for `Session.getActiveUser()` to return actual signed-in user email reliably, select "Anyone within <your G Suite domain>" and "Execute as: User accessing the web app".
6. Authorize scopes (Drive, MailApp, SpreadsheetApp, etc.) when prompted.

Important: `Session.getActiveUser().getEmail()` behavior depends on domain and web-app settings. If you need identity reliably, deploy as "Execute as: User accessing the web app" and restrict access to your Google Workspace domain.

---

## Frontend ↔ Backend integration
- The frontend uses `google.script.run` to invoke server functions. Examples in the UI include:
  - `.getActiveUserEmail()`
  - `.sendVerificationCode(email)`
  - `.verifyRegistrationCode(email, otp)`
  - `.finalizeRegistration(formData)`
  - `.authenticateUser(credentials)`
  - `.generateRfpNumber()`
  - `.getSourceFileNames()`
  - `.getRfpDataAdvanced(sourceFile, rfpInput, invoiceInput)`
  - `.requestPasswordResetCode(email)`
  - `.verifyResetCode(email, otp)`
  - `.submitPasswordReset(email, newPassword)`
  - `.processSubmission(payload)`

The client expects JSON-like responses and uses `.withSuccessHandler()` and `.withFailureHandler()` for handling callbacks.

---

## Exposed backend functions (public API via google.script.run)
Listed here so you know what the client calls and what each does:

- doGet(e) — returns the main HTML (web app entrypoint).
- include(filename) — helper to include other HTML parts if used.
- getActiveUserEmail() — returns current user's email (Session).
- sendVerificationCode(email) — sends registration OTP to email and caches OTP.
- verifyRegistrationCode(email, userCode) — validates OTP from cache.
- finalizeRegistration(formData) — writes new user row into users spreadsheet (gap-filling).
- authenticateUser(credentials) — authenticates a user against USER_DB_ID spreadsheet (checks employee & permission status).
- generateRfpNumber() — thread-safe RFP number generation persisted to REGISTRY_SS_ID.
- getSourceFileNames() — returns list of source files + tab names for selection UI.
- getRfpDataAdvanced(sourceFile, rfpInput, invoiceInput) — high-level search which returns status and data.
- performStrictSearch(selectedSource, rfpInput, invoiceInput) — underlying strict search logic over source spreadsheets.
- requestPasswordResetCode(email) — sends reset OTP and caches it.
- verifyResetCode(email, userCode) — validates reset OTP.
- submitPasswordReset(email, newPassword) — overwrites stored password in users sheet.
- processSubmission(payload) — central submission handler that:
  - validates required fields (invoice required),
  - checks duplicates,
  - uploads attachment to Drive (using FOLDER_ID),
  - optionally creates RFP PDF via createRfpPdf(),
  - appends a submission row to the submissions sheet.
- createRfpPdf(p, folder) — builds a printable HTML and converts to PDF (returned Drive URL).
- getPreviousParticipantEmails() — returns unique set of participant emails used previously (suggestions in UI).

Return shapes: most server functions return objects like `{ success: true, message: '...' }`, arrays, or throw Errors which are caught by `.withFailureHandler`.

---

## UI walkthrough (major views & flows)
- Login view
  - Uses email + password input.
  - Calls `authenticateUser`.
  - Maintains login history in localStorage for suggestions.

- Register view
  - Auto-detects authenticated Google session email via `getActiveUserEmail()`.
  - Request OTP (`sendVerificationCode`) → verify (`verifyRegistrationCode`) → finalize (`finalizeRegistration`).
  - Access level set to "REQUESTOR", initial permission "PENDING".

- Forgot Password
  - Request reset OTP → verify → submit new password (`submitPasswordReset`).

- Dashboard
  - Displays logged in user & navigation cards to:
    - RFP / PEF submission (FormRequest view)
    - Approver / Requestor / Coordinator portals (placeholders)

- FormRequest (Request for Payment)
  - Select Source File (populated by `getSourceFileNames`).
  - Enter RFP No or Invoice No to search source data (`getRfpDataAdvanced`).
  - Input particulars (dynamic `tableFields`).
  - Upload PRF (PDF) — handled client-side by base64 then submitted to server in `processSubmission`.
  - Participants: Primary + Secondary emails with suggestions from `getPreviousParticipantEmails`.
  - Submit: calls `processSubmission` and expects success message.

