# Payment Endorsement Platform (PEP)

## Introduction
The **Payment Endorsement Platform (PEP)** is a secure, enterprise-grade financial workflow tool built within the Google Workspace ecosystem. It streamlines the creation, verification, and endorsement of **Requests for Payment (RFP)** and **Payment Endorsement Forms (PEF)**.

The system ensures financial integrity by cross-referencing requests against master source files, generating unique RFP tracking numbers, and producing tamper-proof PDF documents for audit trails.

### Core Features:
*   **Zero-Trust Registration:** Identity verification via Google Session and OTP (One-Time Password) email validation.
*   **Strict Search Engine:** Advanced lookup logic that validates RFP/Invoice data against specific source spreadsheets before allowing submission.
*   **Dynamic RFP Generation:** Automated sequence logic to generate unique, formatted tracking IDs (e.g., `MALL-YYYY-MM-000000X`).
*   **Automated PDF Engine:** Converts web forms into standardized, print-ready PDF layouts with brand headers and signature lines.
*   **Intelligent Participant Routing:** Suggests previous participants and routes notifications to Primary and Secondary recipients.

---

## Installation Instructions

### 1. Spreadsheet Database Setup
You require three specific Google Sheets to act as the backend:
*   **User Database (`USER_DB_ID`):** A sheet named `PEP` to store user profiles, passwords, and access levels.
*   **RFP Registry (`REGISTRY_SS_ID`):** A sheet named `autogenrfp` to manage the ID sequence and a `SUBMISSIONS` sheet for transaction logs.
*   **Master Index (`MASTER_ID`):** A sheet named `SOURCEFILES` containing `File Name`, `File URL`, and `Tab Name` for the search engine to query.

### 2. Google Drive Configuration
*   Create a dedicated folder in Google Drive to store submitted PDFs and attachments.
*   Copy the **Folder ID** from the URL.

### 3. Script Deployment
1.  Open [Google Apps Script](https://script.google.com).
2.  Create a **New Project**.
3.  Copy the `Index.html` and `Code.js` content into the editor.
4.  **Update Global Constants:** In `Code.js`, ensure the following variables match your IDs:
    *   `USER_DB_ID`
    *   `REGISTRY_SS_ID`
    *   `MASTER_ID`
    *   `FOLDER_ID` (for PDF storage)

### 4. Web App Deployment
1.  Click **Deploy** > **New Deployment**.
2.  Select **Web App**.
3.  Execute As: **Me** (The Admin).
4.  Who has access: **Anyone within [Your Domain]**.
5.  Authorize permissions for **Drive, Gmail, and Sheets**.

---

## Usage Examples

### Executing a Search & Form Population
The following snippet demonstrates how the frontend handles the "Strict Search" to pull data from master spreadsheets into the RFP form.

```javascript
/**
 * Copy-Paste Ready: RFP Search Pattern
 * Triggers the backend lookup of an RFP or Invoice number
 * within a specific source spreadsheet.
 */

function performDataLookup() {
    const sourceFile = "Malls_Master_2024.xlsx";
    const rfpNumber = "MALL-2024-05-0000001";
    
    console.log("Searching for validated record...");

    google.script.run
        .withSuccessHandler((response) => {
            if (response.status === "FOUND") {
                // Populate the Vue.js particulars object
                const record = response.data[0]; 
                console.log("Record identified:", record["PAYEE NAME"]);
                alert("Data Loaded: " + record["SOA AMOUNT"]);
            } else if (response.status === "ALREADY_PAID") {
                alert("Warning: This invoice has already been processed.");
            } else {
                alert("No match found in source file.");
            }
        })
        .withFailureHandler((err) => {
            console.error("Search Engine Error:", err.message);
        })
        .getRfpDataAdvanced(sourceFile, rfpNumber, "");
}
```

---

## Tech Stack
*   **Backend:** Google Apps Script (V8)
*   **Frontend:** Vue.js 3
*   **UI Framework:** Bootstrap 5.3
*   **PDF Engine:** Google HTML-to-PDF Conversion
*   **Security:** Script Cache Service (OTP) & Google Session Authentication

---

> **Warning**  
> **- For MCD Internal Use Only**  
> This application handles sensitive financial endorsement data. Distribution of the source code or database IDs outside of authorized personnel is strictly prohibited.
