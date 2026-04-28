# Payment Endorsement Portal

## Introduction
Payment Endorsement Portal is a Google Apps Script web app that lets PCU Group encode and endorse specific RFP numbers to RBG. It provides a simple HTML/JavaScript front-end with Apps Script server-side handlers to record and confirm endorsements. Intended for internal PCU/MCD use by staff responsible for RFP endorsements.

## Installation Instructions

Prerequisites:
- A Google account with access to Google Apps Script.
- Optional: Node.js + clasp if you prefer local development.

Manual (Apps Script UI):
1. Open https://script.google.com/ and create a new project.
2. Copy the repository files (HTML + JavaScript/Code.gs) into the Apps Script project files.
3. In the Apps Script editor, set `doGet`/`doPost` handlers (if not already present).
4. Deploy → New deployment → Select "Web app".
   - Execute as: Me (or appropriate service account)
   - Who has access: Only myself / Anyone in <your-domain> as required
5. Copy the Web app URL and share with authorized users.

Using clasp (optional):
1. Install clasp: `npm install -g @google/clasp`
2. Authenticate: `clasp login`
3. Clone this repo locally and create a new Apps Script project:
   - `clasp create --title "Payment Endorsement Portal" --type webapp`
4. Push files: `clasp push`
5. Deploy from the Apps Script UI or with `clasp deploy`.

## Usage Examples

Client-side example — submit an endorsement via fetch:
```javascript
// Replace with your deployed web app URL
const WEB_APP_URL = 'https://script.google.com/macros/s/REPLACE_WITH_ID/exec';

async function endorseRfp(rfpNumber, endorsedBy) {
  const resp = await fetch(WEB_APP_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ action: 'endorse', rfp: rfpNumber, endorsedBy })
  });
  return resp.json();
}

// Example usage:
endorseRfp('RFP-2026-001', 'jane.doe@pcu.local').then(console.log).catch(console.error);
```

Server-side Apps Script example — handle POST requests:
```javascript
// Code.gs
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.action === 'endorse' && payload.rfp) {
      // TODO: Validate payload and store endorsement (e.g., in a Sheet or DB)
      // Example response:
      const result = { status: 'success', rfp: payload.rfp, endorsedBy: payload.endorsedBy || null };
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'invalid action or missing rfp' })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}
```

Security and operational notes:
- Validate and sanitize all incoming data before storing.
- Restrict web app access to your organization as needed.
- Log endorsements and maintain an audit trail (e.g., Google Sheets or a secure DB).

- For MCD Internal Use Only
