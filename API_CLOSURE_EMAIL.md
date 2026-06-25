# Test Closure Email API Documentation

## Overview

API endpoints untuk Next.js Dashboard untuk mengirim test closure email dengan PDF attachment dari External Test Report data.

**Base URL:** `https://script.google.com/macros/s/{DEPLOYMENT_ID}/exec`

---

## API Endpoints

### 1. GET External Test Reports

**Endpoint:** `?action=getExternalTestReports`

**Method:** GET

**Description:** Mengambil semua External Test Report data dari QATM modules dengan optional filtering.

**Query Parameters:**
- `project` (optional): Filter by project name
- `module` (optional): Filter by module name
- `status` (optional): Filter by overall status (e.g., "Ready for Closure", "Approved")

**Example Request:**
```bash
# Get all reports
curl "https://script.google.com/.../exec?action=getExternalTestReports"

# Get reports for specific project
curl "https://script.google.com/.../exec?action=getExternalTestReports&project=SIPGN"

# Get reports ready for closure
curl "https://script.google.com/.../exec?action=getExternalTestReports&status=Ready%20for%20Closure"
```

**Response:**
```json
{
  "success": true,
  "data": [
    {
      "moduleId": "1abc...xyz",
      "project": "SIPGN",
      "module": "Manajemen Gizi",
      "submodule": "Nutritionist App",
      "picQA": "John Doe",
      "isExternalQA": true,
      "externalTeam": "Vendor XYZ",
      "statusReview": "In Review",
      "functionalEvidenceUrl": "https://...",
      "functionalReviewStatus": "Approved",
      "performanceEvidenceUrl": "https://...",
      "performanceReviewStatus": "Approved",
      "vaptEvidenceUrl": "https://...",
      "vaptReviewStatus": "In Review",
      "overallStatus": "Ready for Closure",
      "reviewer": "Jane Smith",
      "reviewDate": "2026-06-09",
      "notes": "All tests passed successfully"
    }
  ],
  "total": 1,
  "timestamp": "2026-06-09T10:30:00.000Z"
}
```

---

### 2. POST Send Closure Email

**Endpoint:** `?action=sendClosureEmail` (via doPost)

**Method:** POST

**Content-Type:** application/json

**Description:** Mengirim email closure langsung dengan PDF attachment.

**Request Body:**
```json
{
  "action": "sendClosureEmail",
  "moduleId": "1abc...xyz",
  "emailTo": "stakeholder@example.com",
  "emailCc": "manager@example.com",
  "emailSubject": "Test Closure Report - SIPGN Nutritionist App",
  "emailBody": "<html>...</html>",  // Optional, auto-generated if not provided
  "attachPDF": true  // Optional, default true
}
```

**Example Request:**
```bash
curl -X POST "https://script.google.com/.../exec" \
  -H "Content-Type: application/json" \
  -d '{
    "action": "sendClosureEmail",
    "moduleId": "1abc...xyz",
    "emailTo": "stakeholder@example.com",
    "emailCc": "manager@example.com",
    "emailSubject": "Test Closure Report - SIPGN Nutritionist App",
    "attachPDF": true
  }'
```

**Response:**
```json
{
  "success": true,
  "message": "Email sent successfully",
  "sentTo": "stakeholder@example.com",
  "sentCc": "manager@example.com",
  "timestamp": "2026-06-09T10:35:00.000Z"
}
```

---

### 3. POST Create Email Draft

**Endpoint:** `?action=createClosureEmailDraft` (via doPost)

**Method:** POST

**Content-Type:** application/json

**Description:** Membuat Gmail draft (tidak langsung send), sehingga user bisa edit di Gmail sebelum send.

**Request Body:**
```json
{
  "action": "createClosureEmailDraft",
  "moduleId": "1abc...xyz",
  "emailTo": "stakeholder@example.com",
  "emailCc": "manager@example.com",
  "emailSubject": "Test Closure Report - SIPGN Nutritionist App",
  "emailBody": "<html>...</html>",  // Optional
  "attachPDF": true  // Optional, default true
}
```

**Example Request:**
```bash
curl -X POST "https://script.google.com/.../exec" \
  -H "Content-Type: application/json" \
  -d '{
    "action": "createClosureEmailDraft",
    "moduleId": "1abc...xyz",
    "emailTo": "stakeholder@example.com",
    "emailSubject": "Test Closure Report - SIPGN",
    "attachPDF": true
  }'
```

**Response:**
```json
{
  "success": true,
  "message": "Email draft created successfully",
  "draftId": "r-1234567890",
  "recipient": "stakeholder@example.com",
  "subject": "Test Closure Report - SIPGN",
  "timestamp": "2026-06-09T10:40:00.000Z"
}
```

**Note:** Draft dapat di-edit di Gmail web/mobile sebelum dikirim.

---

## Email Template

Email yang di-generate otomatis memiliki struktur:

### HTML Email Structure
- **Header:** Blue banner dengan project/module name
- **External Test Information:** Team, status, reviewer, date
- **Testing Evidence Table:** Functional, Performance, VAPT dengan evidence links
- **Notes Section:** Additional notes dari External Test Report
- **Footer:** Auto-generated timestamp

### PDF Attachment
- **Filename:** `Test_Closure_Report_YYYYMMDD.pdf`
- **Content:** Summary sheet dari QATM (dapat dikustomisasi)
- **Format:** PDF export dari Google Sheets

---

## Email Sender Configuration

Email dikirim menggunakan akun **departemen.qa@inadigital.co.id**.

### Setup Email Sender (One-time):

**Option 1: Via Script Properties (Recommended)**
```javascript
// Run dari Apps Script Editor
function setupEmailSender() {
  PropertiesService.getScriptProperties()
    .setProperty('EMAIL_SENDER_ADDRESS', 'departemen.qa@inadigital.co.id');
}
```

**Option 2: Default**
Jika tidak di-set, otomatis menggunakan `departemen.qa@inadigital.co.id`.

**Catatan:** Script harus di-authorize dengan akun departemen.qa untuk bisa send email.

---

## Next.js Integration Example

### Fetch External Test Reports
```typescript
// app/api/external-reports/route.ts
export async function GET(request: Request) {
  const { searchParams } = new URL(request.url);
  const project = searchParams.get('project');

  const url = new URL(process.env.APPS_SCRIPT_URL!);
  url.searchParams.set('action', 'getExternalTestReports');
  if (project) url.searchParams.set('project', project);

  const response = await fetch(url.toString());
  const data = await response.json();

  return Response.json(data);
}
```

### Send Closure Email
```typescript
// app/api/send-closure-email/route.ts
export async function POST(request: Request) {
  const body = await request.json();

  const response = await fetch(process.env.APPS_SCRIPT_URL!, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      action: 'sendClosureEmail',
      moduleId: body.moduleId,
      emailTo: body.emailTo,
      emailCc: body.emailCc,
      emailSubject: body.emailSubject,
      emailBody: body.emailBody,  // Optional custom body
      attachPDF: body.attachPDF ?? true
    })
  });

  const data = await response.json();
  return Response.json(data);
}
```

### Create Draft
```typescript
// app/api/create-draft/route.ts
export async function POST(request: Request) {
  const body = await request.json();

  const response = await fetch(process.env.APPS_SCRIPT_URL!, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      action: 'createClosureEmailDraft',
      moduleId: body.moduleId,
      emailTo: body.emailTo,
      emailCc: body.emailCc,
      emailSubject: body.emailSubject,
      attachPDF: true
    })
  });

  const data = await response.json();
  return Response.json(data);
}
```

---

## React Component Example

```tsx
'use client';

import { useState } from 'react';

export default function ClosureEmailForm({ moduleId, project, module }) {
  const [emailTo, setEmailTo] = useState('');
  const [emailCc, setEmailCc] = useState('');
  const [loading, setLoading] = useState(false);

  const handleSendEmail = async () => {
    setLoading(true);
    try {
      const response = await fetch('/api/send-closure-email', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          moduleId,
          emailTo,
          emailCc,
          emailSubject: `Test Closure Report - ${project} ${module}`,
          attachPDF: true
        })
      });

      const data = await response.json();
      if (data.success) {
        alert('Email sent successfully!');
      } else {
        alert('Failed: ' + data.error);
      }
    } catch (error) {
      alert('Error: ' + error.message);
    } finally {
      setLoading(false);
    }
  };

  const handleCreateDraft = async () => {
    setLoading(true);
    try {
      const response = await fetch('/api/create-draft', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          moduleId,
          emailTo,
          emailCc,
          emailSubject: `Test Closure Report - ${project} ${module}`
        })
      });

      const data = await response.json();
      if (data.success) {
        alert(`Draft created! Check Gmail drafts (ID: ${data.draftId})`);
        window.open('https://mail.google.com/mail/#drafts', '_blank');
      }
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="space-y-4">
      <input
        type="email"
        placeholder="To: stakeholder@example.com"
        value={emailTo}
        onChange={(e) => setEmailTo(e.target.value)}
        className="w-full p-2 border rounded"
      />
      <input
        type="email"
        placeholder="Cc: manager@example.com (optional)"
        value={emailCc}
        onChange={(e) => setEmailCc(e.target.value)}
        className="w-full p-2 border rounded"
      />
      <div className="flex gap-2">
        <button
          onClick={handleCreateDraft}
          disabled={loading || !emailTo}
          className="px-4 py-2 bg-blue-500 text-white rounded disabled:opacity-50"
        >
          Create Draft
        </button>
        <button
          onClick={handleSendEmail}
          disabled={loading || !emailTo}
          className="px-4 py-2 bg-green-500 text-white rounded disabled:opacity-50"
        >
          Send Now
        </button>
      </div>
    </div>
  );
}
```

---

## Testing

### 1. Test GET Endpoint
```bash
# Replace {DEPLOYMENT_ID} with actual deployment ID
curl "https://script.google.com/macros/s/{DEPLOYMENT_ID}/exec?action=getExternalTestReports"
```

### 2. Test POST Endpoint (Create Draft)
```bash
curl -X POST "https://script.google.com/macros/s/{DEPLOYMENT_ID}/exec" \
  -H "Content-Type: application/json" \
  -d '{
    "action": "createClosureEmailDraft",
    "moduleId": "YOUR_QATM_SPREADSHEET_ID",
    "emailTo": "your-email@example.com",
    "emailSubject": "Test Closure Report - Test",
    "attachPDF": true
  }'
```

### 3. Verify Email Draft
1. Buka Gmail dengan akun departemen.qa@inadigital.co.id
2. Go to Drafts
3. Cari email draft dengan subject yang baru di-create
4. Verify PDF attachment ada
5. Edit jika perlu, lalu send manual

---

## Error Handling

### Common Errors

**Error: "moduleId is required"**
```json
{
  "success": false,
  "error": "moduleId is required",
  "timestamp": "2026-06-09T10:00:00.000Z"
}
```
**Solution:** Pastikan `moduleId` (QATM Spreadsheet ID) di-include di request body.

**Error: "External Test Report tab not found"**
```json
{
  "success": false,
  "error": "External Test Report tab not found in QATM"
}
```
**Solution:** Jalankan `/deploy-template` atau broadcast `External Test Report` tab ke QATM tersebut.

**Error: "Failed to send email"**
```json
{
  "success": false,
  "error": "Failed to send email: User rate limit exceeded"
}
```
**Solution:** Gmail memiliki quota limit. Wait beberapa menit atau gunakan `createClosureEmailDraft` sebagai alternatif.

---

## Deployment

### 1. Deploy Apps Script as Web App
```bash
cd projects/qa-dashboard
clasp push
clasp deploy --description "Add closure email API"
```

### 2. Get Deployment URL
```bash
clasp deployments
```

Copy deployment URL, contoh:
```
https://script.google.com/macros/s/AKfycbxxx.../exec
```

### 3. Add to Next.js Environment
```bash
# .env.local
APPS_SCRIPT_URL=https://script.google.com/macros/s/AKfycbxxx.../exec
```

### 4. Test API
```bash
curl "https://script.google.com/macros/s/AKfycbxxx.../exec?action=getExternalTestReports"
```

---

## Permissions

Apps Script membutuhkan permissions untuk:
- ✅ Google Sheets (read QATM data)
- ✅ Gmail (send email / create drafts)
- ✅ Drive (generate PDF from sheets)

**First run:** User akan diminta authorize script untuk access Gmail dan Drive.

---

## Changelog

**v1.0.0 - 2026-06-09**
- Initial implementation
- GET `getExternalTestReports` with filtering
- POST `sendClosureEmail` with PDF attachment
- POST `createClosureEmailDraft` untuk editable drafts
- Auto-generated HTML email template
- departemen.qa email sender support

---

## Support

**Documentation:** `ARCHITECTURE.md`, `CLASP_WORKFLOW.md`
**Contact:** departemen.qa@inadigital.co.id
