# How to Get Apps Script Web App URL

## Manual Method (Via Apps Script Editor)

1. **Open Apps Script Editor:**
   ```
   https://script.google.com/d/1lHO8yKyqKs1_n5GV1m-SJMACLS95Jc7yy6dM_ItyT-l_-GdmkGQk3OIO/edit
   ```

2. **Deploy Web App:**
   - Click **Deploy** button (top right)
   - Select **New deployment**

3. **Configure Deployment:**
   - Type: Select **Web app**
   - Description: "Closure Email API - v1.0"
   - Execute as: **Me (your-email@inadigital.co.id)**
   - Who has access: **Anyone** (for Next.js to access)
   - Click **Deploy**

4. **Copy Web App URL:**
   ```
   https://script.google.com/macros/s/AKfycbxxx.../exec
   ```

   This is your `NEXT_PUBLIC_APPS_SCRIPT_URL`

---

## Via Clasp (Command Line)

```bash
cd /Users/qainadigital/WebstormProjects/qa-test-management-template/projects/qa-dashboard

# List deployments
clasp deployments

# Create new deployment
clasp deploy --description "Closure Email API v1.0"

# Output will show Web App URL
```

---

## Important Notes

⚠️ **First-time Access:**
- When Next.js first calls the Apps Script URL, user akan diminta authorize
- Login dengan akun **departemen.qa@inadigital.co.id**
- Accept permissions untuk:
  - Google Sheets (read QATM data)
  - Gmail (send email, create drafts)
  - Google Drive (generate PDF)

⚠️ **URL Format:**
- Correct: `https://script.google.com/macros/s/AKfycbxxx.../exec`
- Wrong: `https://script.google.com/d/1lHO8yKy.../edit` (this is editor URL)

⚠️ **Updates:**
- Setiap `clasp push`, Web App otomatis update (no need redeploy)
- Deployment URL tidak berubah

---

## Save to Next.js Environment

```bash
# .env.local
NEXT_PUBLIC_APPS_SCRIPT_URL=https://script.google.com/macros/s/AKfycbxxx.../exec
NEXT_PUBLIC_APP_URL=http://localhost:3000
```

Restart Next.js dev server setelah update `.env.local`.
