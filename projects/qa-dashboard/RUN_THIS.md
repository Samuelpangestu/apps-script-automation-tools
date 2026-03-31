# ⚡ Quick Setup: Web App URL

## ⚠️ PENTING: Deployment dari Testing Script

**Web App HARUS di-deploy dari Testing Script** (Production Script error saat deployment).

**Data tetap aman** - hardcoded ke Production Spreadsheet.

---

## Langkah Setup (5 menit total)

### Step 1: Deploy Web App (Testing Script Only)

1. **Buka Testing Apps Script:**
   ```
   https://script.google.com/d/1LJ83OATTAp7ChDWGkrSTg0b9KmMhOABISBrAJrB54JksjQ7mi5oNB7C3/edit
   ```

2. **Deploy Web App:**
   - Click: Deploy > New deployment
   - Type: Web app
   - Execute as: Me
   - Who has access: Anyone with Google account
   - Click: Deploy
   - **Copy the Web App URL**

### Step 2: Set URL di Testing Script

1. **Di Testing Apps Script Editor**

2. **Pilih function:** `setupWebAppUrl`

3. **Klik Run** ▶️

4. **Paste Web App URL** yang di-copy dari Step 1

5. **Done!** ✅ Testing Script Properties updated

### Step 3: Set URL di Production Script

1. **Buka Production Apps Script:**
   ```
   https://script.google.com/d/1lHO8yKyqKs1_n5GV1m-SJMACLS95Jc7yy6dM_ItyT-l_-GdmkGQk3OIO/edit
   ```

2. **Pilih function:** `setupWebAppUrl`

3. **Klik Run** ▶️

4. **Paste Web App URL** yang sama dari Step 1

5. **Done!** ✅ Production Script Properties updated

### Current Production Web App URL:
```
https://script.google.com/a/macros/inadigital.co.id/s/AKfycbxswnlu4nYCdkKtsDrt8gWbEl75r7yAatMYTt5AKOXvTu1kBYMHdf-iKsdAOQi5aiMg3A/exec
```

## Verify Setup

**Check di Execution Log:**
```
✅ Web App URL saved: https://script.google.com/a/macros/...
```

**Test Notification:**
1. Dari Dashboard spreadsheet
2. Menu: **Notifications** > **Test Notification Now**
3. Check WhatsApp message
4. Link "📊 Web Dashboard" harus muncul dengan URL yang benar

## What Happens

Function `setupWebAppUrl()` akan:

1. ✅ Prompt untuk paste Web App URL
2. ✅ Validate URL format
3. ✅ Save ke Script Properties
4. ✅ Show confirmation message

Web App URL akan digunakan di:
- **WhatsApp notifications** (Notifications.js)
- Footer setiap daily bug report
- Link "📊 Web Dashboard"

---

## Quick Links

**Testing Apps Script:**
https://script.google.com/d/1LJ83OATTAp7ChDWGkrSTg0b9KmMhOABISBrAJrB54JksjQ7mi5oNB7C3/edit

**Production Apps Script:**
https://script.google.com/d/1lHO8yKyqKs1_n5GV1m-SJMACLS95Jc7yy6dM_ItyT-l_-GdmkGQk3OIO/edit

**Function to Run:** `setupWebAppUrl`

**Estimated Time:** 5 minutes total
