# Web App URL Setup - WhatsApp Notification

## Overview

Web App URL akan digunakan di WhatsApp notifications (line 1369 Notifications.js) sebagai link ke dashboard.

```javascript
message += '📊 Web Dashboard: ' + dashboardWebAppUrl + '\n';
```

## Langkah Setup

### 1. Deploy Web App (jika belum)

1. Buka Apps Script Editor
2. Klik **Deploy** > **New deployment**
3. Select type: **Web app**
4. Configure:
   - Execute as: **Me**
   - Who has access: **Anyone with Google account**
5. Klik **Deploy**
6. **Copy Web App URL**

URL format:
```
https://script.google.com/a/macros/inadigital.co.id/s/AKfycby.../exec
```

### 2. Simpan URL ke Script Properties

**Dari Dashboard Spreadsheet:**

1. Buka Dashboard spreadsheet
2. Menu: **QA Dashboard** > **⚙️ Settings** > **Setup Web App URL**
3. Paste Web App URL yang sudah di-copy
4. Klik **OK**

✅ Web App URL tersimpan dan akan digunakan di WhatsApp notifications!

**Atau via Apps Script Editor:**

1. Buka Apps Script Editor
2. Pilih function: `setupWebAppUrl`
3. Klik **Run**
4. Paste Web App URL
5. Klik **OK**

### 3. Verify Setup

**Check Script Properties:**

1. Buka Apps Script Editor
2. Menu: **Project Settings** (gear icon)
3. Scroll ke **Script Properties**
4. Cari property: `WEB_APP_URL`
5. Value harus berisi Web App URL yang valid

**Test Notification:**

1. Menu: **Notifications** > **Test Notification Now**
2. Check WhatsApp message
3. Link "📊 Web Dashboard" harus mengarah ke Web App URL

## Current Web App URL

**Testing Script:**
```
https://script.google.com/a/macros/inadigital.co.id/s/AKfycbyGQZbca-O1zFVFyh4xvRDdX-M8Hce-FRorhp5d-7SrxSo706ernucw9ge1wPIwEP_qRw/exec
```

**Production Script:**
- Belum di-deploy sebagai Web App
- Gunakan Testing Script URL untuk sementara

## Kapan Update URL?

Update Web App URL jika:

1. **Deploy ulang Web App** dengan deployment ID baru
2. **Ganti environment** (Testing → Production)
3. **Deployment expired** atau di-delete

## Troubleshooting

### WhatsApp link mengarah ke spreadsheet, bukan Web App

**Penyebab:** Script Property `WEB_APP_URL` belum di-set atau kosong

**Solusi:**
1. Jalankan `setupWebAppUrl()` function
2. Paste Web App URL yang benar
3. Test ulang notification

### "Invalid URL" error saat setup

**Penyebab:** URL format tidak valid

**Solusi:**
Pastikan URL format:
```
https://script.google.com/a/macros/inadigital.co.id/s/AKfycby.../exec
```

Harus ada:
- `script.google.com`
- `/exec` di akhir (bukan `/dev`)

### Web App URL tidak berfungsi (404)

**Penyebab:** Deployment expired atau di-delete

**Solusi:**
1. Buka Apps Script Editor
2. Klik **Deploy** > **Manage deployments**
3. Check apakah deployment masih aktif
4. Jika tidak, create new deployment
5. Update URL dengan `setupWebAppUrl()`

## File Locations

**Function:** `setupWebAppUrl()`
- File: `projects/qa-dashboard/src/WebAppBackend.js`
- Lines: 55-95

**Usage:** WhatsApp Notification
- File: `projects/qa-dashboard/src/Notifications.js`
- Line: 1281-1283 (get URL from Script Properties)
- Line: 1369 (include in message)

## Related Functions

- `setupWebAppSpreadsheet()` - Setup spreadsheet ID
- `sendWhatsAppNotification_()` - Send WhatsApp notification
- `getDashboardSpreadsheet_()` - Get dashboard spreadsheet

---

**Last Updated:** March 31, 2026
**Author:** Samuel Pangestu - QA INA Digital
