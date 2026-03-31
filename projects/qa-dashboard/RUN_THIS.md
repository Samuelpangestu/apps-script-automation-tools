# ⚡ Quick Setup: Web App URL

## Langkah Singkat (1 menit)

### Option 1: Run dari Apps Script Editor (Recommended)

1. **Buka Production Apps Script:**
   ```
   https://script.google.com/d/1lHO8yKyqKs1_n5GV1m-SJMACLS95Jc7yy6dM_ItyT-l_-GdmkGQk3OIO/edit
   ```

2. **Pilih function:** `setWebAppUrlDirect`

3. **Klik Run** ▶️

4. **Done!** ✅

Function akan otomatis set Web App URL:
```
https://script.google.com/a/macros/inadigital.co.id/s/AKfycbxswnlu4nYCdkKtsDrt8gWbEl75r7yAatMYTt5AKOXvTu1kBYMHdf-iKsdAOQi5aiMg3A/exec
```

### Option 2: Run dari Dashboard Spreadsheet

1. **Buka Production Dashboard:**
   ```
   https://docs.google.com/spreadsheets/d/1b2RBemEgo5B0YfUJHqAw8D0dH9Pg2Avgcngb7iz1PxY/edit
   ```

2. **Menu:** Extensions > Apps Script

3. **Pilih function:** `setWebAppUrlDirect`

4. **Klik Run** ▶️

5. **Done!** ✅

## Verify Setup

**Check di Execution Log:**
```
✅ Web App URL set successfully: https://script.google.com/a/macros/...
Verified: https://script.google.com/a/macros/...
```

**Test Notification:**
1. Dari Dashboard spreadsheet
2. Menu: **Notifications** > **Test Notification Now**
3. Check WhatsApp message
4. Link "📊 Web Dashboard" harus muncul dengan URL yang benar

## What Happens

Function `setWebAppUrlDirect()` akan:

1. ✅ Save Web App URL ke Script Properties
2. ✅ Verify URL tersimpan dengan benar
3. ✅ Show confirmation message

Web App URL akan digunakan di:
- **WhatsApp notifications** (line 1369 Notifications.js)
- Footer setiap daily bug report
- Link "📊 Web Dashboard"

## URL Yang Di-Set

```
Production Web App:
https://script.google.com/a/macros/inadigital.co.id/s/AKfycbxswnlu4nYCdkKtsDrt8gWbEl75r7yAatMYTt5AKOXvTu1kBYMHdf-iKsdAOQi5aiMg3A/exec
```

## Setelah Setup Selesai

Function `setWebAppUrlDirect()` bisa di-delete kalau mau (optional):

1. Buka `WebAppBackend.js`
2. Delete function `setWebAppUrlDirect()` (lines 97-130)
3. Commit & deploy ulang

Atau biarkan aja, tidak masalah. Function ini hanya untuk setup sekali.

---

**Quick Link to Apps Script Editor:**
https://script.google.com/d/1lHO8yKyqKs1_n5GV1m-SJMACLS95Jc7yy6dM_ItyT-l_-GdmkGQk3OIO/edit

**Function to Run:** `setWebAppUrlDirect`

**Status:** ⏳ Pending - Butuh 1x run manual dari Apps Script Editor

**Estimated Time:** < 1 minute
