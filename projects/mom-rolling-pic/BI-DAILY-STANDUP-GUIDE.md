# 📋 BI-DAILY STANDUP - QUICK START GUIDE

## 🎯 Overview

Sistem bi-daily standup yang sudah di-revamp dengan fitur:
- ✅ **Auto-generate** 3 rows per person (Done, In Progress, Blocker)
- ✅ **WhatsApp Reminder** 30 menit sebelum standup
- ✅ **WhatsApp Summary** dengan detail task breakdown
- ✅ **Multi-project support** (Project A & Project B terpisah)
- ✅ **Configurable schedule** per project dan per hari

---

## 📁 File Structure

```
projects/mom-rolling-pic/src/
├── Config.js           - Config management untuk kedua project
├── StandupLog.js       - Auto-generate rows & data management
├── Notifications.js    - WhatsApp reminder & summary
├── MenuFunctions.js    - Custom menu & user functions
├── Triggers.js         - Auto triggers untuk reminder & summary
└── appsscript.json     - Apps Script configuration
```

---

## 🚀 Deployment Steps

### 1. Deploy ke Apps Script

Karena ada authentication issue, deploy manual via clasp:

```bash
cd projects/mom-rolling-pic
clasp login
clasp push
```

**Atau upload manual via Apps Script Editor:**
1. Buka [script.google.com](https://script.google.com)
2. Buka project dengan Script ID: `1Rddhr1q5W-E1p-pnzMAlFzoAdjegaWFQ0BKiurKmNBPa8BDUiQDH5drg`
3. Delete file `Code.js` lama
4. Upload files baru dari `src/` folder

---

## ⚙️ Setup Guide (Setelah Deploy)

### Step 1: Initialize Config

1. Buka spreadsheet standup
2. Menu: **📋 Bi-Daily Standup → ⚙️ Setup → 1️⃣ Initialize Config**
3. Update settings di tab **Config**:

**Project A:**
- Team Members: `Alice, Bob, Carol` (sesuaikan nama tim Anda)
- Monday Time: `09:00`
- Wednesday Time: `09:00`
- Friday Time: `09:00`
- Reminder Offset: `30` (menit sebelum standup)
- Summary Time: `17:00`
- WhatsApp Group ID: `120363xxx@g.us` (gunakan menu Get Groups)
- Fonnte Token: `[your token]`
- Enable Reminder: `TRUE`
- Enable Summary: `TRUE`

**Project B:**
- (Sama seperti Project A, tapi bisa beda team members & waktu)

### Step 2: Get WhatsApp Group ID

1. Menu: **📋 Bi-Daily Standup → 📱 Test Notifications → Get WhatsApp Groups**
2. Masukkan Fonnte Token
3. Copy Group ID (format: `120363xxx@g.us`)
4. Paste ke Config sheet (Project A: Row 12 Col B, Project B: Row 26 Col B)

### Step 3: Initialize Standup Sheets

1. Menu: **📋 Bi-Daily Standup → ⚙️ Setup → 2️⃣ Initialize Standup Sheets**
2. Akan create 2 tab baru:
   - **Project A Standup**
   - **Project B Standup**

### Step 4: Setup Auto Triggers

1. Menu: **📋 Bi-Daily Standup → ⚙️ Setup → 🔄 Setup Auto Triggers**
2. Authorize Apps Script (jika diminta)
3. Triggers akan dibuat untuk:
   - Monday reminder (Project A & B)
   - Wednesday reminder (Project A & B)
   - Friday reminder (Project A & B)
   - Daily summary (Project A & B)

---

## 🧪 Testing

### Test Reminder

Menu: **📋 Bi-Daily Standup → 📱 Test Notifications → Test Reminder (Project A/B)**

Akan kirim WhatsApp seperti:
```
🔔 BI-DAILY STANDUP REMINDER
📅 Senin, 07 April 2026
━━━━━━━━━━━━━━━━━━━━

⏰ Standup Time: 09:00
📋 Project: Project A

👥 Team: Alice, Bob, Carol

━━━━━━━━━━━━━━━━━━━━

📝 PLEASE UPDATE:
✅ Done - Tasks completed since last standup
📋 In Progress - Tasks until next standup
🚨 Blockers - Any blockers or help needed

━━━━━━━━━━━━━━━━━━━━
📊 Update sheet: [link]
```

### Test Summary

Menu: **📋 Bi-Daily Standup → 📱 Test Notifications → Test Summary (Project A/B)**

Akan kirim WhatsApp dengan detail task breakdown (Opsi 2):
```
📊 BI-DAILY STANDUP SUMMARY
📅 Senin, 07 April 2026
📋 Project: Project A
━━━━━━━━━━━━━━━━━━━━

👤 Alice:
  ✅ Done:
    • Selesai API login
    • Testing module A
    • Deploy staging
  📋 In Progress:
    • Bug fix dashboard
    • Code review
  🚨 Blockers: None

👤 Bob:
  ✅ Done:
    • Deploy UAT
    • Testing done
  📋 In Progress:
    • Production release
  🚨 Blockers:
    • Waiting API keys from vendor

━━━━━━━━━━━━━━━━━━━━
📊 TOTAL:
  ✅ 5 tasks completed
  📋 3 tasks in progress
  🚨 1 blocker(s)

━━━━━━━━━━━━━━━━━━━━
📋 Full details: [link]
```

---

## 📅 Daily Workflow

### Sebelum Standup (Auto):
1. **08:30** - Auto-generate 3 rows untuk setiap person
2. **08:30** - Kirim WhatsApp reminder

### Saat Standup (Manual):
1. Team buka sheet dan isi:
   - **Done**: Tasks yang sudah selesai sejak standup terakhir
   - **In Progress**: Tasks yang akan dikerjakan sampai next standup
   - **Blocker**: Any blockers atau help needed

### Setelah Standup (Auto):
1. **17:00** (atau waktu yang dikonfigurasi) - Kirim WhatsApp summary

---

## 🔧 Troubleshooting

### WhatsApp tidak terkirim

**Check:**
1. Fonnte Token valid?
2. Group ID format benar (`120363xxx@g.us`)?
3. Device Fonnte online?
4. Enable Reminder/Summary = TRUE?

**Fix:**
- Jalankan "Get WhatsApp Groups" untuk refresh group list
- Tunggu 5-15 menit jika group baru dibuat
- Test dengan "Test Reminder" menu

### Rows tidak auto-generate

**Check:**
1. Triggers sudah di-setup?
2. Standup time sudah benar?

**Fix:**
- Re-run "Setup Auto Triggers"
- Manual generate: **📅 Generate Standup → Generate Today**

### Duplicate rows

**Sistem sudah protect** - Jika rows untuk tanggal tertentu sudah ada, tidak akan generate lagi (kecuali manual dengan `skipIfExists=false`)

---

## 📝 Manual Operations

### Generate Standup for Specific Date

Menu: **📋 Bi-Daily Standup → 📅 Generate Standup → Generate for Date...**

Input: `2026-04-09` (format YYYY-MM-DD)

### Remove All Triggers

Menu: **📋 Bi-Daily Standup → ⚙️ Setup → 🗑️ Remove All Triggers**

---

## 🎨 Customization

### Ubah Format Pesan WhatsApp

Edit file `Notifications.js`:
- Line 40-60: Reminder message format
- Line 90-120: Summary message format

### Ubah Jumlah Default Rows

Edit file `StandupLog.js`:
- Line 75-100: Rows generation logic
- Default: 3 rows (Done, In Progress, Blocker)

### Tambah Project C

1. Edit `Config.js` - tambah section Project C
2. Edit `StandupLog.js` - tambah handling Project C
3. Edit `Triggers.js` - tambah trigger handlers untuk Project C
4. Deploy ulang

---

## 📚 Additional Notes

### Standup Days

Default: **Monday, Wednesday, Friday**

Untuk ubah:
- Edit Config sheet (Standup Days)
- Summary hanya kirim di hari standup (logic di Triggers.js)

### Time Format

Harus `HH:MM` (24-hour format):
- ✅ `09:00`, `14:00`, `23:30`
- ❌ `9:00`, `2:00 PM`

### Team Members

Comma-separated, spasi akan di-trim:
- ✅ `Alice, Bob, Carol`
- ✅ `Alice,Bob,Carol`
- ❌ `Alice; Bob; Carol`

---

## 🆘 Support

Issues? Check:
1. Execution log di Apps Script Editor
2. WhatsApp group apakah ada error message
3. Config sheet format

---

**Last Updated:** April 7, 2026
**Version:** 1.0 (Bi-Daily Standup Revamp)
