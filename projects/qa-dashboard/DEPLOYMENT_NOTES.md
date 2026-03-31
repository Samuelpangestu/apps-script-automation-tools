# ⚠️ DEPLOYMENT NOTES - IMPORTANT

## Deployment Script Configuration

### ✅ USE TESTING SCRIPT FOR WEB APP DEPLOYMENTS

**IMPORTANT:** Untuk deployment Web App, **SELALU pakai Testing Script**, BUKAN Production Script.

**Alasan:**
- Production Script **ERROR** saat deployment Web App
- Testing Script berfungsi normal untuk deployment

### Script IDs

**Testing Script (untuk deployment):**
```
Script ID: 1LJ83OATTAp7ChDWGkrSTg0b9KmMhOABISBrAJrB54JksjQ7mi5oNB7C3
URL: https://script.google.com/d/1LJ83OATTAp7ChDWGkrSTg0b9KmMhOABISBrAJrB54JksjQ7mi5oNB7C3/edit
```

**Production Script (untuk code push only, TIDAK untuk deployment):**
```
Script ID: 1lHO8yKyqKs1_n5GV1m-SJMACLS95Jc7yy6dM_ItyT-l_-GdmkGQk3OIO
URL: https://script.google.com/d/1lHO8yKyqKs1_n5GV1m-SJMACLS95Jc7yy6dM_ItyT-l_-GdmkGQk3OIO/edit
```

---

## Deployment Workflow

### ✅ Correct Workflow:

**1. Push Code ke Kedua Script:**
```bash
# Deploy to Testing
cd projects/qa-dashboard
./deploy-testing.sh

# Deploy to Production (code push only)
./deploy-production.sh
```

**2. Deploy Web App (dari Testing Script ONLY):**
```
1. Buka Testing Script Editor
2. Deploy > New deployment
3. Type: Web app
4. Execute as: Me
5. Who has access: Anyone with Google account
6. Click Deploy
7. Copy Web App URL
```

**3. Use Testing Script Web App URL for Production Notifications:**

Testing Script Web App URL akan digunakan untuk:
- WhatsApp notifications footer
- Email notifications
- Production daily triggers

**Data Source:** Tetap Production Spreadsheet (hardcoded di `getDashboardSpreadsheet_()`)

---

## Current Deployment Info

**Active Web App Deployment:**
```
Source: Testing Script (1LJ83OATTAp7ChDWGkrSTg0b9KmMhOABISBrAJrB54JksjQ7mi5oNB7C3)
URL: https://script.google.com/a/macros/inadigital.co.id/s/AKfycbxswnlu4nYCdkKtsDrt8gWbEl75r7yAatMYTt5AKOXvTu1kBYMHdf-iKsdAOQi5aiMg3A/exec
Data: Production Spreadsheet (1b2RBemEgo5B0YfUJHqAw8D0dH9Pg2Avgcngb7iz1PxY)
```

**Other Known Deployments:**
```
v62: AKfycbzh1ujm97GtRKvccYb7zLg7cEKzzUmzCcozGRYyRqvJBrJVpSSJm8v-Ct3FvAao5rorbg
```

---

## Why Testing Script for Deployment?

**Architecture:**
```
Testing Script (Deployment)
    ↓ Web App
    ↓ getDashboardSpreadsheet_()
    ↓ HARDCODED: Production Spreadsheet ID
    ↓
Production Data ✅
```

**Key Points:**
1. ✅ Testing Script bisa deploy Web App
2. ❌ Production Script error saat deploy
3. ✅ Testing Script Web App tetap baca Production data
4. ✅ Code di kedua script SAMA (di-sync via deploy scripts)
5. ✅ Testing deployment URL dipakai untuk production notifications

**Kenapa Aman:**
- Data source hardcoded ke Production Spreadsheet (line 33 WebAppBackend.js)
- Code di Testing & Production sama (synchronized)
- Web App URL dari Testing dipakai di Production notifications
- Script Properties independent (tidak cross-contaminate)

---

## Troubleshooting

### Production Script Deploy Error

**Error:** Production Script tidak bisa deploy Web App

**Solution:**
1. ✅ Gunakan Testing Script untuk deployment
2. ✅ Push code ke Production Script untuk sync
3. ✅ Set WEB_APP_URL di Production Script Properties untuk notifications

### Wrong Data Displayed

**Check:** `getDashboardSpreadsheet_()` function hardcoded to Production

```javascript
function getDashboardSpreadsheet_() {
  // FORCE PRODUCTION SPREADSHEET ONLY
  return SpreadsheetApp.openById('1b2RBemEgo5B0YfUJHqAw8D0dH9Pg2Avgcngb7iz1PxY');
}
```

### Deployment Version Mismatch

**Check:**
```bash
clasp deployments
clasp versions
```

Jika ada deployment yang perlu di-update:
1. Create new deployment dari Testing Script
2. Update WEB_APP_URL di kedua Script Properties
3. Test notification

---

## Best Practices

### ✅ DO:
- Use Testing Script for Web App deployments
- Push code to BOTH Testing and Production
- Set WEB_APP_URL in BOTH Script Properties
- Test notifications after deployment
- Keep code synchronized between scripts

### ❌ DON'T:
- ❌ Try to deploy Web App from Production Script (will error)
- ❌ Forget to push code to Production Script
- ❌ Use different code versions in Testing vs Production
- ❌ Forget to update WEB_APP_URL after new deployment

---

## Deployment Checklist

**When deploying new version:**

- [ ] 1. Update code locally
- [ ] 2. Push to Testing Script: `./deploy-testing.sh`
- [ ] 3. Push to Production Script: `./deploy-production.sh`
- [ ] 4. Deploy Web App from Testing Script Editor
- [ ] 5. Copy new Web App URL
- [ ] 6. Run `setupWebAppUrl()` in Testing Script (paste Web App URL)
- [ ] 7. Run `setupWebAppUrl()` in Production Script (paste same URL)
- [ ] 8. Test notification: Menu > Notifications > Test Now
- [ ] 9. Verify Web App URL in WhatsApp message
- [ ] 10. Commit deployment URL to git/documentation

---

**Last Updated:** March 31, 2026
**Verified By:** Samuel Pangestu - QA INA Digital

**Key Takeaway:**
🔴 **ALWAYS deploy Web App from Testing Script, NOT Production Script**
✅ **Data remains Production regardless of deployment source**
