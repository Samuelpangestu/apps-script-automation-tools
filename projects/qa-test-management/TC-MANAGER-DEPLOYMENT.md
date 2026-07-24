# TC Manager - Deployment Guide

**Version:** v6
**Last Updated:** 2025-07-24

---

## 📋 Overview

TC Manager is a custom Apps Script menu for QATM (QA Test Management) that ensures TC_Master and TC_Execution sheets stay aligned when inserting/deleting rows.

**Problem Solved:**
- Manual insert/delete breaks alignment between TC_Master and TC_Execution
- Screenshots and test results become misaligned with TC_IDs
- TC Manager maintains perfect 1:1 row synchronization

---

## 🚀 Quick Deployment

### **Option 1: Using Deployment Script (Recommended)**

```bash
cd projects/qa-test-management
./deploy-tcmanager.sh <SCRIPT_ID>
```

**Example:**
```bash
./deploy-tcmanager.sh 1EMrV7bdWIWN7jRMDe9TyAMRjpnBwz5MsCWqWkezgtI8uJuETIlA38mDD
```

### **Option 2: Manual Deployment via clasp**

1. Get Script ID (see instructions below)
2. Create temporary deployment folder:
   ```bash
   mkdir -p /tmp/qatm-deploy/src
   cd /tmp/qatm-deploy
   ```

3. Create `.clasp.json`:
   ```json
   {
     "scriptId": "YOUR_SCRIPT_ID_HERE",
     "rootDir": "./src"
   }
   ```

4. Copy files:
   ```bash
   cp path/to/TCManager.js ./src/
   cp path/to/appsscript.json ./src/
   ```

5. Deploy:
   ```bash
   clasp push --force
   ```

### **Option 3: Manual Copy-Paste**

1. Open target QATM spreadsheet
2. Click: **Extensions** → **Apps Script**
3. Create new file: `TCManager.gs`
4. Copy entire content from `src/TCManager.js`
5. Paste into `TCManager.gs`
6. Save (Ctrl/Cmd + S)
7. Refresh spreadsheet

---

## 🔑 How to Get Script ID

### **Method 1: Via Apps Script Editor**

1. Open target QATM spreadsheet
2. Click: **Extensions** → **Apps Script**
3. Click: **⚙️ Project Settings** (gear icon in left sidebar)
4. Copy **Script ID** under "IDs" section

### **Method 2: From URL**

1. Open target QATM spreadsheet
2. Click: **Extensions** → **Apps Script**
3. Look at URL in address bar:
   ```
   https://script.google.com/home/projects/1ABC123def456GHI789jkl/edit
   ```
4. Script ID = `1ABC123def456GHI789jkl` (after `/projects/`)

---

## 📦 Deployed Instances

| Spreadsheet | Script ID | Date Deployed | Version |
|-------------|-----------|---------------|---------|
| QATM Testing | `1MU8JF5Ve39LC2VG4hc2O9lEcaaAzQUNYFHE4Xu-F3zH31BD11Ok1AIbm` | 2025-07-23 | v6 |
| Production 1 | `1EMrV7bdWIWN7jRMDe9TyAMRjpnBwz5MsCWqWkezgtI8uJuETIlA38mDD` | 2025-07-24 | v6 |
| Production 2 | `1Y-ww5rtFRcXJtvgSd4KogUZG6d84ejQs3vCs8P2WmDBHVFCEBr23G3y0` | 2025-07-24 | v6 |
| Main Template | `1qfesXuPqt6IgAoRLAK7cl3Jy8i68HL2wLdG-Y9dF-jX7p7QbO64e7fA_` | Pending | - |

**To add new deployment:**
1. Get Script ID from target spreadsheet
2. Run deployment script
3. Update this table with Script ID and date

---

## ✅ Verification After Deployment

### **1. Check Menu Appears**

After deployment, refresh spreadsheet and check menu bar:

✅ **Should see:** `🔧 TC Manager` menu

Menu items:
- 📋 Insert TC Here
- 📋 Bulk Insert
- 🗑️ Delete TC
- ⚠️ Mark as Deprecated
- ℹ️ Help

### **2. Test Warning Popup**

1. Close spreadsheet completely
2. Re-open spreadsheet
3. **Should see:** Warning popup on load

```
⚠️ Important Warning

🚫 DO NOT insert/delete rows manually!

Always use 🔧 TC Manager menu:
• Insert TC Here
• Bulk Insert
• Delete TC

Manual insert/delete breaks TC alignment and test results.
```

(Popup shows once every 6 hours)

### **3. Test Insert TC Here**

1. Go to TC_Master sheet
2. Select any row (row 3 or below)
3. Click: `🔧 TC Manager` → `Insert TC Here`
4. Confirm dialog
5. **Verify:**
   - ✅ New row inserted in TC_Master
   - ✅ New row inserted in TC_Execution (at correct position)
   - ✅ Row is empty (ARRAYFORMULA will auto-populate)
   - ✅ Rows below shifted down in BOTH sheets

### **4. Test Bulk Insert**

1. Select row in TC_Master
2. Click: `🔧 TC Manager` → `Bulk Insert`
3. Enter number (e.g., 5)
4. Confirm
5. **Verify:**
   - ✅ 5 rows inserted in TC_Master
   - ✅ 5 rows inserted in TC_Execution
   - ✅ All rows empty
   - ✅ Dialog shows exact row ranges

### **5. Test Delete TC (Multiple Rows)**

1. Select multiple rows in TC_Master (Shift+click or drag)
2. Click: `🔧 TC Manager` → `Delete TC`
3. **Verify dialog shows:**
   - TC_Master rows: 5 to 10
   - TC_Execution rows: 11 to 16
   - Total TCs: 6
   - TC_IDs: TC-A-1, TC-A-2, ...
4. Confirm
5. **Verify:**
   - ✅ All selected rows deleted from both sheets
   - ✅ Rows below shifted up correctly

### **6. Test Mark as Deprecated (Multiple Rows)**

1. Select multiple rows in TC_Master
2. Click: `🔧 TC Manager` → `Mark as Deprecated`
3. Confirm
4. **Verify:**
   - ✅ All scenarios have `[DEPRECATED]` prefix
   - ✅ All rows grayed out (#F5F5F5 background, #999999 text)
   - ✅ Test results intact

---

## 🧹 Cleanup Old Code (If Needed)

If old menus still appear after deployment (e.g., "QATM Review", "QATM Tools"):

1. Open Apps Script editor: **Extensions** → **Apps Script**
2. In **Files** sidebar, delete all files **EXCEPT**:
   - ✅ `TCManager.gs` (keep)
   - ✅ `appsscript.json` (keep)
3. Delete files by clicking **⋮** (three dots) → **Remove**
4. Save (Ctrl/Cmd + S)
5. Refresh spreadsheet
6. **Verify:** Only `🔧 TC Manager` menu appears

---

## 📊 Features

### **Insert TC Here**
- Insert 1 row at current cursor position
- Syncs TC_Master + TC_Execution
- Shows exact row numbers in dialog
- Leaves rows empty for ARRAYFORMULA

### **Bulk Insert**
- Insert multiple rows (1-100) at once
- Enter number of rows via prompt
- Faster than inserting one-by-one
- Shows row ranges in confirmation

### **Delete TC**
- Supports single & multiple rows
- Shows TC_IDs before deletion
- Deletes from both sheets simultaneously
- Bottom-to-top deletion (avoids index shifting)
- **Warning:** Test results will be lost!

### **Mark as Deprecated**
- Supports single & multiple rows
- Adds `[DEPRECATED]` prefix to scenario
- Grays out rows visually
- **Safer than delete** - keeps test results intact

### **Session Reminder**
- Warning popup on spreadsheet open
- Warns against manual insert/delete
- Shows once every 6 hours (cached)
- Clear, simple message

---

## 🛠️ Technical Details

### **Sheet Structure Requirements**

- **TC_Master:**
  - Data starts at row 3
  - Column C = TC_ID
  - Column K (11) = Scenario

- **TC_Execution:**
  - Data starts at row 9
  - Columns B-G = ARRAYFORMULA auto-populated
  - Row offset: `execRow = 9 + (masterRow - 3)`

### **Row Offset Calculation**

```javascript
const execStartRow = 9;
const masterRow = tcMaster.getActiveRange().getRow(); // e.g., 5
const execRow = execStartRow + (masterRow - 3); // 9 + (5 - 3) = 11
```

**Example:**
- TC_Master row 3 → TC_Execution row 9
- TC_Master row 5 → TC_Execution row 11
- TC_Master row 10 → TC_Execution row 16

### **ARRAYFORMULA Compatibility**

TC Manager **does NOT copy formulas** to new rows. Rows are left empty so ARRAYFORMULA in header row can auto-populate.

**Before (v5 - wrong):**
```javascript
sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMULA);
// ❌ Breaks ARRAYFORMULA - static formula overrides auto-populate
```

**After (v6 - correct):**
```javascript
// Insert rows (leave empty - ARRAYFORMULA will auto-populate)
tcMaster.insertRowBefore(activeRow);
tcExecution.insertRowBefore(execRow);
// ✅ Row stays empty, ARRAYFORMULA works correctly
```

---

## 📝 Version History

| Version | Date | Changes |
|---------|------|---------|
| v1 | 2025-07-23 | Initial implementation (Add, Insert, Delete, Deprecated) |
| v2 | 2025-07-23 | Added Bulk Insert feature |
| v3 | 2025-07-23 | Added multi-row selection for Mark as Deprecated |
| v4 | 2025-07-23 | Added session reminder popup |
| v5 | 2025-07-23 | Added multi-row selection for Delete TC |
| v6 | 2025-07-24 | **Removed formula copy** (ARRAYFORMULA compatibility fix) |

**Current Production Version:** v6

---

## 🐛 Troubleshooting

### **Menu doesn't appear after deployment**

**Solution:**
1. Close spreadsheet completely
2. Clear browser cache (Ctrl+Shift+Del)
3. Re-open spreadsheet
4. Hard refresh (Ctrl+Shift+R)

### **Warning popup doesn't show**

**Reason:** Popup cached for 6 hours

**Solution (for testing):**
1. Wait 6 hours, OR
2. Clear Apps Script cache:
   - Run in Apps Script editor:
   ```javascript
   function clearCache() {
     CacheService.getScriptCache().remove('tc_manager_reminder_shown');
   }
   ```
3. Re-open spreadsheet

### **ARRAYFORMULA not populating new rows**

**Check:**
1. Row is actually empty (no static formulas)
2. ARRAYFORMULA in header row is working
3. ARRAYFORMULA range covers new rows

**Fix:**
- Verify TC Manager v6 is deployed (not v5)
- v6 does NOT copy formulas

### **"Error: Sheets not found"**

**Reason:** TC_Master or TC_Execution sheet name is different

**Solution:**
- Ensure exact sheet names: `TC_Master` and `TC_Execution`
- Case-sensitive!

### **Authentication error when deploying**

**Error:** `invalid_grant` or `invalid_rapt`

**Solution:**
```bash
clasp logout
clasp login
```

Then retry deployment.

---

## 📞 Support

**Issues or Questions:**
1. Check this documentation first
2. Review Apps Script execution logs:
   - Apps Script editor → **View** → **Execution log**
3. Check browser console (F12)
4. Contact: Claude Code for debugging

---

## 📄 Files

| File | Description |
|------|-------------|
| `src/TCManager.js` | **Production version** (v6) - deploy this |
| `src/TCManager-v2.js` | Version history (reference only) |
| `src/TCManager-v3.js` | Version history (reference only) |
| `src/TCManager-v4.js` | Version history (reference only) |
| `src/TCManager-v5.js` | Version history (reference only) |
| `src/TCManager-v6.js` | Version history (reference only) |
| `src/appsscript.json` | Apps Script manifest |
| `deploy-tcmanager.sh` | **Automated deployment script** |
| `TC-MANAGER-DEPLOYMENT.md` | This documentation |
| `DEPLOY-TCMANAGER.md` | Old deployment guide (deprecated) |

---

**Created:** 2025-07-23
**Last Updated:** 2025-07-24
**Version:** v6
**Author:** Claude Code

---

## 🎯 Quick Reference Card

```bash
# Deploy to new spreadsheet
./deploy-tcmanager.sh <SCRIPT_ID>

# Get Script ID
Extensions → Apps Script → ⚙️ Project Settings → Copy Script ID

# Verify deployment
Refresh spreadsheet → Check for 🔧 TC Manager menu

# Test features
Insert TC Here    → Single row insert
Bulk Insert       → Multiple rows (1-100)
Delete TC         → Single/multiple row delete
Mark as Deprecated → Single/multiple row deprecate

# Clean up old code
Apps Script editor → Delete all files except TCManager.gs + appsscript.json
```

---

**Ready to deploy? Run:**
```bash
cd projects/qa-test-management
./deploy-tcmanager.sh <YOUR_SCRIPT_ID>
```
