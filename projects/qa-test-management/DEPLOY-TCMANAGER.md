# 🚀 Deploy TC Manager to QATM Testing

**QATM Testing ID:** `1MU8JF5Ve39LC2VG4hc2O9lEcaaAzQUNYFHE4Xu-F3zH31BD11Ok1AIbm`

---

## ✅ Step-by-Step Deployment

### **1. Open QATM Testing Spreadsheet**

Link: https://docs.google.com/spreadsheets/d/1MU8JF5Ve39LC2VG4hc2O9lEcaaAzQUNYFHE4Xu-F3zH31BD11Ok1AIbm/edit

---

### **2. Open Apps Script Editor**

1. Click menu: **Extensions** → **Apps Script**
2. Apps Script editor akan terbuka di tab baru

---

### **3. Clean Up Existing Code (PENTING!)**

**Delete semua code/menu yang tidak terpakai:**

Anda bilang ada menu-menu yang sudah tidak terpakai. Mari kita bersihkan:

1. Di Apps Script editor, cek file-file yang ada
2. **BACKUP dulu** (optional): Copy existing code ke notepad
3. **Delete/Comment out** menu yang tidak terpakai di function `onOpen()`

**Contoh cleanup `onOpen()`:**

```javascript
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // TC Manager Menu - BARU (keep this)
  ui.createMenu('🔧 TC Manager')
    .addItem('➕ Add New TC (at bottom)', 'addNewTC')
    // ... etc

  // DELETE atau COMMENT menu lama yang tidak terpakai:
  // ui.createMenu('Old Menu')  ← DELETE THIS
  //   .addItem('...', '...')
}
```

---

### **4. Add TC Manager Code**

**Option A: Create New File (Recommended)**

1. Di Apps Script editor, click **+** di sebelah "Files"
2. Pilih **Script**
3. Name: `TCManager`
4. Copy-paste code dari file: `TCManager-Clean.js`

**Option B: Add to Existing File**

1. Buka existing `Code.gs` atau file utama
2. Append code dari `TCManager-Clean.js` di bawah

---

### **5. Code to Copy**

File location: `projects/qa-test-management/src/TCManager-Clean.js`

**Full code ada di file tersebut** (311 lines, sudah clean tanpa bloat)

---

### **6. Save & Deploy**

1. **Save** (Ctrl/Cmd + S)
2. Click **Run** → Select `onOpen` → **Run** (untuk test)
3. **Authorize** kalau diminta permission
4. **Close** Apps Script editor
5. **Refresh** QATM Testing spreadsheet

---

### **7. Verify Deployment**

Setelah refresh, cek menu bar di spreadsheet:

✅ **Menu baru muncul:** `🔧 TC Manager`

Menu contents:
- ➕ Add New TC (at bottom)
- 📋 Insert TC Here (at cursor)
- ───────────────
- 🗑️ Delete TC (current row)
- ⚠️ Mark as Deprecated
- ───────────────
- ℹ️ Help
  - 📖 How to Use
  - ⚡ Why Use This?

---

## 🧪 Testing the Functions

### **Test 1: Add New TC (Safest)**

1. Go to `TC_Master` sheet
2. Click menu: `🔧 TC Manager` → `➕ Add New TC`
3. Dialog muncul dengan info row baru
4. **Verify:**
   - ✅ New row di TC_Master (di bottom)
   - ✅ New row di TC_Execution (same position)
   - ✅ Formula di TC_Execution columns B-G copied

---

### **Test 2: Insert TC at Position**

1. Select row di tengah TC_Master (misal row 5)
2. Click: `🔧 TC Manager` → `📋 Insert TC Here`
3. Confirm dialog
4. **Verify:**
   - ✅ New row inserted BEFORE row 5
   - ✅ TC_Execution row inserted di posisi yang sama
   - ✅ Rows below shifted down di BOTH sheets

---

### **Test 3: Mark as Deprecated**

1. Select TC yang mau di-deprecate
2. Click: `🔧 TC Manager` → `⚠️ Mark as Deprecated`
3. **Verify:**
   - ✅ Scenario prefix: `[DEPRECATED] ...`
   - ✅ Row jadi gray color

---

### **Test 4: Delete TC (Use with Caution)**

1. Select TC yang mau di-delete
2. Click: `🔧 TC Manager` → `🗑️ Delete TC`
3. Read warning, confirm
4. **Verify:**
   - ✅ Row deleted from TC_Master
   - ✅ Row deleted from TC_Execution
   - ✅ Rows below shifted up di BOTH sheets

---

## ⚠️ Important Notes

### **DO:**
- ✅ Use TC Manager menu untuk add/insert/delete
- ✅ Use "Add New TC" untuk most cases (safest)
- ✅ Use "Mark as Deprecated" instead of delete

### **DON'T:**
- ❌ Manual insert/delete rows di TC_Master
- ❌ Manual insert/delete rows di TC_Execution
- ❌ Delete TC yang sudah ada test results (use deprecate)

---

## 🐛 Troubleshooting

### **Menu tidak muncul setelah refresh:**
- Close and re-open spreadsheet
- Hard refresh browser (Ctrl+Shift+R)
- Check Apps Script execution log for errors

### **Error saat run function:**
- Check TC_Master and TC_Execution sheet names (exact match)
- Check row structure (TC_Master starts row 3, TC_Execution row 9)
- Check Apps Script permissions granted

### **Formula tidak ke-copy:**
- Check source row has formulas di columns B-G
- Manual copy formula from row above if needed

---

## 📊 Sheet Structure Requirements

Script expects:
- **TC_Master**: Data starts at row 3, column C = TC_ID
- **TC_Execution**: Data starts at row 9, columns B-G = formulas

If structure berbeda, adjust constants di script:
```javascript
const execStartRow = 9;  // Change if TC_Execution starts at different row
```

---

## 🎯 Next Steps After Testing

Kalau testing berhasil:

1. ✅ Document hasil testing
2. ✅ Train users cara pakai TC Manager
3. ✅ Deploy ke QATM Template
4. ✅ Broadcast info ke team

---

## 📞 Support

Issues? Check:
1. Apps Script execution logs (View → Execution log)
2. Browser console (F12)
3. Claude Code for debugging

---

**Created:** 2025-01-23
**Version:** 1.0 (Clean)
