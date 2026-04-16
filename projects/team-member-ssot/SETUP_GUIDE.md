# Setup Guide — Team Member SSOT Formatting Script

Quick guide to install the formatting script to the SSOT spreadsheet.

## Prerequisites

- Access to SSOT spreadsheet: `1PKZTLAhjcBAoBNcfhhOGdy7JUixEqvLbOwK5YAueSa4SdT5QzH1OHh_U`
- Edit permissions on the spreadsheet

## Installation Steps

### Step 1: Open SSOT Spreadsheet

1. Open browser and navigate to:
   ```
   https://docs.google.com/spreadsheets/d/1PKZTLAhjcBAoBNcfhhOGdy7JUixEqvLbOwK5YAueSa4SdT5QzH1OHh_U
   ```

2. Verify you're on the correct spreadsheet (should have "Team Member" tab)

### Step 2: Open Apps Script Editor

1. Click **Extensions** in the menu bar
2. Click **Apps Script**
3. A new tab will open with the script editor

### Step 3: Delete Existing Code (if any)

1. If there's any existing code in `Code.gs`, delete it
2. Or rename it to backup (e.g., `Code_backup.gs`)

### Step 4: Create TeamMemberFormat.js

1. Click **+** next to "Files"
2. Select **Script**
3. Name it: `TeamMemberFormat`
4. Copy entire content from `src/TeamMemberFormat.js`
5. Paste into the editor
6. Save (Ctrl+S or Cmd+S)

### Step 5: Create MenuFunctions.js

1. Click **+** next to "Files"
2. Select **Script**
3. Name it: `MenuFunctions`
4. Copy entire content from `src/MenuFunctions.js`
5. Paste into the editor
6. Save (Ctrl+S or Cmd+S)

### Step 6: Grant Permissions

1. Click **Run** (play button) on any function (e.g., `onOpen`)
2. Click **Review permissions**
3. Select your Google account
4. Click **Advanced**
5. Click **Go to [Project Name] (unsafe)**
6. Click **Allow**

### Step 7: Test Installation

1. Go back to the spreadsheet tab
2. Refresh the page (F5 or Cmd+R)
3. You should see new menu: **👥 Team Management**
4. Click it to see options:
   - 🎨 Format Team Member Tab
   - ➕ Add New Team Member
   - 📊 View Summary
   - ℹ️ About

### Step 8: Format the Tab

1. Click **👥 Team Management** → **🎨 Format Team Member Tab**
2. Click **YES** to confirm
3. Wait for completion message
4. Tab should now be formatted!

## Verification Checklist

After installation, verify:

- ✅ Menu "👥 Team Management" appears
- ✅ Header row is blue with white text
- ✅ Columns have optimal widths
- ✅ Status column has dropdowns
- ✅ Status "Onboard" is green
- ✅ Status "Resign" is red
- ✅ Row 1 is frozen
- ✅ Filter is enabled

## Using the Script

### Format Tab

Anytime you need to reformat:
1. **Menu** → **👥 Team Management** → **🎨 Format Team Member Tab**

### Add New Member

To add someone:
1. **Menu** → **👥 Team Management** → **➕ Add New Team Member**
2. Enter name
3. Fill details in new row

### View Statistics

To see summary:
1. **Menu** → **👥 Team Management** → **📊 View Summary**

## Troubleshooting

### Menu doesn't appear

**Solution:**
1. Refresh the spreadsheet
2. Wait a few seconds
3. Check Apps Script editor for errors

### Permission error

**Solution:**
1. Go to Apps Script editor
2. Run `onOpen` function manually
3. Grant permissions again

### Formatting doesn't apply

**Solution:**
1. Check you're on "Team Member" tab
2. Make sure you have edit access
3. Try running again

### "Tab not found" error

**Solution:**
1. Verify tab name is exactly "Team Member"
2. Check for extra spaces in tab name
3. Rename tab if needed

## Alternative: Using clasp (Advanced)

If you want to use `clasp` for deployment:

1. **Get Script Project ID:**
   - Open Apps Script editor
   - Click ⚙️ (Project Settings)
   - Copy "Script ID"

2. **Update .clasp.json:**
   ```json
   {
     "scriptId": "PASTE_SCRIPT_ID_HERE",
     "rootDir": "./src"
   }
   ```

3. **Deploy:**
   ```bash
   cd projects/team-member-ssot
   clasp push
   ```

## Next Steps

After installation:

1. ✅ Format the tab (Menu → Format Team Member Tab)
2. ✅ Review formatted structure
3. ✅ Test adding a new member
4. ✅ View summary statistics
5. ✅ Setup KPI Tracker to sync from this SSOT

## Integration with KPI Tracker

Once Team Member tab is formatted:

1. Open KPI Tracker spreadsheet
2. Menu → **👥 Team Members (SSOT)** → **🔧 Test SSOT Connection**
3. If successful, run: **🔄 Sync from SSOT**
4. Team members will auto-populate in KPI Tracker!

---

Questions? Contact QA Team Lead.
