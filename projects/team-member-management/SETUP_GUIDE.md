# Setup Guide — Team Member Management Tool

Complete guide to install and use the Team Member Management tool in any Google Spreadsheet.

## Prerequisites

- Any Google Spreadsheet where you want team member management
- Edit permissions on the spreadsheet
- Google account with Apps Script access

## Installation Steps

### Step 1: Open Your Spreadsheet

1. Open any Google Spreadsheet where you want team member management
2. Or create a new spreadsheet: https://sheets.new

### Step 2: Open Apps Script Editor

1. Click **Extensions** in the menu bar
2. Click **Apps Script**
3. A new tab will open with the Apps Script editor

### Step 3: Prepare the Editor

1. If there's default `Code.gs` with placeholder code, you can:
   - Delete it, OR
   - Keep it and add new files alongside it

### Step 4: Create TeamMemberManagement.js

1. Click **+** next to "Files" (left sidebar)
2. Select **Script**
3. Name it: `TeamMemberManagement`
4. Open this repo's file: `projects/team-member-management/src/TeamMemberManagement.js`
5. Copy the entire content
6. Paste into the Apps Script editor
7. Save (Ctrl+S or Cmd+S)

### Step 5: Create MenuFunctions.js

1. Click **+** next to "Files"
2. Select **Script**
3. Name it: `MenuFunctions`
4. Open this repo's file: `projects/team-member-management/src/MenuFunctions.js`
5. Copy the entire content
6. Paste into the Apps Script editor
7. Save (Ctrl+S or Cmd+S)

### Step 6: Grant Permissions

1. In Apps Script editor, select `onOpen` function from dropdown (top toolbar)
2. Click **Run** (play button ▶️)
3. Click **Review permissions**
4. Select your Google account
5. Click **Advanced**
6. Click **Go to [Untitled project] (unsafe)**
7. Click **Allow**
8. Wait for execution to complete

### Step 7: Test Installation

1. Go back to the spreadsheet tab
2. Refresh the page (F5 or Cmd+R)
3. Wait a few seconds
4. You should see new menu: **👥 Team Management**
5. Click it to see options:
   - 🔨 Create Team Member Tab
   - ➕ Add Team Member
   - 📊 View Summary
   - 📤 Export Data
   - ℹ️ About
   - ❓ Help

### Step 8: Create the Team Member Tab

1. Click **👥 Team Management** → **🔨 Create Team Member Tab**
2. Click **YES** to confirm
3. Wait for completion message (5-10 seconds)
4. A new "Team Member" tab will be created with:
   - Professional blue header
   - 15 columns
   - Sample data (4 team members)
   - Data validation
   - Conditional formatting

## Verification Checklist

After creating the tab, verify:

- ✅ Menu "👥 Team Management" appears in menu bar
- ✅ "Team Member" tab exists
- ✅ Header row is blue (#1a73e8) with white text
- ✅ 15 columns (A-O) with proper labels
- ✅ 4 sample team members in rows 2-5
- ✅ Status column (I) has dropdown validation
- ✅ Role column (C) has dropdown validation
- ✅ Level column (D) has dropdown validation
- ✅ Status "Onboard" has green background
- ✅ Row 1 (header) is frozen
- ✅ Filter views are enabled

## Using the Tool

### Create/Rebuild Tab

To create a new tab or rebuild existing:
1. **Menu** → **👥 Team Management** → **🔨 Create Team Member Tab**
2. Click **YES** to confirm
3. **Warning:** This deletes all existing data in "Team Member" tab!

### Add Team Member

To add a new team member:
1. **Menu** → **👥 Team Management** → **➕ Add Team Member**
2. New row is added with defaults (Join Date: today, Status: Onboard, Role: Quality Engineer)
3. Fill in the remaining details manually

### View Statistics

To see team summary:
1. **Menu** → **👥 Team Management** → **📊 View Summary**
2. View statistics:
   - Total members (active/inactive)
   - Count by status
   - Count by role
   - Count by level

### Export Data

To export team member data:
1. **Menu** → **👥 Team Management** → **📤 Export Data**
2. Click **YES** to confirm
3. CSV file is created in the same folder as your spreadsheet

## Troubleshooting

### Menu doesn't appear

**Symptoms:** After refreshing, "👥 Team Management" menu is not visible

**Solutions:**
1. Refresh the spreadsheet again (F5 or Cmd+R)
2. Wait 10-15 seconds for menu to load
3. Check Apps Script editor for errors (View → Logs)
4. Try closing and reopening the spreadsheet
5. Clear browser cache and retry

### Permission error when running

**Symptoms:** "Authorization required" or permission popup

**Solutions:**
1. Go to Apps Script editor
2. Select `onOpen` from function dropdown
3. Click Run (▶️)
4. Grant permissions again following Step 6

### Create Tab fails

**Symptoms:** Error when clicking "Create Team Member Tab"

**Solutions:**
1. Check you have edit access to the spreadsheet
2. Verify no sheet is named "Team Member" with protection
3. Try running `createTeamMemberTab` directly from Apps Script editor
4. Check execution logs in Apps Script (View → Logs)

### Data validation not working

**Symptoms:** Dropdowns don't appear or don't work

**Solutions:**
1. Rebuild the tab (Menu → Create Team Member Tab)
2. Check cell values match dropdown options exactly (case-sensitive)
3. Verify no manual edits to validation ranges

### Conditional formatting not applying

**Symptoms:** Status colors not showing

**Solutions:**
1. Rebuild the tab to reapply rules
2. Verify Status column values match exactly: "Onboard", "Resign", etc.
3. Check for extra spaces in status values
4. Clear any conflicting conditional formatting rules

## Alternative: Using clasp (Advanced)

For developers who want to use `clasp` to push code updates:

### Prerequisites
```bash
npm install -g @google/clasp
clasp login
```

### Setup

1. **Get Script Project ID:**
   - Open your spreadsheet
   - Extensions → Apps Script
   - Click ⚙️ (Project Settings)
   - Copy "Script ID"

2. **Update .clasp.json:**
   ```json
   {
     "scriptId": "YOUR_SCRIPT_ID_HERE",
     "rootDir": "./src"
   }
   ```

3. **Deploy:**
   ```bash
   cd projects/team-member-management
   clasp push
   ```

4. **Pull changes:**
   ```bash
   clasp pull
   ```

## Next Steps

After successful installation:

1. ✅ Create the Team Member tab
2. ✅ Review sample data and structure
3. ✅ Delete sample data and add real team members
4. ✅ Test adding a new member via menu
5. ✅ Test viewing summary statistics
6. ✅ Test exporting data to CSV

## Advanced Usage

### Customizing Columns

To add/modify columns:

1. Edit `TeamMemberManagement.js` → `COLUMNS` constant
2. Add your new column definition:
   ```javascript
   NEW_COL: { index: 16, letter: 'P', width: 150, header: 'New Column' }
   ```
3. Update `createHeader()` to include new header
4. Update `addSampleData()` to include sample data for new column
5. Redeploy and rebuild tab

### Customizing Dropdowns

To modify dropdown values:

1. Edit `TeamMemberManagement.js` → `addDataValidation()` function
2. Modify arrays for Status, Role, or Level
3. Redeploy and rebuild tab

### Integration as SSOT

To use this as Single Source of Truth for other tools:

1. Deploy this tool to your central spreadsheet
2. Note the Spreadsheet ID
3. In other tools (e.g., KPI Tracker), configure SSOT sync
4. Use the Spreadsheet ID to fetch team member data

## Best Practices

1. **Backup First**: Always export data before rebuilding tab
2. **Consistent Data**: Use dropdowns to maintain data consistency
3. **Regular Updates**: Keep team member information up to date
4. **Access Control**: Limit edit access to prevent accidental changes
5. **Version Control**: Keep script code in git repository

## Support

For questions or issues:

1. Check **Help** dialog: Menu → Team Management → Help
2. Check **About** dialog: Menu → Team Management → About
3. Review README.md in this repository
4. Contact your QA Team Lead or project administrator

---

**Pro Tip:** This tool is designed to be portable. You can copy the scripts to any Google Spreadsheet in any company and it will work the same way!
