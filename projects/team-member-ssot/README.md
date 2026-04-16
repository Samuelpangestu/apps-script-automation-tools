# Team Member SSOT — Formatting & Management

Professional formatting and management tools for the centralized Team Member spreadsheet (Single Source of Truth).

## Purpose

This script provides formatting and management utilities for the Team Member SSOT spreadsheet, making it:
- ✅ Clean and professional looking
- ✅ Easy to read and navigate
- ✅ Manageable with built-in validations
- ✅ Consistent across all data
- ✅ Ready to be used by other tools (KPI Tracker, etc.)

## Features

### 🎨 Professional Formatting

**Header Styling:**
- Blue background (#1a73e8)
- White text, bold, centered
- Proper column headers
- Frozen for easy scrolling

**Data Rows:**
- Alternating colors (white/light gray)
- Optimal column widths
- Proper alignment (center for dates, status, etc.)
- Borders for clear separation

**Conditional Formatting:**
- 🟢 **Green** (Onboard): Active team members
- 🟡 **Yellow** (Tidak Ada Kabar): Pending status
- 🔴 **Red** (Resign/Contract End/Digispark): Inactive

### ✅ Data Validation

**Dropdown Lists:**
- **Status**: Onboard, Tidak Ada Kabar, Digispark, Resign, Contract End
- **Role**: QA Team Lead, Senior QE, QE, PIC QE, Lead Project, Intern QE, Security Engineer, UX Research

**Checkboxes:**
- Automation (TRUE/FALSE)
- VPN ABC (TRUE/FALSE)
- VPN Huwawei (TRUE/FALSE)

**Date Formatting:**
- Join Date: YYYY-MM-DD format

### ➕ Easy Team Management

**Add New Team Member:**
- One-click add via menu
- Auto-formatted row
- Default values:
  - Join Date: Today
  - Status: Onboard
  - Role: Quality Engineer
  - All checkboxes: FALSE

**Auto-placement:**
- Inserts before section headers
- Maintains proper row grouping

### 📊 Summary Statistics

View live statistics:
- Total team members
- Active vs Inactive count
- Count by status
- Count by role

## Installation

### Method 1: Manual Copy-Paste (Recommended)

1. **Open SSOT Spreadsheet**
   ```
   https://docs.google.com/spreadsheets/d/1PKZTLAhjcBAoBNcfhhOGdy7JUixEqvLbOwK5YAueSa4SdT5QzH1OHh_U
   ```

2. **Open Apps Script Editor**
   - Extensions → Apps Script

3. **Copy Files**
   - Create `TeamMemberFormat.js`: Copy content from `src/TeamMemberFormat.js`
   - Create `MenuFunctions.js`: Copy content from `src/MenuFunctions.js`

4. **Save & Refresh**
   - Save the project
   - Refresh the spreadsheet
   - New menu "👥 Team Management" will appear

### Method 2: Using clasp (For Developers)

**Note**: You need to update `.clasp.json` with the actual Apps Script project ID bound to the SSOT spreadsheet.

1. **Get Script ID**
   - Open SSOT spreadsheet
   - Extensions → Apps Script
   - Project Settings → Copy "Script ID"

2. **Update .clasp.json**
   ```json
   {
     "scriptId": "YOUR_ACTUAL_SCRIPT_ID_HERE",
     "rootDir": "./src"
   }
   ```

3. **Deploy**
   ```bash
   cd projects/team-member-ssot
   clasp push
   ```

## Usage

### First Time Setup

1. **Open SSOT Spreadsheet**
2. **Menu** → **👥 Team Management** → **🎨 Format Team Member Tab**
3. Click **YES** to confirm
4. Wait for formatting to complete
5. Done! Tab is now formatted

### Adding New Team Member

1. **Menu** → **👥 Team Management** → **➕ Add New Team Member**
2. Enter full name
3. Click OK
4. Fill in remaining details in the new row

### Viewing Statistics

1. **Menu** → **👥 Team Management** → **📊 View Summary**
2. See total members, status distribution, role distribution

## Column Structure

| Column | Name | Width | Type | Description |
|--------|------|-------|------|-------------|
| A | Name | 200px | Text | Full name |
| B | Join Date | 100px | Date | Start date (YYYY-MM-DD) |
| C | Title | 150px | Text | Job title |
| D | Lead/PIC | 120px | Text | Project lead designation |
| E | Project | 250px | Text | Assigned projects |
| F | NP | 80px | Text | Employee number |
| G | Email | 250px | Email | Primary work email |
| H | Email 2 | 250px | Email | Secondary email |
| I | Status | 120px | Dropdown | Employment status |
| J | Automation | 100px | Checkbox | Automation access |
| K | Github | 180px | Text | GitHub username |
| L | Phone | 130px | Text | Contact number |
| M | Role | 200px | Dropdown | Job role |
| N | VPN ABC | 90px | Checkbox | VPN access |
| O | VPN Huwawei | 120px | Checkbox | VPN access |

## Status Values

- **Onboard** (Active) - 🟢 Green background
- **Tidak Ada Kabar** (Pending) - 🟡 Yellow background
- **Digispark** (Inactive) - 🔴 Red background
- **Resign** (Inactive) - 🔴 Red background
- **Contract End** (Inactive) - 🔴 Red background

## Role Values

QA/QE Roles (synced to KPI Tracker):
- QA Team Lead
- Senior Quality Engineer
- Quality Engineer
- PIC QE
- Lead Project
- Intern Quality Engineer

Other Roles (not synced):
- Security Engineer
- UX Research

## Integration

This SSOT is used by:
- **KPI Tracker**: Auto-syncs active QA/QE members
- **QA Dashboard**: Team member data
- **Other QA Tools**: Centralized team data

## File Structure

```
projects/team-member-ssot/
├── src/
│   ├── appsscript.json          # Manifest
│   ├── TeamMemberFormat.js      # Formatting functions
│   └── MenuFunctions.js         # Menu UI
├── .clasp.json                  # Deployment config
└── README.md                    # This file
```

## Maintenance

### Re-formatting

Run formatting anytime to:
- Restore column widths
- Fix row colors
- Reapply conditional formatting
- Clean up structure

**Menu** → **👥 Team Management** → **🎨 Format Team Member Tab**

### Updating Validation

Edit `TeamMemberFormat.js` → `addDataValidation()` function to add/modify dropdown values.

### Updating Conditional Formatting

Edit `TeamMemberFormat.js` → `addConditionalFormatting()` function to modify color rules.

## Tips

1. **Keep Section Headers**: Don't delete rows like "MBG", "Riset/Security" - they organize the data
2. **Use Dropdowns**: Always use dropdowns for Status and Role for consistency
3. **Run Formatting**: After bulk changes, run formatting to restore structure
4. **Check Summary**: Use summary to verify data integrity
5. **Freeze Panes**: Header row is frozen for easy navigation

## Troubleshooting

**Menu not appearing:**
- Refresh the spreadsheet
- Check if script is installed correctly
- Try closing and reopening the spreadsheet

**Formatting not applying:**
- Check if you have edit access
- Verify you're on the "Team Member" tab
- Try running again

**Conditional formatting not working:**
- Make sure Status column values match exactly (case-sensitive)
- Re-run formatting to reapply rules

## Support

For issues or questions:
1. Check the About dialog: Menu → Team Management → About
2. Review this README
3. Contact QA Team Lead

---

Built with ❤️ for QA Team PERURI
