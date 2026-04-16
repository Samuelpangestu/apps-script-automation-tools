# Team Member Management

Professional team member management system that creates and manages Team Member tabs from scratch. Designed to be portable across different companies and projects.

## Purpose

This tool creates a complete Team Member management tab with:
- ✅ Professional structure and formatting
- ✅ Comprehensive member tracking (15 columns)
- ✅ Built-in data validation
- ✅ Color-coded status visualization
- ✅ Statistics and export capabilities
- ✅ Ready to be used by other tools (KPI Tracker, etc.)

## Features

### 🔨 Create from Scratch

**Complete Tab Creation:**
- Creates "Team Member" tab with full structure
- 15 columns for comprehensive tracking
- Sample data included for reference
- Can rebuild existing tabs (clears and recreates)
- Portable design for any company/project

**Professional Header:**
- Blue background (#1a73e8)
- White text, bold, centered
- Frozen for easy scrolling
- Clear column labels

**Smart Structure:**
- 100 rows pre-allocated
- Optimal column widths
- Professional formatting
- Filter views enabled

### ✅ Data Validation

**Dropdown Lists:**
- **Status**: Onboard, Tidak Ada Kabar, Digispark, Resign, Contract End
- **Role**: QA Team Lead, Senior QE, Quality Engineer, PIC QE, Lead Project, Intern QE, Security Engineer, UX Research
- **Level**: Lead, Senior, Middle, Junior, Intern

**Checkboxes:**
- Automation Skills
- VPN ABC Access
- VPN Huwawei Access

**Date Formatting:**
- Join Date: YYYY-MM-DD format

### 🎨 Conditional Formatting

**Status Color Coding:**
- 🟢 **Green** (Onboard): Active team members
- 🟡 **Yellow** (Tidak Ada Kabar): Inactive/pending
- 🔴 **Red** (Resign/Contract End/Digispark): Terminated

### ➕ Team Management

**Add Team Member:**
- One-click add via menu
- Default values:
  - Join Date: Today
  - Status: Onboard
  - Role: Quality Engineer
- Professional row formatting

**Export Data:**
- Export to CSV format
- Includes all columns and data
- Creates file in same folder

### 📊 Summary Statistics

View live statistics:
- Total members (active/inactive)
- Count by status
- Count by role
- Count by level

## Installation

### Method 1: Manual Copy-Paste (Recommended)

1. **Open Target Spreadsheet**
   - Open any Google Spreadsheet where you want team member management

2. **Open Apps Script Editor**
   - Extensions → Apps Script

3. **Copy Files**
   - Create `TeamMemberManagement.js`: Copy content from `src/TeamMemberManagement.js`
   - Create `MenuFunctions.js`: Copy content from `src/MenuFunctions.js`

4. **Save & Refresh**
   - Save the project (Ctrl/Cmd + S)
   - Refresh the spreadsheet
   - New menu "👥 Team Management" will appear

### Method 2: Using clasp (For Developers)

1. **Get Script ID**
   - Open target spreadsheet
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
   cd projects/team-member-management
   clasp push
   ```

## Usage

### First Time Setup

1. **Open Your Spreadsheet**
2. **Menu** → **👥 Team Management** → **🔨 Create Team Member Tab**
3. Click **YES** to confirm
4. Wait for tab creation to complete
5. Done! "Team Member" tab is created with sample data

**Note:** If the tab already exists, it will be cleared and rebuilt from scratch.

### Adding Team Members

1. **Menu** → **👥 Team Management** → **➕ Add Team Member**
2. New row is added with default values
3. Fill in the member details

### Viewing Statistics

1. **Menu** → **👥 Team Management** → **📊 View Summary**
2. See total members, status distribution, role distribution, level distribution

### Exporting Data

1. **Menu** → **👥 Team Management** → **📤 Export Data**
2. Click **YES** to confirm
3. CSV file created in same folder as spreadsheet

## Column Structure

| Column | Name | Width | Type | Description |
|--------|------|-------|------|-------------|
| A | Name | 200px | Text | Full name of team member |
| B | Join Date | 100px | Date | Date joined the team (YYYY-MM-DD) |
| C | Role | 180px | Dropdown | QA role (QA Team Lead, Senior QE, etc.) |
| D | Level | 100px | Dropdown | Seniority level (Lead, Senior, Middle, Junior, Intern) |
| E | Project Assignment | 250px | Text | Current projects (comma-separated) |
| F | NIP | 100px | Text | Employee ID number |
| G | Email | 250px | Email | Work email address |
| H | Division/Unit | 150px | Text | Department or division |
| I | Status | 130px | Dropdown | Employment status |
| J | Automation | 100px | Checkbox | Has automation skills (TRUE/FALSE) |
| K | Slack Username | 150px | Text | Slack handle (@username) |
| L | WhatsApp Number | 140px | Text | Contact number |
| M | Role AMS | 150px | Text | AMS-specific role |
| N | VPN ABC | 100px | Checkbox | Has VPN ABC access (TRUE/FALSE) |
| O | VPN Huwawei | 120px | Checkbox | Has VPN Huwawei access (TRUE/FALSE) |

## Status Values

- **Onboard** (Active) - 🟢 Green background
- **Tidak Ada Kabar** (Pending) - 🟡 Yellow background
- **Digispark** (Inactive) - 🔴 Red background
- **Resign** (Inactive) - 🔴 Red background
- **Contract End** (Inactive) - 🔴 Red background

## Role Values

Available QA/QE roles:
- **QA Team Lead**: Team lead position
- **Senior Quality Engineer**: Senior QE
- **Quality Engineer**: Standard QE role
- **PIC QE**: Project-in-charge QE
- **Lead Project**: Project lead
- **Intern Quality Engineer**: Intern/trainee
- **Security Engineer**: Security testing
- **UX Research**: UX research role

## Level Values

Available seniority levels:
- **Lead**: Team/project lead
- **Senior**: Senior level
- **Middle**: Mid-level
- **Junior**: Junior level
- **Intern**: Internship/trainee

## Integration

This tool can be used as:
- **Standalone**: Independent team member management
- **SSOT**: Single source of truth for other tools
- **KPI Tracker Integration**: Sync team members to KPI Tracker
- **QA Dashboard**: Central team member data
- **Portable**: Use across different companies/projects

## File Structure

```
projects/team-member-management/
├── src/
│   ├── appsscript.json              # Apps Script manifest
│   ├── TeamMemberManagement.js      # Core tab creation & management
│   └── MenuFunctions.js             # Menu UI handlers
├── .clasp.json                      # Clasp deployment config
├── README.md                        # This file
└── SETUP_GUIDE.md                   # Detailed setup guide
```

## Maintenance

### Rebuilding Tab

To rebuild the entire tab from scratch:

**Menu** → **👥 Team Management** → **🔨 Create Team Member Tab**

This will:
- Clear all existing data
- Recreate the tab structure
- Add sample data
- Reapply all formatting and validation

**Warning:** This deletes all existing data! Export first if needed.

### Customizing

**Update Dropdown Values:**
Edit `TeamMemberManagement.js` → `addDataValidation()` function to modify:
- Status options
- Role options
- Level options

**Update Color Rules:**
Edit `TeamMemberManagement.js` → `addConditionalFormatting()` function to modify status colors.

**Update Columns:**
Edit `TeamMemberManagement.js` → `COLUMNS` constant to add/remove/modify columns.

## Tips

1. **Use Sample Data**: Review sample data after first creation to understand structure
2. **Export Before Rebuild**: Always export data before rebuilding the tab
3. **Use Dropdowns**: Always use dropdowns for Status, Role, and Level for consistency
4. **Check Summary**: Regularly view summary to verify data integrity
5. **Portable Design**: This tool can be deployed to any spreadsheet

## Troubleshooting

**Menu not appearing:**
- Refresh the spreadsheet (F5 or Ctrl/Cmd + R)
- Check if scripts are installed correctly in Apps Script editor
- Try closing and reopening the spreadsheet
- Check browser console for errors

**Tab creation fails:**
- Verify you have edit access to the spreadsheet
- Check if sheet name "Team Member" is available
- Try running from Apps Script editor directly

**Validation not working:**
- Ensure values match dropdown options exactly (case-sensitive)
- Rebuild tab to reapply validation rules
- Check for hidden spaces in cell values

## Support

For issues or questions:
1. **Help Dialog**: Menu → Team Management → Help
2. **About Dialog**: Menu → Team Management → About
3. Review this README and SETUP_GUIDE.md
4. Contact QA Team

## Version History

**v1.0.0** (Initial Release)
- Create Team Member tab from scratch
- 15 comprehensive columns
- Data validation (dropdowns, checkboxes)
- Conditional formatting (status colors)
- Add team member functionality
- View summary statistics
- Export to CSV
- Portable design for multi-company use

---

Built with professionalism for QA Teams everywhere
