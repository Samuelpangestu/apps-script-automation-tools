# KPI Tracker — QA Department PERURI

Maintainable and flexible KPI tracking tool for QA team with integrated 360 Review system.

## Features

- ✅ **SSOT Integration**: Sync team members from centralized Team Member spreadsheet
- ✅ **Centralized Configuration**: Manage team members in one place
- ✅ **Maintainable KPI Definitions**: Update targets, formulas, or add new KPIs anytime
- ✅ **Period-based Tracking**: Sprint/Monthly/Quarterly KPI tracking with auto-calculation
- ✅ **Dashboard & Analytics**: Overview of all periods with success rates
- ✅ **Automated 360 Review**: Google Form integration with weighted scoring
- ✅ **Flexible & Scalable**: Easy to adapt when team or KPIs change
- ✅ **Pre-configured**: Ready with QA Department KPIs

## Quick Start

### First Time Setup

1. Open the Google Sheet
2. Click **🎯 KPI Tracker** menu → **Setup** → **Initial Setup**
3. Follow the prompts to create all tabs
4. Update team members in **Config** tab
5. Create 360 Review Form from menu

### Structure

```
📁 Tabs
├── Dashboard                 # Overview & summary of all periods
├── Config                    # Team members management
├── KPI Definition            # Maintainable KPI definitions per role
├── KPI - <Period>            # Actual KPI tracking per period (dynamic)
└── Form Responses 1          # Auto-created by 360 Review Form
```

## Roles & KPIs

### QA Team Lead (CoE)
- On-time Delivery Rate (≥ 80%)
- Defect Detection Efficiency (≥ 75%)
- CoE Tool & Template Adoption Rate (≥ 80%)
- Team Training Completion Rate (≥ 85%)

### QA Lead (Project Dedicated)
- Test Automation Coverage (≥ 70%)
- Defect Escape Rate (≤ 10%)
- 360 Review Score (≥ 3.5/5)
- Test Pass Rate (≥ 85%)

### PIC Project (QE + Koordinator)
- Test Automation Coverage (≥ 65%)
- Defect Escape Rate (≤ 10%)
- 360 Review Score (≥ 3.5/5)
- Test Pass Rate (≥ 80%)

### Quality Engineer (QE)
- Test Automation Coverage (≥ 60%)
- Defect Escape Rate (≤ 10%)
- 360 Review Score (≥ 3.5/5)
- Test Pass Rate (≥ 80%)

## Usage

### Syncing Team Members from SSOT

**SSOT (Single Source of Truth)** = Centralized Team Member spreadsheet

**One-click Sync:**
1. Menu → **👥 Team Members (SSOT)** → **🔄 Sync from SSOT**
2. Confirm sync
3. Team members auto-populate in Config tab

**What gets synced:**
- Name, Role, Email, Status
- Join date (Start Date)
- Only QA/QE roles (filters out Security, UX, etc.)
- Auto-maps SSOT roles to KPI Tracker roles:
  - `Senior Quality Engineer` → `QA Team Lead (CoE)`
  - `QA Team Lead` → `QA Team Lead (CoE)`
  - `Quality Engineer` → `Quality Engineer (QE)`
  - `Intern Quality Engineer` → `Quality Engineer (QE)`
  - `PIC QE` → `PIC Project (QE + Koordinator)`
  - `Lead Project` → `QA Lead (Project Dedicated)`
- Auto-maps status:
  - `Onboard` → `Aktif`
  - Others → `Non-Aktif`

**Testing Connection:**
- Menu → **👥 Team Members (SSOT)** → **🔧 Test SSOT Connection**

**View SSOT Info:**
- Menu → **👥 Team Members (SSOT)** → **ℹ️ Show SSOT Info**
- Shows: Total members, role distribution, status distribution

### Managing Team Members Manually

1. Open **Config** tab
2. Add/edit members in the table (row 6 onwards)
3. Set **Status** to "Aktif" or "Non-Aktif"
4. Changes automatically reflect everywhere

**Note**: Manual edits will be overwritten when syncing from SSOT

### Updating KPI Definitions

1. Open **KPI Definition** tab
2. Edit targets, formulas, or data sources directly
3. Add new rows for new KPIs
4. Changes are immediately effective

### Tracking KPIs per Period

**Creating Period Tracker:**
1. Menu → **KPI Tracking** → **Create Period Tracker**
2. Enter period name (e.g., "Sprint 24", "Jan 2026", "Q1 2026")
3. Select period type:
   - **Sprint**: Tracks sprint-based KPIs only
   - **Monthly**: Tracks monthly KPIs only
   - **Quarterly**: Tracks quarterly KPIs only
4. New sheet created (e.g., "KPI - Sprint 24")

**Filling Actual Values:**
1. Open the period tracker sheet
2. Fill **Actual** column with real measured values
3. **Achievement %** auto-calculates
4. **Status** auto-updates (✅ Met / ❌ Not Met / ⚪ Pending)
5. Summary stats update automatically

**Viewing Dashboard:**
1. Open **Dashboard** tab
2. See overview of all periods
3. Check success rates
4. Refresh: Menu → **KPI Tracking** → **Refresh Dashboard**

### 360 Review Process

**Creating Form (Once):**
1. Menu → **360 Review** → **Create Review Form**
2. Form is auto-linked to spreadsheet
3. Share form URL with reviewers

**Updating Reviewees:**
1. Update team members in Config tab
2. Menu → **360 Review** → **Update Reviewee List**
3. Form automatically updates

**Scoring:**
- Weighted formula: TL 35% + PM 25% + Peer 15% + Self 10% + Other 15%
- Auto-calculated from form responses
- 5 criteria: Technical, Delivery, Communication, Leadership, Quality

## Maintenance

### When Team Changes

1. **New member joins:**
   - Add to Config tab
   - Set Status = "Aktif"
   - Run "Update Reviewee List" (if eligible for review)

2. **Member leaves:**
   - Set Status = "Non-Aktif" (keeps history)
   - OR delete row (removes from all tracking)
   - Run "Update Reviewee List"

### When KPI Changes

1. Open **KPI Definition** tab
2. Edit target value, formula, or data source
3. Add new KPI as new row
4. No code changes needed!

## File Structure

```
projects/kpi-tracker/
├── src/
│   ├── appsscript.json       # Apps Script manifest
│   ├── SSOTSync.js           # SSOT team member sync
│   ├── TeamConfig.js         # Team member management
│   ├── KPIDefinition.js      # KPI definitions
│   ├── KPITracker.js         # Period tracker creation
│   ├── Dashboard.js          # Dashboard & overview
│   ├── Review360.js          # 360 Review form & scoring
│   └── MenuFunctions.js      # Menu UI & setup
├── .clasp.json               # Deployment config
└── README.md                 # This file
```

## Deployment

```bash
# Deploy to Apps Script
cd projects/kpi-tracker
clasp push

# Pull latest from Apps Script
clasp pull
```

## Version History

- **v1.0.0** (2026-04-16): Initial release
  - Config tab with team member management
  - KPI Definition tab with all QA KPIs
  - Period tracker creation (Sprint/Monthly/Quarterly)
  - Dashboard with overview & summary
  - Auto-calculation of achievement % and status
  - 360 Review Form integration
  - Weighted scoring calculation

## Support

For issues or questions:
1. Check the User Guide: Menu → Help → User Guide
2. Review this README
3. Contact QA Team Lead

---

Built with ❤️ for QA Team PERURI
