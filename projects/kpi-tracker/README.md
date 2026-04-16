# KPI Tracker — QA Department PERURI

Maintainable and flexible KPI tracking tool for QA team with integrated 360 Review system.

## Features

- ✅ **Centralized Configuration**: Manage team members in one place
- ✅ **Maintainable KPI Definitions**: Update targets, formulas, or add new KPIs anytime
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
├── Config                    # Team members management
├── KPI Definition            # Maintainable KPI definitions per role
├── Form Responses 1          # Auto-created by 360 Review Form
└── (Future: KPI Tracker)     # Actual KPI tracking per period
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

### Managing Team Members

1. Open **Config** tab
2. Add/edit members in the table (row 6 onwards)
3. Set **Status** to "Aktif" or "Non-Aktif"
4. Changes automatically reflect everywhere

### Updating KPI Definitions

1. Open **KPI Definition** tab
2. Edit targets, formulas, or data sources directly
3. Add new rows for new KPIs
4. Changes are immediately effective

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
│   ├── Config.js             # Team member management
│   ├── KPIDefinition.js      # KPI definitions
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
  - 360 Review Form integration
  - Weighted scoring calculation

## Support

For issues or questions:
1. Check the User Guide: Menu → Help → User Guide
2. Review this README
3. Contact QA Team Lead

---

Built with ❤️ for QA Team PERURI
