# VAPT Integration Plan

## Overview

Integrasi data VAPT dari 2 sumber ke QA Dashboard:
1. **Ad Hoc VAPT** - VAPT untuk development/improvement
2. **Regular VAPT** - VAPT rutin per beberapa bulan

**Goals:**
- Tampilkan summary + table di Dashboard tab "VAPT"
- Daily tracking untuk trendline (History tab)
- Integrasi ke Web App dengan visualisasi

---

## Data Structure Analysis

### Ad Hoc VAPT (Range: C1:U30)
```
Columns (21 total):
- No
- Aplikasi
- PIC VAPT
- Scope
- VAPT Status (Done, In Progress, Not Started)
- Report

Findings (grouped by status):
├─ Ready to Retest: Critical, High, Medium, Low, Info (5 cols)
├─ Open: Critical, High, Medium, Low, Info (5 cols)
└─ Closed: Critical, High, Medium, Low, Info (5 cols)

- Prod (TRUE/FALSE)
- Formula Updated (TRUE/FALSE)
```

### Regular VAPT (Range: C1:AB30)
```
Columns (28 total):
- Aplikasi
- PIC VAPT
- PIC QA
- Product Owner
- VAPT MSSP Status (Done, In Progress, Not Started, Todo)
- MSSP Report
- MSSP Cheklist Status
- MSSP Checklist Report
- Internal VAPT Status
- Report
- Report to PMO
- MSSP Reported (TRUE/FALSE)
- Internal Reported (TRUE/FALSE)

Findings (grouped by status):
├─ Ready to Retest: Critical, High, Medium, Low, Info (5 cols)
├─ Open: Critical, High, Medium, Low, Info (5 cols)
└─ Closed: Critical, High, Medium, Low, Info (5 cols)

- Prod (TRUE/FALSE)
- Formula Updated (TRUE/FALSE)
```

### Common Fields:
- **Aplikasi** - Application name
- **PIC VAPT** - VAPT PIC
- **VAPT Status** - Status (Done, In Progress, Not Started, Todo)
- **Report** - Report link/name
- **Findings by Severity & Status**:
  - Ready to Retest: Critical, High, Medium, Low, Info
  - Open: Critical, High, Medium, Low, Info
  - Closed: Critical, High, Medium, Low, Info
- **Prod** - Production status
- **Formula Updated** - Update flag

---

## Dashboard Implementation

### 1. VAPT Tab (New)

**Layout:**

```
┌─────────────────────────────────────────────────────────────┐
│ VAPT DASHBOARD - SIPGN                                      │
│ Last Updated: [timestamp]                                   │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│ === SUMMARY METRICS ===                                     │
│                                                             │
│ Total Applications: 50                                      │
│ Total Findings: 234 (Critical: 5, High: 10, Med: 89...)    │
│                                                             │
│ By Status:                                                  │
│   - Ready to Retest: 15 (Crit: 1, High: 3, Med: 8, Low: 3) │
│   - Open: 45 (Crit: 2, High: 5, Med: 25, Low: 13)         │
│   - Closed: 174 (Crit: 2, High: 2, Med: 56, Low: 114)     │
│                                                             │
│ By VAPT Status:                                            │
│   - Done: 40 apps                                          │
│   - In Progress: 8 apps                                    │
│   - Not Started: 2 apps                                    │
│                                                             │
│ Production Status:                                          │
│   - In Production: 45 apps                                 │
│   - Not in Prod: 5 apps                                    │
│                                                             │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│ === DETAILED TABLE ===                                      │
│                                                             │
│ [Headers]                                                   │
│ Type | Aplikasi | PIC VAPT | Status | Ready to Retest |   │
│      |          |          |        | Open | Closed | Prod│
│                                                             │
│ [Data rows combined from Ad Hoc + Regular]                 │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

**Table Columns:**
1. **Type** - "Ad Hoc" or "Regular"
2. **Aplikasi** - Application name
3. **PIC VAPT** - VAPT PIC
4. **VAPT Status** - Done/In Progress/Not Started/Todo
5. **Ready to Retest** - Crit/High/Med/Low/Info counts
6. **Open** - Crit/High/Med/Low/Info counts
7. **Closed** - Crit/High/Med/Low/Info counts
8. **Prod** - Production status (✓/✗)
9. **Report** - Link to report

### 2. VAPT History Tab (New)

**Purpose:** Daily snapshots for trendline analysis

**Layout:**
```
Timestamp | Type | Total_Findings | Critical | High | Medium | Low | Info |
          |      | Ready_Retest   | Open     | Closed | Prod_Count | Apps_Done |
```

**Tracking:**
- Daily snapshot at refresh time
- Aggregate totals from both Ad Hoc + Regular
- Separate rows for Ad Hoc and Regular (for type-specific trends)
- Combined row for overall portfolio

---

## Code Implementation

### File: `src/MasterDashboard.js`

#### Function: `buildVAPT(ss)`

**Purpose:** Create VAPT tab with summary + table

**Steps:**
1. Get VAPT spreadsheet ID from Config
2. Fetch data from Ad Hoc VAPT tab (C1:U30)
3. Fetch data from Regular VAPT tab (C1:AB30)
4. Process and combine data:
   - Parse findings by severity and status
   - Calculate summary metrics
   - Combine into unified table
5. Create tab with headers
6. Write summary section (rows 1-20)
7. Write table section (rows 22+)
8. Apply formatting (colors, borders, conditional formatting)

**Data Processing:**
```javascript
// For each row from Ad Hoc VAPT:
{
  type: 'Ad Hoc',
  aplikasi: row[1],
  picVapt: row[2],
  status: row[4],
  readyToRetest: {
    critical: row[6],
    high: row[7],
    medium: row[8],
    low: row[9],
    info: row[10]
  },
  open: {
    critical: row[11],
    high: row[12],
    medium: row[13],
    low: row[14],
    info: row[15]
  },
  closed: {
    critical: row[16],
    high: row[17],
    medium: row[18],
    low: row[19],
    info: row[20]
  },
  prod: row[21],
  report: row[5]
}

// Similar for Regular VAPT (adjust column indices)
```

#### Function: `buildVAPTHistory(ss)`

**Purpose:** Create VAPT History tab for daily tracking

**Layout:**
```
Row 1: Headers
Row 2+: Historical data entries
```

**Columns:**
- Timestamp
- Type (Ad Hoc / Regular / Combined)
- Total Findings
- By Severity: Critical, High, Medium, Low, Info
- By Status: Ready to Retest, Open, Closed
- Prod Count
- Apps Done
- Apps In Progress

#### Function: `refreshVAPTData()`

**Purpose:** Fetch latest VAPT data and update dashboard

**Called by:** `refreshDashboard()`

**Steps:**
1. Get VAPT spreadsheet ID from Config
2. Fetch Ad Hoc VAPT data
3. Fetch Regular VAPT data
4. Update VAPT tab summary + table
5. Append daily snapshot to VAPT History

#### Function: `appendVAPTHistory_(ss, vaptData, timestamp)`

**Purpose:** Append daily snapshot to VAPT History

**Parameters:**
- `ss` - Dashboard spreadsheet
- `vaptData` - Processed VAPT data (Ad Hoc + Regular)
- `timestamp` - Current timestamp

**Steps:**
1. Calculate aggregate metrics for Ad Hoc
2. Calculate aggregate metrics for Regular
3. Calculate combined totals
4. Append 3 rows (Ad Hoc, Regular, Combined) to History

### File: `src/WebAppBackend.js`

#### Function: `getVAPTData_(ss)`

**Purpose:** Fetch VAPT data for Web App

**Returns:**
```javascript
{
  summary: {
    totalApps: 50,
    totalFindings: 234,
    bySeverity: {
      critical: 5,
      high: 10,
      medium: 89,
      low: 120,
      info: 10
    },
    byStatus: {
      readyToRetest: { critical: 1, high: 3, ... },
      open: { critical: 2, high: 5, ... },
      closed: { critical: 2, high: 2, ... }
    },
    byVaptStatus: {
      done: 40,
      inProgress: 8,
      notStarted: 2
    },
    prodCount: 45
  },
  table: [
    // Array of VAPT entries
    ['Ad Hoc', 'Portal SIPGN', 'Afin', 'Done', ...],
    ['Regular', 'Fleet Management', 'Harits', 'In Progress', ...],
    ...
  ],
  history: [
    // Array of historical data for charts
    ['2026-03-01', 'Combined', 234, 5, 10, 89, ...],
    ['2026-03-02', 'Combined', 235, 5, 10, 90, ...],
    ...
  ]
}
```

#### Update: `getDashboardData()`

Add VAPT data to main dashboard data:

```javascript
const vaptData = getVAPTData_(ss);

return {
  summary: summaryData,
  history: historyData,
  modules: modules,
  bugsTable: bugsTableData,
  vapt: vaptData,  // NEW
  timestamp: new Date().toISOString()
};
```

### File: `src/WebApp.html`

#### New Tab: "VAPT"

**Layout:**
```html
<div id="vapt-tab" class="tab-content">
  <!-- Summary Cards -->
  <div class="summary-section">
    <div class="card">Total Apps: <span id="vapt-total-apps"></span></div>
    <div class="card">Total Findings: <span id="vapt-total-findings"></span></div>
    <div class="card">Critical Open: <span id="vapt-critical-open"></span></div>
    <div class="card">High Open: <span id="vapt-high-open"></span></div>
  </div>

  <!-- Charts -->
  <div class="charts-section">
    <canvas id="vapt-severity-chart"></canvas>
    <canvas id="vapt-status-chart"></canvas>
    <canvas id="vapt-trend-chart"></canvas>
  </div>

  <!-- Table -->
  <div class="table-section">
    <table id="vapt-table">
      <!-- Populated by JavaScript -->
    </table>
  </div>
</div>
```

**Charts:**
1. **Severity Distribution** - Pie chart (Critical, High, Medium, Low, Info)
2. **Status Breakdown** - Stacked bar chart (Ready to Retest, Open, Closed)
3. **Trend Line** - Line chart over time (Critical + High findings)

---

## Configuration Updates

### Config Tab

Add new row:
```
Row N: VAPT Spreadsheet ID | [Spreadsheet ID input] | [Link to VAPT spreadsheet]
```

**Spreadsheet ID:** `17qeErP3VHxN7qcNQqhT6zGLukxZU4OKLmBMbsgsl1Rk`

### Update `buildConfig(ss)`

Add VAPT config row after existing module configs.

---

## Refresh Workflow

### `refreshDashboard()` Update

Add VAPT refresh step:

```javascript
function refreshDashboard() {
  // ... existing code ...

  // Refresh VAPT data
  refreshVAPTData();

  // ... existing code ...
}
```

### Daily Trigger

Existing daily trigger will automatically include VAPT:
- `refreshDashboard()` → `refreshVAPTData()` → `appendVAPTHistory_()`

---

## Implementation Phases

### Phase 1: Dashboard Backend (Priority 1)
- [ ] Create `buildVAPT(ss)` function
- [ ] Create `buildVAPTHistory(ss)` function
- [ ] Create `refreshVAPTData()` function
- [ ] Create helper functions:
  - [ ] `getVAPTSpreadsheetId_(config)`
  - [ ] `fetchAdHocVAPTData_(vaptSs)`
  - [ ] `fetchRegularVAPTData_(vaptSs)`
  - [ ] `processVAPTData_(adHocData, regularData)`
  - [ ] `calculateVAPTSummary_(processedData)`
  - [ ] `appendVAPTHistory_(ss, vaptData, timestamp)`
- [ ] Update `buildConfig(ss)` to include VAPT config row
- [ ] Update `createDashboard()` to include VAPT tabs
- [ ] Update `refreshDashboard()` to include VAPT refresh

### Phase 2: Web App Backend (Priority 2)
- [ ] Create `getVAPTData_(ss)` function
- [ ] Update `getDashboardData()` to include VAPT data
- [ ] Test API response with VAPT data

### Phase 3: Web App Frontend (Priority 3)
- [ ] Add VAPT tab to navigation
- [ ] Create VAPT summary cards
- [ ] Create VAPT charts (Chart.js):
  - [ ] Severity distribution pie chart
  - [ ] Status breakdown stacked bar chart
  - [ ] Trend line chart
- [ ] Create VAPT data table
- [ ] Add filtering/sorting to table

### Phase 4: Testing & Deployment (Priority 4)
- [ ] Test with sample data
- [ ] Test daily refresh trigger
- [ ] Test Web App VAPT tab
- [ ] Deploy to Testing Script
- [ ] Verify in Testing environment
- [ ] Deploy to Production Script
- [ ] Update documentation

---

## Data Flow Diagram

```
VAPT Spreadsheet (Source)
  ├─ Ad Hoc VAPT (C1:U30)
  └─ Regular VAPT (C1:AB30)
         ↓
    fetchAdHocVAPTData_()
    fetchRegularVAPTData_()
         ↓
    processVAPTData_()
         ↓
    Combined VAPT Data
         ↓
    ├─→ buildVAPT() → VAPT Tab (Dashboard)
    ├─→ appendVAPTHistory_() → VAPT History Tab
    └─→ getVAPTData_() → Web App API
              ↓
         Web App Frontend
           └─→ VAPT Tab with Charts
```

---

## Summary Metrics to Calculate

### Overall:
- Total Applications (Ad Hoc + Regular)
- Total Findings (all severities, all statuses)
- Total Open Findings (Ready to Retest + Open)
- Total Closed Findings
- Production Count

### By Severity:
- Critical: Total, Ready to Retest, Open, Closed
- High: Total, Ready to Retest, Open, Closed
- Medium: Total, Ready to Retest, Open, Closed
- Low: Total, Ready to Retest, Open, Closed
- Info: Total, Ready to Retest, Open, Closed

### By VAPT Status:
- Done
- In Progress
- Not Started
- Todo

### By Type:
- Ad Hoc: count, findings
- Regular: count, findings

### KPIs for Web App:
1. **Open Critical/High Findings** (Red alert if > threshold)
2. **Ready to Retest Backlog** (Findings waiting for retest)
3. **Closure Rate** (Closed / Total over time)
4. **New Findings per Week** (Trend from history)

---

## Timeline Estimate

**Phase 1 (Dashboard Backend):** 4-6 hours
- Complex data processing from 2 sources
- Summary calculations
- History tracking logic

**Phase 2 (Web App Backend):** 1-2 hours
- Relatively straightforward API addition

**Phase 3 (Web App Frontend):** 2-3 hours
- HTML/CSS/JS for VAPT tab
- Chart.js integration
- Table with filtering

**Phase 4 (Testing & Deployment):** 1-2 hours
- Testing with real data
- Deployment to both environments
- Documentation

**Total Estimate:** 8-13 hours

---

## Notes & Considerations

1. **Data Quality:**
   - Handle empty rows
   - Handle missing values (null/undefined)
   - Validate numeric values (findings counts)

2. **Performance:**
   - Batch read from VAPT spreadsheet (read entire range at once)
   - Cache processed data during refresh
   - Minimize API calls in Web App

3. **Error Handling:**
   - VAPT spreadsheet not accessible
   - Invalid spreadsheet ID in Config
   - Missing tabs (Ad Hoc VAPT, Regular VAPT)
   - Data format changes

4. **Future Enhancements:**
   - Filter by PIC VAPT
   - Filter by application
   - Export VAPT report
   - Email alerts for critical findings
   - Integration with Jira (if VAPT findings tracked there)

---

**Created:** March 31, 2026
**Author:** Samuel Pangestu - QA INA Digital
**Status:** Planning - Ready for Implementation
