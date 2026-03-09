# 📊 QA Portfolio Dashboard

**Multi-project test management aggregation dashboard with real-time KPIs**

> Centralized monitoring system aggregating test results from multiple QA projects with automated refresh and broadcast capabilities.

---

## 📋 **Overview**

QA Portfolio Dashboard is a Google Sheets-based aggregation system that consolidates test management data from multiple projects into a single, real-time dashboard. Built with Google Apps Script, it provides QA managers with comprehensive oversight of all active testing projects.

**Key Capabilities:**
- 🎛️ Aggregate data from 10+ test management sheets simultaneously
- 📊 Real-time KPIs for Web, Mobile, API, and Performance testing
- 🔄 Automated data refresh with time-based triggers
- 📢 Broadcast updates/templates to all connected projects
- 📈 Executive-level reporting and team capacity tracking

---

## ✨ **Features**

### 🎛️ **Multi-Project Aggregation**
- Pull test case counts, pass rates, and execution status from multiple sheets
- Support for 10+ active modules concurrently
- Configurable source sheet URLs
- Auto-detection of sheet structure

### 📊 **Real-Time Dashboard**
- Consolidated pass rate across all projects
- Test coverage metrics by module
- Execution status distribution
- Team PIC tracking
- Visual charts and statistics

### 🔄 **Auto-Refresh System**
- Time-based triggers (every 1 hour by default)
- Manual refresh option via custom menu
- Incremental updates to minimize API calls
- Error handling and retry logic

### 📢 **Broadcast Functionality**
- Push template updates to all projects at once
- Update Notes & References sections
- Add new columns (e.g., QA Team Lead)
- Recreate Appendix sections with consistent formatting

### 📈 **Advanced Features**
- **BroadcastFix** - Fix common issues across all sheets
- **BroadcastAllFixes** - Run comprehensive fix suite
- **AppendixFullSetup** - Standardize documentation sections
- **Custom Menu** - Quick access to all functions

---

## 🛠️ **Tech Stack**

- **Google Apps Script** (V8 Runtime)
- **SpreadsheetApp API** - Advanced range operations
- **UrlFetchApp** - HTTP requests for cross-sheet data
- **Time-based Triggers** - Automated scheduling
- **Custom Menu System** - User-friendly interface

---

## 📁 **Files**

### **Core Scripts**

- **`MasterDashboard.js`** - Main aggregation and dashboard logic
- **`BroadcastFix.js`** - Broadcast fixes to all sheets
- **`BroadcastAllFixes.js`** - Master runner for all fixes
- **`BroadcastFixesLead.js`** - Add QA Team Lead column
- **`BroadcastFixNoteAndTemplate.js`** - Update Notes section
- **`BroadcastRecreateAppendix.js`** - Recreate Appendix

### **Appendix Management**

- **`AppendixContentFix.js`** - Targeted content updates
- **`AppendixFullSetup.js`** - Complete Appendix rebuild
- **`AppendixRevert.js`** - Rollback to previous version

---

## 🚀 **Setup**

### **Prerequisites**

- Google Sheets with test management data
- Apps Script enabled
- clasp installed (for development)

### **Development & Deployment**

#### **Environments**

This project has two environments:

- **Testing**: Development and testing (default)
  - Script ID: `1LJ83OATTAp7ChDWGkrSTg0b9KmMhOABISBrAJrB54JksjQ7mi5oNB7C3`

- **Production**: Live dashboard used by QA team
  - Script ID: `1lHO8yKyqKs1_n5GV1m-SJMACLS95Jc7yy6dM_ItyT-l_-GdmkGQk3OIO`

See [../../ENVIRONMENTS.md](../../ENVIRONMENTS.md) for full details.

#### **Deployment Commands**

```bash
# Deploy to testing only
./deploy-testing.sh

# Deploy to production only (will ask for confirmation)
./deploy-production.sh

# Deploy to both testing and production
./deploy-all.sh
```

### **Configuration**

1. **Create Dashboard Sheet**
   - Create new Google Sheet
   - Open Extensions → Apps Script
   - Deploy code from `src/`

2. **Configure Source Sheets**
   - Add source sheet URLs in Config
   - Ensure sheets have consistent structure
   - Verify column mappings

3. **Set Up Triggers**
   - Run from menu: Dashboard → Setup Auto-Refresh
   - Default: Every 10 minutes
   - Customizable in Config tab (Q4-R4)

---

## 💻 **Usage**

### **Dashboard Menu**

Access from: **Dashboard** (custom menu)

- **Refresh All Data** - Manual data refresh
- **Broadcast Fix** - Send fixes to all projects
- **Run All Fixes** - Comprehensive fix suite
- **Setup Auto-Refresh** - Configure triggers

### **Common Operations**

**Refresh Data:**
```
Dashboard → Refresh All Data
```

**Broadcast Update:**
```
Dashboard → Broadcast Fix → Select Fix Type
```

**Add New Project:**
1. Add sheet URL to Config
2. Run Refresh All Data
3. Verify data in dashboard

---

## 📊 **Dashboard Sections**

### **1. Summary Dashboard**
- Total projects tracked
- Overall pass rate
- Total test cases
- Execution status breakdown

### **2. Project List**
- Project name
- Module/submodule
- PIC (Person In Charge)
- QA Team Lead
- Last updated timestamp

### **3. Metrics View**
- Pass rate trends
- Coverage by module
- Priority distribution
- Test type breakdown (Web/Mobile/API/Performance)

---

## 🔧 **Configuration**

### **Source Sheet Format**

Expected columns in source sheets:
- Test Case ID
- Module/Submodule
- Test Scenario
- Status (Pass/Fail/Blocked)
- Priority
- PIC
- Last Execution Date

### **Broadcast Settings**

Customize in code:
```javascript
var CONFIG = {
  REFRESH_INTERVAL: 60, // minutes
  MAX_RETRIES: 3,
  TIMEOUT: 30000 // milliseconds
};
```

---

## 🐛 **Troubleshooting**

### **Data Not Refreshing**

1. Check trigger status (Extensions → Apps Script → Triggers)
2. Verify source sheet URLs are correct
3. Check execution logs for errors
4. Ensure sheets have correct permissions

### **Broadcast Fails**

1. Verify target sheets exist and are accessible
2. Check for structural differences in sheets
3. Review execution transcript
4. Run fixes individually to isolate issue

### **Performance Issues**

1. Reduce refresh frequency
2. Limit number of source sheets
3. Optimize data range selections
4. Use batch operations

---

## 📈 **Impact Metrics**

- **10+ modules** aggregated in real-time
- **Daily reports** with auto-refresh
- **5+ team members** using centralized view
- **50% reduction** in manual status reporting

---

## 🔗 **Related Projects**

- [QA Test Management](../qa-test-management/) - Individual project template
- [MOM Rolling & PIC](../mom-rolling-pic/) - Meeting automation

---

## 📞 **Support**

For issues or questions:
- Check [Troubleshooting](#troubleshooting) section
- Review execution logs in Apps Script Editor
- Contact: departemen.qa@inadigital.co.id

---

<div align="center">

**Part of [Apps Script Automation Tools](../../README.md)**

Built for scalable QA management across multiple projects

</div>
