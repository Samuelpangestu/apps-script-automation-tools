# 📋 QA Test Management Template

**Comprehensive test case management system for Web, Mobile, API & Performance testing**

> Production-ready Google Sheets template for managing test cases, tracking execution, and monitoring quality metrics per module/project.

---

## 📋 **Overview**

QA Test Management Template is a standardized Google Sheets-based system for comprehensive test case management. Built with Google Apps Script automation, it provides QA teams with a complete workflow from test design to execution tracking and reporting.

**Key Capabilities:**
- ✨ One-click setup - Generate complete workspace instantly
- 📋 Separate repositories for Web/Mobile and API test cases
- 🎯 Execution tracking with real-time status updates
- 📊 Auto-generated charts (Pass Rate trends, Status distribution)
- ⚡ Performance testing integration (K6/JMeter)
- 🔐 RBAC testing support
- 📈 Coverage metrics calculation per module

---

## ✨ **Features**

### 📝 **Test Case Repository**

**Web & Mobile Testing:**
- TC_ID, Module/Submodule hierarchy
- Test Scenario & Steps
- Expected Result & Test Data
- Priority (P0-P4)
- Test Type (Functional, Regression, Smoke, etc.)
- RBAC roles
- Automation status

**API Testing:**
- Endpoint & HTTP Method
- Request/Response specifications
- Authentication & Authorization
- Performance SLA
- Expected status codes

### 🎯 **Execution Tracking**

**Eksekusi Sheet:**
- Link to test cases (auto-populated)
- Execution date & tester PIC
- Status (Pass/Fail/Blocked/N/A)
- Notes & Screenshots
- Actual vs Expected results
- Bug tracking reference

### 📊 **Dashboard & Reporting**

**Auto-Generated Insights:**
- Pass rate percentage
- Test coverage by module
- Status distribution (Pass/Fail/Blocked)
- Priority coverage (P0-P4 breakdown)
- Execution trends over time
- Visual charts

### ⚡ **Performance Testing**

**Performance Sheet:**
- Load testing scenarios
- K6/JMeter integration
- SLA thresholds
- Response time metrics
- Throughput measurements
- Error rate tracking

### 📚 **Documentation**

**Appendix Section:**
- Test hierarchy explanation
- Sheet structure guide
- Status definitions
- Priority levels
- RBAC roles
- Automation status codes
- Performance metrics glossary
- HTTP methods reference

---

## 🛠️ **Tech Stack**

- **Google Apps Script** (V8 Runtime)
- **SpreadsheetApp API** - Sheet manipulation
- **Charts API** - Data visualization
- **Custom Formulas** - Dynamic calculations
- **Conditional Formatting** - Visual status indicators
- **Data Validation** - Input constraints

---

## 🚀 **Setup Guide**

### **Step 1: Initial Setup**

1. Open Google Sheets
2. Go to **Extensions → Apps Script**
3. Copy code from `src/MasterQATCM.js`
4. Save project

### **Step 2: Generate Template**

Run from custom menu:
```
QA Template → Setup Awal (Jalankan Pertama Kali)
```

This creates:
- ✅ Sheet "WEB_MOBILE_Master" - Web/Mobile test cases
- ✅ Sheet "API_Master" - API test cases
- ✅ Sheet "Eksekusi" - Execution tracking
- ✅ Sheet "Performance" - Performance testing
- ✅ Sheet "Dashboard" - KPI visualization
- ✅ Sheet "Appendix" - Documentation

### **Step 3: Configuration**

**Customize settings in code:**
```javascript
var PROJECT_NAME = "Your Project Name";
var MODULES = ["Module1", "Module2", "Module3"];
```

**Set up validation:**
- Priority dropdown (P0-P4)
- Status dropdown (Pass/Fail/Blocked/N/A)
- Test Type options

---

## 💻 **Usage**

### **Custom Menu Actions**

Access from: **QA Template** menu

1. **Setup Awal** - Initial template generation
2. **Update Dashboard** - Refresh charts & metrics
3. **Export Report** - Generate PDF report
4. **Import Test Cases** - Bulk upload from CSV

### **Adding Test Cases**

**Web/Mobile Test:**
1. Go to "WEB_MOBILE_Master" sheet
2. Fill in columns:
   - TC_ID (e.g., WEB-001)
   - Module/Submodule
   - Test Scenario
   - Test Steps
   - Expected Result
   - Priority
3. Auto-link to Eksekusi sheet

**API Test:**
1. Go to "API_Master" sheet
2. Fill in:
   - TC_ID (e.g., API-001)
   - Endpoint URL
   - HTTP Method
   - Request Body/Headers
   - Expected Response
   - SLA requirements

### **Tracking Execution**

1. Go to "Eksekusi" sheet
2. Select test case (auto-populated from Master)
3. Execute test
4. Update status (Pass/Fail/Blocked)
5. Add notes & screenshots
6. Dashboard auto-updates

---

## 📊 **Dashboard Metrics**

### **Key Performance Indicators**

- **Pass Rate** - Percentage of passed tests
- **Coverage** - Test cases per module
- **Execution Rate** - Completed vs Total
- **Priority Distribution** - P0-P4 breakdown
- **Trend Analysis** - Pass rate over time

### **Visual Reports**

- **Pass Rate Chart** - Line graph over time
- **Status Pie Chart** - Pass/Fail/Blocked distribution
- **Coverage Bar Chart** - Tests per module
- **Priority Heatmap** - Critical test coverage

---

## 📋 **Sheet Structure**

### **WEB_MOBILE_Master**

| Column | Description |
|--------|-------------|
| TC_ID | Unique test case identifier |
| Module | High-level module name |
| Submodule | Specific feature/submodule |
| Test Scenario | What is being tested |
| Test Steps | Step-by-step instructions |
| Expected Result | Expected outcome |
| Test Data | Sample data for testing |
| Priority | P0 (Critical) to P4 (Low) |
| Test Type | Functional, Regression, etc. |
| RBAC | Required user role |
| Automation | Manual/Automated/Automatable |

### **API_Master**

| Column | Description |
|--------|-------------|
| TC_ID | API test identifier |
| Endpoint | API endpoint URL |
| HTTP Method | GET/POST/PUT/DELETE |
| Request Headers | Required headers |
| Request Body | JSON/XML payload |
| Expected Response | Response structure |
| Expected Status | HTTP status code |
| SLA | Performance requirement |
| Auth Required | Authentication type |

### **Eksekusi**

| Column | Description |
|--------|-------------|
| Date | Execution date |
| TC_ID | Test case reference |
| Tester | PIC name |
| Status | Pass/Fail/Blocked/N/A |
| Notes | Execution notes |
| Screenshot | Evidence URL |
| Bug ID | Linked bug ticket |

---

## 🔧 **Customization**

### **Add Custom Columns**

Edit in `MasterQATCM.js`:
```javascript
var CUSTOM_COLUMNS = [
  "Environment",
  "Test Iteration",
  "Browser Version"
];
```

### **Modify Status Options**

```javascript
var STATUS_OPTIONS = [
  "Pass",
  "Fail",
  "Blocked",
  "In Progress",
  "Skipped"
];
```

### **Configure Auto-Refresh**

```javascript
function setupTrigger() {
  ScriptApp.newTrigger('updateDashboard')
    .timeBased()
    .everyHours(1)
    .create();
}
```

---

## 📈 **Best Practices**

### **Test Case Design**
- Use clear, descriptive TC_IDs
- Follow naming convention: `[TYPE]-[MODULE]-[NUMBER]`
- Keep test steps atomic and repeatable
- Include all necessary test data

### **Execution Tracking**
- Update status immediately after testing
- Add detailed notes for failures
- Attach screenshots for visual evidence
- Link related bug tickets

### **Maintenance**
- Regular cleanup of obsolete test cases
- Archive completed test cycles
- Update dashboard weekly
- Review and optimize formulas

---

## 🐛 **Troubleshooting**

### **Setup Fails**

1. Check Apps Script execution permissions
2. Verify sheet quota not exceeded
3. Review execution transcript
4. Clear browser cache

### **Charts Not Updating**

1. Run "Update Dashboard" from menu
2. Verify data ranges are correct
3. Check formula references
4. Refresh browser

### **Performance Issues**

1. Reduce number of test cases per sheet
2. Archive old execution data
3. Optimize formulas (use ARRAYFORMULA)
4. Limit conditional formatting rules

---

## 📈 **Impact Metrics**

- **12+ projects** using this template
- **500+ test cases** managed
- **95%+ pass rate** tracked
- **80% time saved** on setup vs manual creation

---

## 🎯 **Use Cases**

- **Web Application Testing** - UI/UX validation
- **Mobile App Testing** - iOS/Android testing
- **API Testing** - Backend integration testing
- **Performance Testing** - Load/stress testing
- **Regression Testing** - Release validation
- **Compliance Testing** - RBAC & security testing

---

## 🔗 **Related Projects**

- [QA Dashboard](../qa-dashboard/) - Multi-project aggregation
- [MOM Rolling & PIC](../mom-rolling-pic/) - Team automation

---

## 📞 **Support**

For issues or questions:
- Check [Troubleshooting](#troubleshooting) section
- Review Apps Script execution logs
- Contact: departemen.qa@inadigital.co.id

---

<div align="center">

**Part of [Apps Script Automation Tools](../../README.md)**

Standardized test management for quality assurance teams

</div>
