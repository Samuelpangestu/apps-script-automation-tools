# 🛠️ Apps Script Automation Tools

**Production-ready automation tools built with Google Apps Script and Google Sheets**

> A collection of enterprise-grade automation solutions for test management, team collaboration, and workflow optimization - all powered by Google Workspace integration.

---

## 🎯 **Overview**

This repository contains a suite of automation tools built with **Google Apps Script** and **Google Sheets** as the core dependencies. Each tool is designed to solve real-world workflow challenges through intelligent automation.

**Core Technologies:**
- 🔧 **Google Apps Script** - Server-side JavaScript automation platform
- 📊 **Google Sheets** - Data storage, UI, and computational engine
- 🔄 **Apps Script APIs** - SpreadsheetApp, DocumentApp, DriveApp, MailApp, UrlFetchApp
- 🚀 **clasp** - Modern development workflow with version control
- ⚙️ **GitHub Actions** - Automated CI/CD deployment pipeline

**Capabilities Demonstrated:**
- ✅ **Automated Workflows** - Reduce manual tasks with smart automation
- ✅ **Data Integration** - Aggregate and process data from multiple sources
- ✅ **Real-time Dashboards** - Live KPIs and metrics visualization
- ✅ **Notification Systems** - Email, Google Chat, and webhook integrations
- ✅ **Document Generation** - Automated template creation and management
- ✅ **Scheduling & Triggers** - Time-based and event-driven automation

---

## 🚀 **Projects**

### 1. **QA Test Management System**
> Comprehensive test case management and execution tracking for Web, Mobile & API testing

**📁 Location:** [`projects/qa-test-management/`](projects/qa-test-management/)

**Features:**
- ✨ **Instant Setup** - Generate complete test management workspace with one click
- 📋 **Test Case Repository** - Master registry for Web/Mobile and API test cases
- 🎯 **Execution Tracking** - Track test runs with real-time status updates
- 📊 **Smart Dashboard** - Auto-generated charts (Pass Rate trends, Status distribution)
- ⚡ **Performance Testing** - K6/JMeter integration with SLA validation
- 🔐 **RBAC Testing** - Built-in support for access control verification
- 📈 **Coverage Metrics** - Automatic calculation per module/feature

**Tech Stack:**
- Google Apps Script (V8 Runtime)
- SpreadsheetApp API
- Charts API
- Custom formulas & conditional formatting

**Use Cases:**
- QA teams managing 100+ test cases
- Multi-module testing (SIPGN, INAGOV, e-Meterai projects)
- Compliance testing with audit trails

---

### 2. **QA Portfolio Dashboard**
> Centralized monitoring dashboard aggregating test results from multiple projects

**📁 Location:** [`projects/qa-dashboard/`](projects/qa-dashboard/)

**Features:**
- 🎛️ **Multi-Project Aggregation** - Pull data from 10+ test management sheets
- 📊 **Consolidated KPIs** - Web, Mobile, API, Performance metrics in one view
- 🔄 **Auto-Refresh** - Scheduled triggers for real-time data sync
- 📢 **Broadcast Updates** - Push notes/templates to all projects at once
- 🎨 **Visual Reports** - Executive dashboard with charts & statistics
- 👥 **Team Overview** - PIC tracking, project health status

**Tech Stack:**
- Google Apps Script
- SpreadsheetApp (Advanced Range Operations)
- Time-based Triggers
- Custom menu actions

**Use Cases:**
- QA Managers monitoring multiple projects
- Executive reporting
- Team capacity planning

---

### 3. **MOM Rolling & PIC Reminder**
> Automated rotation system for meeting note-takers with integrated reminders

**📁 Location:** [`projects/mom-rolling-pic/`](projects/mom-rolling-pic/)

**Features:**
- 🎲 **Smart Rotation** - Weighted random assignment based on history
- 📅 **Schedule Management** - Configurable days (Mon, Wed, Fri)
- 📝 **Auto Doc Creation** - Generate MOM template in Google Drive
- 📧 **Email Notifications** - Automated reminders to assigned PIC
- 💬 **Google Chat Integration** - Team notifications via webhook
- 📊 **History Tracking** - Assignment audit trail

**Tech Stack:**
- Google Apps Script
- DocumentApp (Template generation)
- DriveApp (Folder management)
- MailApp & UrlFetchApp (Notifications)
- Time-based Triggers

**Use Cases:**
- Daily standup meeting management
- Fair PIC rotation across team members
- Automatic meeting documentation

---

## 🛠️ **Technical Highlights**

### **Modern Development Workflow**

```
Local Development (VS Code/WebStorm)
         ↓ clasp push
   Google Apps Script (Cloud)
         ↓ clasp pull
   Git Repository (Version Control)
         ↓ GitHub Actions
   Auto-Deploy (CI/CD)
```

**Tools & Setup:**
- **clasp** - Command-line Apps Script management
- **Git** - Version control with GitHub
- **GitHub Actions** - Automated deployment pipeline
- **ESLint** - Code quality enforcement (optional)
- **Custom scripts** - Sync automation

### **Code Quality**

- ✅ Modular architecture with separation of concerns
- ✅ Comprehensive error handling with try-catch
- ✅ Logging for debugging and audit trails
- ✅ Idempotent operations (safe to re-run)
- ✅ Input validation and sanitization
- ✅ Performance optimized (batch operations, minimal API calls)

### **Documentation**

- ✅ Inline JSDoc comments
- ✅ README for each project
- ✅ Setup guides and troubleshooting
- ✅ API reference documentation

---

## 📦 **Repository Structure**

```
google-apps-script-portfolio/
├── projects/
│   ├── qa-test-management/      # Test case management template
│   │   ├── src/
│   │   │   ├── QATestManagement.js
│   │   │   └── appsscript.json
│   │   ├── .clasp.json
│   │   └── README.md
│   │
│   ├── qa-dashboard/            # Portfolio dashboard
│   │   ├── src/
│   │   │   ├── Dashboard.js
│   │   │   ├── BroadcastFix.js
│   │   │   ├── BroadcastAllFixes.js
│   │   │   └── appsscript.json
│   │   ├── .clasp.json
│   │   └── README.md
│   │
│   └── mom-rolling-pic/         # MOM rotation & reminders
│       ├── src/
│       │   ├── Code.js
│       │   └── appsscript.json
│       ├── .clasp.json
│       └── README.md
│
├── .github/
│   └── workflows/
│       └── deploy-apps-script.yml
│
├── docs/                        # Portfolio documentation
│   ├── ARCHITECTURE.md
│   ├── DEPLOYMENT.md
│   └── screenshots/
│
├── scripts/
│   ├── setup-clasp.sh
│   └── sync-from-appscript.sh
│
├── CLASP_WORKFLOW.md
└── README.md                    # This file
```

---

## 🚀 **Getting Started**

### **Prerequisites**

```bash
# Install Node.js (v14+)
node --version

# Install clasp globally
npm install -g @google/clasp

# Login to Google account
clasp login
```

### **Enable Apps Script API**

1. Visit: https://script.google.com/home/usersettings
2. Toggle **ON** "Google Apps Script API"

### **Clone & Setup**

```bash
# Clone repository
git clone https://github.com/YOUR_USERNAME/google-apps-script-portfolio.git
cd google-apps-script-portfolio

# Quick setup
./scripts/setup-clasp.sh
```

### **Deploy a Project**

```bash
# Deploy QA Dashboard
cd projects/qa-dashboard
clasp push

# Or use Claude Code commands
/deploy-dashboard
```

---

## 🎓 **Skills Demonstrated**

### **Programming & Scripting**
- ✅ JavaScript (ES6+)
- ✅ Google Apps Script APIs
- ✅ Bash scripting
- ✅ Git version control

### **Software Engineering**
- ✅ Modular architecture design
- ✅ Error handling & logging
- ✅ Performance optimization
- ✅ Code documentation
- ✅ Testing strategies

### **DevOps & Automation**
- ✅ CI/CD pipelines (GitHub Actions)
- ✅ Command-line tools (clasp)
- ✅ Deployment automation
- ✅ Environment management

### **Google Workspace**
- ✅ Google Sheets (Advanced formulas, Apps Script)
- ✅ Google Docs (Template generation)
- ✅ Google Drive (File management)
- ✅ Gmail (Automated emails)
- ✅ Google Chat (Webhooks)

### **QA & Testing**
- ✅ Test case design
- ✅ Test execution tracking
- ✅ Performance testing (K6/JMeter)
- ✅ API testing
- ✅ RBAC testing
- ✅ Metrics & reporting

---

## 📈 **Impact & Results**

### **QA Test Management System**
- 📊 **12+ projects** using the template
- ⚡ **500+ test cases** managed
- 🎯 **95%+ pass rate** tracked
- ⏱️ **80% time saved** on setup vs manual creation

### **QA Portfolio Dashboard**
- 🎛️ **10+ modules** aggregated in real-time
- 📊 **Daily reports** with auto-refresh
- 👥 **5+ team members** using centralized view
- 📉 **50% reduction** in manual status reporting

### **MOM Rolling System**
- 🎲 **Fair rotation** across 8+ team members
- 📝 **100% documentation** compliance
- ⏰ **Zero missed** assignments with auto-reminders
- 💬 **Instant notifications** via Google Chat

---

## 📚 **Documentation**

- **[Clasp Workflow Guide](CLASP_WORKFLOW.md)** - Complete development workflow
- **[Architecture Overview](docs/ARCHITECTURE.md)** - System design & patterns
- **[Deployment Guide](docs/DEPLOYMENT.md)** - Production deployment steps
- **[API Reference](docs/API_REFERENCE.md)** - Function documentation

---

## 🔗 **Live Demo**

> Note: Live demos require Google Workspace access. Contact for demo credentials.

- **QA Dashboard** - [View Dashboard](https://docs.google.com/spreadsheets/d/DASHBOARD_ID)
- **Test Management** - [View Template](https://docs.google.com/spreadsheets/d/TEMPLATE_ID)

---

## 🤝 **Contributing**

This is a portfolio project, but suggestions and feedback are welcome!

```bash
# Fork the repository
# Create feature branch
git checkout -b feature/your-feature

# Commit changes
git commit -m "Add your feature"

# Push and create PR
git push origin feature/your-feature
```

---

## 📄 **License**

**Proprietary** - QA INA Digital Portfolio Projects

Individual projects are proprietary to their respective teams:
- QA Test Management: © QA INA Digital
- MOM Rolling: © INA Digital Team

This repository is for **portfolio showcase purposes** only.

---

## 👤 **Author**

**Samuel Pangestu**
- 💼 QA Engineer @ INA Digital (Peruri)
- 📧 Email: [your.email@example.com]
- 🔗 LinkedIn: [Your LinkedIn]
- 🌐 Portfolio: [Your Portfolio Site]

---

## 🎯 **For Recruiters**

This portfolio demonstrates:

✅ **Full-stack automation development** - End-to-end solutions from requirement to deployment

✅ **Enterprise-grade code quality** - Production-ready with error handling, logging, documentation

✅ **Modern DevOps practices** - CI/CD, version control, automated testing

✅ **Cross-functional expertise** - QA, automation, scripting, API integration

✅ **Real-world impact** - Measurable results in team productivity and quality

✅ **Self-driven innovation** - Identified problems, designed solutions, implemented independently

---

## 📞 **Contact**

Interested in discussing these projects or potential opportunities?

📧 **Email:** [your.email@example.com]
💼 **LinkedIn:** [Your LinkedIn URL]
📱 **Phone:** [Your Phone] (optional)

---

<div align="center">

**⭐ Star this repo if you find it useful!**

Made with ❤️ using Google Apps Script

[View Projects](#projects) • [Documentation](#documentation) • [Contact](#contact)

</div>
