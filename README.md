# 🧪 QA Test Management Template

Template dan dashboard portfolio untuk manajemen testing QA di INA Digital.

Script otomasi untuk membangun ekosistem Test Management di Google Sheets secara instan, dengan integrasi modern development workflow menggunakan clasp.

---

## 📦 Components

### 1. QA Test Management Template
Template standardized untuk test case management:
- **TC_Master** - Test case repository Web/Mobile
- **TC_Execution** - Execution tracking & results
- **API_Master** - API test cases
- **API_Execution** - API test results
- **Summary** - KPI dashboard & metrics
- **Performance** - Performance test results (K6/JMeter)
- **BugReport** - Bug tracking

📄 Docs: [`qa-test-management/qa-test-management.md`](qa-test-management/qa-test-management.md)

### 2. QA Portfolio Dashboard
Dashboard agregasi untuk monitoring multiple projects:
- Pull data dari multiple QA Test Management sheets
- Consolidated metrics & KPIs
- Project health overview
- Broadcast updates to team
- Auto-refresh with triggers

📄 Docs: [`qa-dashboard/qa-dashboard.md`](qa-dashboard/qa-dashboard.md)

---

## 🚀 Getting Started

### Option A: Traditional (Web Editor)

1. Buka [Google Sheets](https://sheets.new)
2. **Extensions** > **Apps Script**
3. Copy kode dari `qa-test-management/src/QATestManagement.js`
4. Save dan run `createQASheet()`
5. Authorize permissions

### Option B: Modern Development (Clasp - Recommended!)

```bash
# Install clasp
npm install -g @google/clasp

# Login to Google
clasp login

# Enable Apps Script API (one-time setup)
# Visit: https://script.google.com/home/usersettings
# Toggle ON "Google Apps Script API"

# Setup project
./setup-clasp.sh

# Start development with auto-sync
cd qa-dashboard
clasp push --watch
```

📖 **Full guide**: [`CLASP_WORKFLOW.md`](CLASP_WORKFLOW.md)

---

## 💻 Development Workflow

### Local Development (Auto-sync)

```bash
# Start watch mode for auto-deploy on save
cd qa-dashboard
clasp push --watch

# Edit files in src/ folder
# Changes auto-push to Apps Script ✨
```

### Git-based Workflow (CI/CD)

```bash
# 1. Edit files in src/ folders using VS Code/WebStorm
# 2. Commit changes
git add qa-dashboard/src/
git commit -m "Update calculation logic"

# 3. Push to GitHub
git push origin main

# 4. GitHub Actions auto-deploys to Apps Script! 🚀
```

---

## 📁 Project Structure

```
qa-test-management-template/
├── qa-dashboard/                    # Portfolio Dashboard
│   ├── src/
│   │   ├── Dashboard.js             # Main dashboard code
│   │   ├── BroadcastFix.js          # Broadcast utilities
│   │   └── appsscript.json          # Apps Script manifest
│   ├── .clasp.json                  # Clasp config (scriptId)
│   └── qa-dashboard.md              # Documentation
│
├── qa-test-management/              # QA Template
│   ├── src/
│   │   ├── QATestManagement.js      # Template code
│   │   └── appsscript.json          # Apps Script manifest
│   ├── .clasp.json                  # Clasp config (scriptId)
│   └── qa-test-management.md        # Documentation
│
├── .github/workflows/
│   └── deploy-apps-script.yml       # Auto-deploy CI/CD pipeline
│
├── .claude/commands/                # Claude Code commands
│   ├── deploy-dashboard.md          # Deploy dashboard
│   ├── deploy-template.md           # Deploy template
│   ├── sync-clasp.md                # Sync with Apps Script
│   ├── fix-dashboard.md             # Debug & fix issues
│   ├── update-template.md           # Add features
│   └── qa-help.md                   # Show help
│
├── CLASP_WORKFLOW.md                # Detailed workflow guide
├── setup-clasp.sh                   # Quick setup script
└── README.md                        # This file
```

---

## 🤖 Claude Code Commands

Custom commands untuk memudahkan development:

```bash
/deploy-dashboard    # Deploy Dashboard to Apps Script
/deploy-template     # Deploy Template to Apps Script
/sync-clasp          # Pull latest from Apps Script
/fix-dashboard       # Debug and fix issues
/update-template     # Add new features
/qa-help             # Show comprehensive help
```

📖 Lihat [`.claude/README.md`](.claude/README.md) untuk detail commands.

---

## 📚 Documentation

- **[CLASP_WORKFLOW.md](CLASP_WORKFLOW.md)** - Complete clasp workflow guide
  - Setup instructions
  - Development workflow
  - Best practices
  - Troubleshooting
  - Command reference

- **[qa-dashboard/qa-dashboard.md](qa-dashboard/qa-dashboard.md)** - Dashboard documentation
  - Architecture overview
  - Data pulling mechanism
  - KPI calculations
  - Broadcast features

- **[qa-test-management/qa-test-management.md](qa-test-management/qa-test-management.md)** - Template documentation
  - Sheet structure
  - Formulas & calculations
  - Status tracking
  - RBAC testing

---

## 🔗 Apps Script Projects

- **Dashboard**: [Script Editor](https://script.google.com/home/projects/1lHO8yKyqKs1_n5GV1m-SJMACLS95Jc7yy6dM_ItyT-l_-GdmkGQk3OIO/edit)
- **Template**: [Script Editor](https://script.google.com/home/projects/11-ITJByolVjYDwVJ8og8-lC16G6g78pUD07cR20Dsvz6DQfpmhqo1_yw/edit)

---

## ⚙️ GitHub Actions Setup

### Enable Auto-Deploy

1. **Get clasp credentials:**
   ```bash
   clasp login
   cat ~/.clasprc.json
   ```

2. **Add GitHub Secret:**
   - Go to: `Settings → Secrets → Actions`
   - Click "New repository secret"
   - Name: `CLASPRC_JSON`
   - Value: Paste entire JSON from `~/.clasprc.json`

3. **Auto-deploy triggers on:**
   - Push to `main` branch
   - Changes in `qa-dashboard/src/` or `qa-test-management/src/`
   - Manual workflow dispatch

---

## 🛠️ Common Commands

```bash
# Push changes to Apps Script
cd qa-dashboard
clasp push

# Pull latest from Apps Script
clasp pull

# Watch mode (auto-push on save)
clasp push --watch

# Open in browser
clasp open

# View execution logs
clasp logs

# Run specific function
clasp run functionName
```

---

## 💡 Best Practices

1. ✅ **Edit in `src/` folders only** - Never edit root .js files
2. ✅ **Use watch mode** during development: `clasp push --watch`
3. ✅ **Pull before starting**: `clasp pull` to sync latest
4. ✅ **Commit frequently** with descriptive messages
5. ✅ **Use branches** for major changes
6. ✅ **Test before deploy** with `clasp run functionName`

---

## 🐛 Troubleshooting

### "No .clasp.json found"
```bash
cd qa-dashboard  # or qa-test-management
clasp status
```

### "User has not enabled the Apps Script API"
**Required before first clasp push/pull:**
1. Go to: https://script.google.com/home/usersettings
2. Toggle ON "Google Apps Script API"
3. Wait 1-2 minutes, then retry `clasp push`

### "Push failed - Authorization required"
```bash
clasp logout
clasp login
```

### Changes not appearing
```bash
clasp push --force
```

📖 See [CLASP_WORKFLOW.md](CLASP_WORKFLOW.md) for complete troubleshooting guide.

---

## ✨ Features

### QA Test Management Template
* ⚡ **Instant Setup** - 7 tabs standar QA dengan satu klik
* 🔄 **Auto-Sync** - Execution data tersinkron dengan Master
* 📊 **Smart Dashboard** - Pie chart & line chart real-time
* 🚀 **Performance Monitoring** - K6/JMeter results dengan SLA calculation
* 🔐 **RBAC Focus** - Kolom khusus untuk access control testing

### Portfolio Dashboard
* 📈 **Multi-project aggregation** - Monitor semua project dalam satu view
* 🎯 **KPI consolidation** - Web, API, Performance metrics
* 📢 **Broadcast features** - Send updates to team
* ⏰ **Auto-refresh** - Scheduled triggers untuk sync data
* 🎨 **Visual reports** - Charts dan statistics

---

## 👥 Team

**QA INA Digital**
- Email: departemen.qa@inadigital.co.id

---

## 📄 License

© QA INA Digital - Proprietary Template
Dilarang digunakan/disebarluaskan tanpa izin.

---

## 🔄 Version

- **Template**: v38+
- **Dashboard**: v1.0.0
- **Clasp Integration**: v1.0.0

---

**🚀 Quick Start**: Run `./setup-clasp.sh` dan baca [`CLASP_WORKFLOW.md`](CLASP_WORKFLOW.md)
**❓ Need Help?**: Type `/qa-help` in Claude Code
