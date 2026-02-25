---
description: Show help for QA Test Management project commands and workflow
---

Provide comprehensive help for the QA Test Management project:

## 📚 Available Commands

Show user all available Claude commands for this project:
- `/deploy-dashboard` - Deploy Dashboard to Apps Script
- `/deploy-template` - Deploy Test Management template
- `/sync-clasp` - Pull latest from Apps Script
- `/fix-dashboard` - Debug and fix Dashboard issues
- `/update-template` - Add features to template
- `/qa-help` - This help message

## 🏗️ Project Structure

```
qa-test-management-template/
├── qa-dashboard/          # Portfolio Dashboard
│   └── src/               # Active development files
│       ├── Dashboard.js
│       └── BroadcastFix.js
├── qa-test-management/    # QA Template
│   └── src/               # Active development files
│       └── QATestManagement.js
└── .github/workflows/     # Auto-deploy on push
```

## 🚀 Quick Workflows

**Local Development:**
```bash
cd qa-dashboard
clasp push --watch
```

**Git-based Deploy:**
```bash
git add qa-dashboard/src/
git commit -m "Update logic"
git push origin main  # Auto-deploys!
```

## 📖 Documentation

- `CLASP_WORKFLOW.md` - Full clasp workflow guide
- `qa-dashboard/qa-dashboard.md` - Dashboard documentation
- `qa-test-management/qa-test-management.md` - Template docs

## 🔗 Apps Script Links

- Dashboard: https://script.google.com/home/projects/1lHO8yKyqKs1_n5GV1m-SJMACLS95Jc7yy6dM_ItyT-l_-GdmkGQk3OIO/edit
- Template: https://script.google.com/home/projects/11-ITJByolVjYDwVJ8og8-lC16G6g78pUD07cR20Dsvz6DQfpmhqo1_yw/edit

## 💡 Tips

- Always edit files in `src/` folders, not root files
- Use `clasp open` to open project in browser
- Use `clasp logs` to view execution logs
- GitHub Actions auto-deploys on push to main

Ask me anything about the project!
