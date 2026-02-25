# 🚀 Clasp Workflow - Apps Script Development Guide

## 📋 Overview

Project ini menggunakan **clasp** (Command Line Apps Script Projects) untuk integrasi development modern dengan Google Apps Script. Setup ini memungkinkan:

- ✅ Edit code di VS Code/WebStorm dengan autocomplete penuh
- ✅ Version control dengan Git
- ✅ Auto-deploy ke Apps Script saat push ke GitHub
- ✅ Multiple files dengan proper module management
- ✅ TypeScript support (opsional)

---

## 📁 Structure

```
qa-test-management-template/
├── qa-dashboard/
│   ├── src/
│   │   ├── Dashboard.js           # Main dashboard code
│   │   ├── BroadcastFix.js        # Broadcast fix utilities
│   │   └── appsscript.json        # Apps Script manifest
│   ├── .clasp.json                 # Clasp config (Dashboard)
│   ├── .claspignore                # Files to exclude
│   └── qa-dashboard.js             # Original (kept for reference)
│
├── qa-test-management/
│   ├── src/
│   │   ├── QATestManagement.js    # Template code
│   │   └── appsscript.json        # Apps Script manifest
│   ├── .clasp.json                 # Clasp config (Template)
│   ├── .claspignore                # Files to exclude
│   └── qa-test-management.js      # Original (kept for reference)
│
└── .github/
    └── workflows/
        └── deploy-apps-script.yml  # Auto-deploy workflow
```

---

## 🔧 Setup Guide

### 1. Install clasp

```bash
npm install -g @google/clasp
```

### 2. Login to Google Account

```bash
clasp login
```

Browser akan terbuka untuk authorize clasp. Login dengan akun Google yang memiliki akses ke Apps Script projects.

### 3. Verify Configuration

```bash
# Check Dashboard config
cd qa-dashboard
clasp status

# Check Test Management config
cd qa-test-management
clasp status
```

---

## 💻 Development Workflow

### Option 1: Local Development + Manual Push

```bash
# 1. Edit files di src/ folder menggunakan VS Code/WebStorm
# 2. Test changes locally (jika perlu)
# 3. Push to Apps Script

cd qa-dashboard
clasp push

# Atau push dengan watch mode (auto-push on save)
clasp push --watch
```

### Option 2: Auto-Deploy via GitHub (Recommended!)

```bash
# 1. Edit files di src/ folder
# 2. Commit changes
git add qa-dashboard/src/
git commit -m "Update dashboard logic"

# 3. Push to GitHub
git push origin main

# 4. GitHub Actions will automatically deploy to Apps Script! 🎉
```

**Status check:**
- Go to: https://github.com/YOUR_USERNAME/qa-test-management-template/actions
- Monitor deployment progress

---

## 📝 Common Commands

### Push Changes

```bash
cd qa-dashboard
clasp push              # Push all changes
clasp push --force      # Force push (overwrite remote)
clasp push --watch      # Auto-push on file save
```

### Pull Changes

```bash
cd qa-dashboard
clasp pull              # Pull latest from Apps Script
```

Gunakan ini jika ada perubahan yang dibuat langsung di Apps Script web editor.

### Open in Browser

```bash
cd qa-dashboard
clasp open              # Open Apps Script editor
clasp open --webapp     # Open web app
```

### View Logs

```bash
clasp logs              # View execution logs
clasp logs --watch      # Watch logs in real-time
```

### Run Functions

```bash
clasp run functionName  # Run a specific function
```

---

## 🤖 GitHub Actions Setup

### Setup Secrets (PENTING!)

Untuk enable auto-deploy, tambahkan secret `CLASPRC_JSON` di GitHub:

1. **Get clasp credentials:**
   ```bash
   cat ~/.clasprc.json
   ```

2. **Add to GitHub Secrets:**
   - Go to: https://github.com/YOUR_USERNAME/qa-test-management-template/settings/secrets/actions
   - Click "New repository secret"
   - Name: `CLASPRC_JSON`
   - Value: Paste entire content dari `~/.clasprc.json`
   - Click "Add secret"

### Workflow Behavior

Auto-deploy akan trigger ketika:
- Push ke branch `main`
- Ada perubahan di `qa-dashboard/src/**` atau `qa-test-management/src/**`
- Manual trigger via GitHub Actions tab

---

## 🎯 Best Practices

### 1. Always Edit in `src/` Folder

```bash
✅ CORRECT: qa-dashboard/src/Dashboard.js
❌ WRONG:   qa-dashboard/qa-dashboard.js
```

File di luar `src/` folder adalah backup/reference saja.

### 2. Test Before Push

```bash
# Option A: Run function directly
clasp run myFunction

# Option B: Deploy as web app and test
clasp deploy --description "Testing changes"
```

### 3. Sync Regularly

```bash
# Pull perubahan dari Apps Script (jika ada edit manual)
clasp pull

# Push perubahan lokal
clasp push
```

### 4. Use Branches for Major Changes

```bash
# Create feature branch
git checkout -b feature/new-calculation

# Make changes in src/
# Test locally with: clasp push --watch

# When ready, commit and create PR
git add qa-dashboard/src/
git commit -m "Add new calculation logic"
git push origin feature/new-calculation

# Create PR in GitHub, auto-deploy happens after merge to main
```

---

## 🔄 Migration from Old Workflow

### Old Workflow (Manual)
1. Edit di Apps Script web editor
2. Copy-paste code ke local Git
3. Commit to GitHub
4. No auto-sync ❌

### New Workflow (Clasp)
1. Edit di VS Code/WebStorm ✨
2. Push to GitHub
3. Auto-deploy to Apps Script 🚀
4. Always in sync ✅

### Files Mapping

| Old File | New Location | Purpose |
|----------|--------------|---------|
| `qa-dashboard.js` | `qa-dashboard/src/Dashboard.js` | Active development |
| `broadcast-fix.js` | `qa-dashboard/src/BroadcastFix.js` | Active development |
| `qa-test-management.js` | `qa-test-management/src/QATestManagement.js` | Active development |

**Note:** Original files tetap ada untuk backward compatibility, tapi **jangan edit** file tersebut!

---

## 🐛 Troubleshooting

### Error: "No .clasp.json found"

```bash
# Make sure you're in the correct directory
cd qa-dashboard  # or qa-test-management
clasp status
```

### Error: "User has not enabled the Apps Script API"

1. Go to: https://script.google.com/home/usersettings
2. Enable "Google Apps Script API"

### Error: "Push failed - Authorization required"

```bash
# Re-login to clasp
clasp logout
clasp login
```

### Changes Not Appearing in Apps Script

```bash
# Force push
clasp push --force

# Or clear clasp cache
rm -rf ~/.clasp
clasp login
clasp push
```

### GitHub Actions Failing

1. Check secret is set: `CLASPRC_JSON`
2. Verify secret value is valid JSON
3. Check workflow logs: https://github.com/YOUR_USERNAME/qa-test-management-template/actions

---

## 📚 Resources

- [clasp Documentation](https://github.com/google/clasp)
- [Apps Script API Reference](https://developers.google.com/apps-script/api/reference/rest)
- [Apps Script Best Practices](https://developers.google.com/apps-script/guides/best-practices)

---

## ✨ Quick Reference

```bash
# Development
clasp push --watch        # Auto-push on save
clasp open               # Open in browser

# Deployment
git push origin main     # Auto-deploy via GitHub Actions

# Sync
clasp pull               # Pull from Apps Script
clasp push               # Push to Apps Script

# Debugging
clasp logs               # View logs
clasp run functionName   # Run function
```

---

## 📧 Support

Questions? Contact: departemen.qa@inadigital.co.id
