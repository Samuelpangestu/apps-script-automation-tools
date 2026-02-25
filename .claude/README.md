# Claude Code Commands - QA Test Management

Custom commands untuk memudahkan development QA Test Management projects.

## 📋 Available Commands

Ketik command ini di Claude Code untuk menjalankan workflow otomatis:

### 🚀 Deployment

- **`/deploy-dashboard`** - Deploy QA Dashboard ke Apps Script
  - Check status & uncommitted changes
  - Push menggunakan clasp
  - Verify deployment berhasil

- **`/deploy-template`** - Deploy QA Test Management template ke Apps Script
  - Check status & uncommitted changes
  - Push menggunakan clasp
  - Verify deployment berhasil

### 🔄 Sync & Update

- **`/sync-clasp`** - Pull latest changes dari Apps Script projects
  - Sync Dashboard project
  - Sync Template project
  - Show diff dan ask untuk commit

### 🔧 Development

- **`/fix-dashboard`** - Debug dan fix Dashboard issues
  - Analyze error atau issue
  - Propose fixes dengan explanation
  - Implement dan test changes
  - Commit dan deploy

- **`/update-template`** - Update template dengan features baru
  - Add new functionality
  - Maintain backward compatibility
  - Update documentation
  - Deploy changes

### 📚 Help

- **`/qa-help`** - Show comprehensive help
  - List semua commands
  - Project structure
  - Quick workflows
  - Documentation links
  - Tips & tricks

## 🎯 Usage Examples

```bash
# Deploy dashboard after making changes
/deploy-dashboard

# Pull latest from Apps Script before starting work
/sync-clasp

# Get help when stuck
/qa-help

# Fix a specific issue
/fix-dashboard
# Then describe the issue in natural language
```

## 💡 Tips

1. **Start with `/qa-help`** untuk overview lengkap
2. Gunakan **`/sync-clasp`** sebelum mulai development
3. Gunakan **`/deploy-*`** untuk quick deployment tanpa manual clasp commands
4. Gunakan **`/fix-*`** untuk debugging dengan AI assistance

## 🔗 Related Files

- `CLASP_WORKFLOW.md` - Detailed clasp workflow guide
- `setup-clasp.sh` - Quick setup script
- `.github/workflows/deploy-apps-script.yml` - CI/CD pipeline

## 📝 Notes

Commands ini adalah wrapper untuk common tasks yang:
- Reduce manual typing
- Ensure best practices
- Automate repetitive workflows
- Provide contextual help

Setiap command sudah include project-specific context (Script IDs, folder structure, etc.).
