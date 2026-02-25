---
description: Update QA Test Management template with new features
---

Update the QA Test Management template:

1. Ask user what feature or improvement they want to add
2. Read current template from qa-test-management/src/QATestManagement.js
3. Review the structure:
   - Sheet creation functions (createQASheet)
   - TC_Master, TC_Execution sheets
   - API_Master, API_Execution sheets
   - Summary sheet with KPIs
   - Dashboard and visualization
4. Implement the requested changes
5. Test changes if possible
6. Update documentation if needed
7. Deploy to Apps Script
8. Commit changes with descriptive message

Key considerations:
- Maintain backward compatibility
- Follow existing code style (PERURI branding, color schemes)
- Add proper error handling
- Document new features in comments
- Update version/copyright if major changes

Reference: qa-test-management/qa-test-management.md for template structure.
