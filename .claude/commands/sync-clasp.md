---
description: Sync local files with Apps Script projects (pull latest changes)
---

Sync both Apps Script projects with local files:

1. Pull latest changes from Dashboard Apps Script project
2. Pull latest changes from Test Management Apps Script project
3. Show diff if there are any changes
4. Ask user if they want to commit the pulled changes

This is useful when:
- Someone edited directly in Apps Script web editor
- You want to ensure local files are up to date
- Before starting new development work

Use:
- `cd qa-dashboard && clasp pull`
- `cd qa-test-management && clasp pull`

Show summary of what was pulled.
