---
description: Debug and fix QA Dashboard issues
---

Help debug and fix issues with the QA Dashboard:

1. Ask user to describe the issue or error they're seeing
2. Read relevant source files from qa-dashboard/src/
3. Check for common issues:
   - STATUS column reading (col Z issues)
   - KPI calculation from Summary sheet
   - Module data pulling
   - Broadcast functions
   - Trigger setup
4. Propose fixes with explanation
5. Implement fixes in qa-dashboard/src/ files
6. Test by deploying with clasp
7. Commit and push changes

Reference files:
- qa-dashboard/src/MasterDashboard.js - Main logic
- qa-dashboard/src/BroadcastFix.js - Broadcast utilities
- qa-dashboard/qa-dashboard.md - Documentation

Always preserve existing functionality and add comments for changes.
