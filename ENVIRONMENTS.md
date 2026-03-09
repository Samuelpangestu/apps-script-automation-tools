# Environment Configuration

This repository contains multiple Apps Script projects with separate **Testing** and **Production** environments.

## QA Dashboard

**Location:** `projects/qa-dashboard/`

### Testing Environment
- **Script ID:** `1LJ83OATTAp7ChDWGkrSTg0b9KmMhOABISBrAJrB54JksjQ7mi5oNB7C3`
- **Spreadsheet:** [Link to Testing Dashboard]
- **Purpose:** Development and testing new features
- **Default:** `.clasp.json` points to this by default

### Production Environment
- **Script ID:** `1lHO8yKyqKs1_n5GV1m-SJMACLS95Jc7yy6dM_ItyT-l_-GdmkGQk3OIO`
- **Spreadsheet:** [Link to Production Dashboard]
- **Purpose:** Live production dashboard used by QA team
- **Usage:** Temporarily change `.clasp.json` scriptId to push to production

---

## QA Test Management Template

**Location:** `projects/qa-test-management/`

### Testing Environment
- **Script ID:** `[TESTING_SCRIPT_ID]`
- **Spreadsheet:** [Link to Testing Template]
- **Purpose:** Development and testing template changes

### Production Environment
- **Script ID:** `[PRODUCTION_SCRIPT_ID]`
- **Spreadsheet:** [Link to Production Template]
- **Purpose:** Live template used to create new QA modules

---

## Rolling MoM

**Location:** `projects/mom-rolling-pic/`

### Testing Environment
- **Script ID:** `[TESTING_SCRIPT_ID]`
- **Spreadsheet:** [Link to Testing MoM]

### Production Environment
- **Script ID:** `[PRODUCTION_SCRIPT_ID]`
- **Spreadsheet:** [Link to Production MoM]

---

## Deployment Workflow

### Option 1: Manual Script ID Switch (Current)

```bash
cd projects/qa-dashboard

# Push to Testing (default)
clasp push

# Push to Production
# 1. Edit .clasp.json - change scriptId to production
# 2. Run:
clasp push
# 3. Revert .clasp.json back to testing scriptId
```

### Option 2: Using Helper Scripts (Recommended)

We provide helper scripts to simplify deployment:

```bash
cd projects/qa-dashboard

# Push to testing
./deploy-testing.sh

# Push to production
./deploy-production.sh

# Push to both
./deploy-all.sh
```

---

## Best Practices

1. **Always test in Testing environment first** before pushing to Production
2. **Default `.clasp.json` should point to Testing** to prevent accidental production deployments
3. **Document all script IDs** in this file when adding new projects
4. **Use deployment scripts** to avoid manual script ID changes
5. **Verify environment** before running destructive operations (e.g., database changes)

---

## For AI Assistants

When asked to deploy or push changes:

1. **Default behavior:** Push to **Testing environment only**
2. **If user says "push to production":** Push to **Production environment** using the script IDs above
3. **If user says "push to both":** Push to Testing first, then Production
4. **Always revert `.clasp.json`** back to Testing script ID after pushing to Production
5. **Read this file** to get the correct script IDs for each environment

---

**Last Updated:** 2026-03-07
