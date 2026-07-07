# 🚀 Apps Script to Next.js Migration Plan

**Project:** QA Dashboard Data Processing Migration
**Date:** 2026-07-06
**Status:** Planning Phase
**Dashboard Live:** ✅ https://qa-platform.inadigital.co.id

---

## 📋 Executive Summary

**Current State:**
- ✅ **Next.js Dashboard:** Already live and production-ready (visualization layer)
- ❌ **Apps Script Backend:** Still handling all data processing (slow, 60-90s per refresh)

**Goal:**
Migrate **data processing logic** from Apps Script to Next.js backend APIs while keeping dashboard UI unchanged.

**Impact:**
- 🚀 **10x faster refresh** (60s → 6s)
- ✅ **No execution limits** (6-min timeout → unlimited)
- ✅ **Real-time updates** (webhooks instead of polling)
- ✅ **Better scalability** (parallel operations)

---

## 🎯 Migration Scope

### **What to Migrate**

| Function | Current (Apps Script) | Target (Next.js API) | Priority |
|----------|----------------------|---------------------|----------|
| **refreshDashboard()** | Pull from QATMs → Write to Dashboard | `/api/refresh-dashboard` | 🔴 HIGH |
| **refreshAutomationHistory()** | Read Automation Runs → Write History | `/api/refresh-automation` | 🔴 HIGH |
| **Jenkins webhook** | Python script → Sheets | `/api/automation-webhook` | 🟡 MEDIUM |
| **Scheduled refresh** | Apps Script time trigger | Jenkins Cron / K8s CronJob | 🟢 LOW |

### **What to Keep (For Now)**

| Feature | Reason | Future Plan |
|---------|--------|-------------|
| Email automation | Gmail API simpler in Apps Script | Phase 3: Migrate to Next.js |
| Sheet formatting | Apps Script SpreadsheetApp easier | Keep in Apps Script |
| Manual menu items | UI convenience | Keep in Apps Script |

---

## 🏗️ Target Architecture

```
┌──────────────────────────────────────────────────────────────┐
│                  QA Test Management (QATM)                    │
│              Individual Module Spreadsheets                   │
│                                                               │
│  ┌──────────┐  ┌──────────┐  ┌──────────┐                  │
│  │ Module 1 │  │ Module 2 │  │ Module 3 │  ...              │
│  │  QATM    │  │  QATM    │  │  QATM    │                  │
│  └────┬─────┘  └────┬─────┘  └────┬─────┘                  │
└───────┼─────────────┼─────────────┼─────────────────────────┘
        │             │             │
        │   Google Sheets API v4 (Service Account)
        │             │             │
┌───────▼─────────────▼─────────────▼─────────────────────────┐
│                                                               │
│           📊 QA Dashboard Spreadsheet                        │
│          (Central aggregated data storage)                   │
│                                                               │
│  Tabs: Config | Overview | Bugs | VAPT | History |          │
│        Automation Runs | VAPT History                        │
│                                                               │
└───────┬──────────────────────────────────────────────┬──────┘
        │                                               │
        │  Google Sheets API (Read/Write)              │
        │                                               │
┌───────▼──────────────────────────────────────────────▼──────┐
│                                                               │
│       🚀 Next.js Backend (NEW - API Routes)                 │
│                                                               │
│  ┌────────────────────────────────────────────────────┐     │
│  │  POST /api/refresh-dashboard                       │     │
│  │  ├─ getConfigModules() → Read Config              │     │
│  │  ├─ pullAllModuleData() → Parallel QATM fetch     │     │
│  │  │  (10x faster than Apps Script!)                 │     │
│  │  └─ writeDashboardTabs() → Update sheets          │     │
│  │                                                     │     │
│  │  POST /api/refresh-automation                      │     │
│  │  ├─ getAutomationRuns() → Read Automation Runs   │     │
│  │  ├─ buildPatches() → Build history patches        │     │
│  │  └─ upsertHistory() → Update History tab          │     │
│  │                                                     │     │
│  │  POST /api/automation-webhook                      │     │
│  │  ├─ validateJenkins() → Auth check                │     │
│  │  ├─ appendAutomationRun() → Write to sheet        │     │
│  │  └─ triggerRefresh() → Background job             │     │
│  └────────────────────────────────────────────────────┘     │
│                                                               │
│  Libraries:                                                   │
│  ✅ googleapis (Google Sheets API v4)                        │
│  ✅ Service Account auth (no OAuth needed)                   │
│  ✅ Parallel Promise.all() operations                        │
│                                                               │
└───────┬──────────────────────────────────────────────┬──────┘
        │                                               │
        │  REST API (JSON)                             │
        │                                               │
┌───────▼──────────────────────────────────────────────▼──────┐
│                                                               │
│        💻 Next.js Frontend (EXISTING - No Changes)          │
│                   Already Live & Working                     │
│                                                               │
│  /dashboard        → Charts, filters, real-time updates     │
│  /dashboard/bugs   → Bug metrics                            │
│  /dashboard/vapt   → Security findings                      │
│                                                               │
└───────────────────────────────────────────────────────────────┘

┌──────────────────────────────────────────────────────────────┐
│                                                               │
│              ⏰ Scheduled Refresh (NEW)                      │
│                                                               │
│  Option A: Jenkins Cron Job                                  │
│  ├─ Schedule: */10 * * * * (every 10 min)                  │
│  └─ Action: curl POST /api/refresh-dashboard                │
│                                                               │
│  Option B: Kubernetes CronJob                                │
│  ├─ Schedule: */10 * * * *                                  │
│  └─ Action: curl POST /api/refresh-dashboard                │
│                                                               │
└───────────────────────────────────────────────────────────────┘
```

---

## 📊 Performance Analysis

### **Current Bottlenecks (Apps Script)**

```javascript
// MasterDashboard.js - SLOW Sequential Operations
function refreshDashboard() {
  const modules = getModuleList_(ss); // 20 modules

  // ❌ BOTTLENECK #1: Sequential QATM fetches
  modules.forEach(module => {
    const ss = SpreadsheetApp.openById(module.spreadsheetId);
    // ⏱️ 2-3 seconds PER module!

    const summary = ss.getSheetByName('Summary');
    const data = summary.getDataRange().getValues();
    // ⏱️ 1-2 seconds per sheet read
  });

  // Total: 20 modules × 3s = 60 seconds minimum!

  // ❌ BOTTLENECK #2: Sequential writes
  writeOverview(ss, allData);      // 3-5s
  writeBugs(ss, allData);          // 3-5s
  writeVAPT(ss, allData);          // 2-3s
  writeSmoke(ss, allData);         // 2-3s
  writeFailureScenario(ss, allData); // 1-2s
  writeCoverage(ss, allData);      // 1-2s
  appendHistory(ss, allData);      // 2-3s

  // Total: ~20 seconds for writes

  // 💥 GRAND TOTAL: 80-90 seconds!
}
```

**Why so slow?**
1. `SpreadsheetApp.openById()` - Heavy API call, not optimized
2. Sequential operations - No parallelization
3. Apps Script engine overhead
4. Network latency per spreadsheet operation

---

### **Target Performance (Next.js API)**

```typescript
// POST /api/refresh-dashboard - FAST Parallel Operations
export async function POST(request: Request) {
  const modules = await getConfigModules(); // 20 modules

  // ✅ OPTIMIZATION #1: Parallel QATM fetches
  const allData = await Promise.all(
    modules.map(async (module) => {
      // Google Sheets API v4 - Direct, optimized calls
      const data = await sheets.spreadsheets.values.batchGet({
        spreadsheetId: module.spreadsheetId,
        ranges: ['Summary!A:Z', 'BUG BLOCKER!A:Z', 'VAPT!A:Z']
      });
      return parseModuleData(data);
    })
  );
  // ⏱️ 3-5 seconds for ALL 20 modules in parallel!

  // ✅ OPTIMIZATION #2: Parallel writes
  await Promise.all([
    updateOverviewSheet(allData),      // Parallel
    updateBugsSheet(allData),          // Parallel
    updateVAPTSheet(allData),          // Parallel
    updateHistorySheet(allData),       // Parallel
  ]);
  // ⏱️ 2-3 seconds for ALL writes in parallel!

  // 🚀 GRAND TOTAL: 5-8 seconds!
}
```

**Why so fast?**
1. Google Sheets API v4 - Optimized, batch operations
2. Parallel Promise.all() - No waiting
3. Node.js performance - Fast runtime
4. Batch operations - Fewer API calls

---

### **Performance Comparison**

| Metric | Apps Script | Next.js API | Improvement |
|--------|------------|-------------|-------------|
| **Pull 20 modules** | 60s (sequential) | 5s (parallel) | **12x faster** |
| **Write 7 tabs** | 20s (sequential) | 3s (parallel) | **6.7x faster** |
| **Total refresh** | 80-90s | 8-10s | **9x faster** |
| **API response** | 2-5s | 50-200ms | **20x faster** |
| **Concurrent users** | ~5 (slow) | 100+ (fast) | **20x scale** |
| **Execution limit** | 6 minutes | No limit | **Unlimited** |

---

## 🛠️ Implementation Plan

### **Phase 1: Core Refresh API (4 weeks)** 🔴

**Goal:** Replace `refreshDashboard()` with Next.js API

#### **Week 1-2: Read Operations**

**Tasks:**
1. ✅ Implement `getConfigModules()` - Already exists in `google-sheets.ts`
2. ✅ Implement `pullModuleDataFromQATM()` - Batch read from QATM Summary/Bugs/VAPT
3. ✅ Add parallel fetching with `Promise.all()`
4. ✅ Error handling & retry logic
5. ✅ Logging & monitoring

**Deliverables:**
```typescript
// src/lib/qatm-fetcher.ts (NEW)
export async function pullAllModuleData(modules: Module[]): Promise<ModuleData[]> {
  return await Promise.all(
    modules.map(module => pullModuleDataFromQATM(module))
  );
}

async function pullModuleDataFromQATM(module: Module): Promise<ModuleData> {
  const sheets = await getSheetsClient();

  // Batch read multiple tabs at once
  const response = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: module.spreadsheetId,
    ranges: [
      'Summary!A1:AZ1000',
      'BUG BLOCKER!A1:Z1000',
      'VAPT!A1:Z1000',
      'Coverage!A1:Z1000'
    ],
    majorDimension: 'ROWS'
  });

  return parseModuleData(response.data.valueRanges);
}
```

**Testing:**
```bash
# Test local
curl -X POST http://localhost:3000/api/test-fetch \
  -H 'Content-Type: application/json' \
  -d '{"moduleLimit": 5}'

# Verify: Should complete in <5 seconds for 5 modules
```

---

#### **Week 3-4: Write Operations**

**Tasks:**
1. ✅ Implement `updateOverviewSheet()`
2. ✅ Implement `updateBugsSheet()`
3. ✅ Implement `updateVAPTSheet()`
4. ✅ Implement `updateHistorySheet()` - Preserve automation columns!
5. ✅ Add batch update operations
6. ✅ Transaction-like behavior (rollback on error)

**Deliverables:**
```typescript
// src/lib/dashboard-writer.ts (NEW)
export async function writeDashboardTabs(allData: ModuleData[]): Promise<void> {
  const sheets = await getSheetsClient();

  // Parallel writes to all tabs
  await Promise.all([
    updateOverviewSheet(sheets, allData),
    updateBugsSheet(sheets, allData),
    updateVAPTSheet(sheets, allData),
    updateSmokeSheet(sheets, allData),
    updateFailureScenarioSheet(sheets, allData),
    updateCoverageSheet(sheets, allData),
    updateHistorySheet(sheets, allData), // Preserve automation columns!
  ]);
}

async function updateOverviewSheet(sheets: any, allData: ModuleData[]): Promise<void> {
  const rows = allData.map(data => buildOverviewRow(data));

  // Batch update - single API call
  await sheets.spreadsheets.values.update({
    spreadsheetId: DASHBOARD_SPREADSHEET_ID,
    range: 'Overview!A6:AZ1000',
    valueInputOption: 'USER_ENTERED',
    resource: { values: rows }
  });
}
```

**Critical: History Sheet Preservation**
```typescript
// src/lib/dashboard-writer.ts
async function updateHistorySheet(sheets: any, allData: ModuleData[]): Promise<void> {
  // ⚠️ IMPORTANT: Read existing automation columns first!
  const existingHistory = await sheets.spreadsheets.values.get({
    spreadsheetId: DASHBOARD_SPREADSHEET_ID,
    range: 'History!A3:ZZ1000'
  });

  const existingRows = existingHistory.data.values || [];
  const historyIndex = buildHistoryIndex(existingRows);

  // Build patches (exclude automation fields)
  const patches = allData.map(data => ({
    key: `${formatDate(now)}|${data.project}|${data.module}|${data.submodule}`,
    values: buildHistoryCoreFields(data) // No automation fields!
  }));

  // Merge with existing automation data
  const finalRows = patches.map(patch => {
    const existing = historyIndex.get(patch.key);
    return mergeHistoryRow(existing, patch); // Preserve automation columns
  });

  // Write merged rows
  await sheets.spreadsheets.values.update({
    spreadsheetId: DASHBOARD_SPREADSHEET_ID,
    range: 'History!A3:ZZ1000',
    valueInputOption: 'USER_ENTERED',
    resource: { values: finalRows }
  });
}
```

**Testing:**
```bash
# Test write (dry-run first)
curl -X POST http://localhost:3000/api/test-write \
  -H 'Content-Type: application/json' \
  -d '{"dryRun": true, "moduleLimit": 2}'

# Verify: Check Dashboard spreadsheet, no data loss
```

---

#### **API Route Implementation**

```typescript
// src/app/api/refresh-dashboard/route.ts (NEW)
export async function POST(request: Request) {
  const startTime = Date.now();

  try {
    // 1. Get modules from Config
    const modules = await getConfigModules();

    if (modules.length === 0) {
      return Response.json({
        error: 'No active modules in Config'
      }, { status: 400 });
    }

    // 2. Pull data from QATMs (parallel)
    console.log(`Fetching data from ${modules.length} modules...`);
    const allData = await pullAllModuleData(modules);

    // 3. Write to Dashboard tabs (parallel)
    console.log('Writing to Dashboard tabs...');
    await writeDashboardTabs(allData);

    // 4. Update timestamps
    await updateRefreshTimestamps();

    const duration = Date.now() - startTime;

    return Response.json({
      success: true,
      modules: modules.length,
      duration: `${(duration / 1000).toFixed(1)}s`,
      timestamp: new Date().toISOString()
    });

  } catch (error) {
    console.error('Dashboard refresh failed:', error);

    return Response.json({
      success: false,
      error: error.message,
      duration: `${((Date.now() - startTime) / 1000).toFixed(1)}s`
    }, { status: 500 });
  }
}
```

**Testing:**
```bash
# Production test
curl -X POST https://qa-platform.inadigital.co.id/api/refresh-dashboard

# Expected response:
{
  "success": true,
  "modules": 20,
  "duration": "8.2s",
  "timestamp": "2026-07-06T14:30:00.000Z"
}
```

---

### **Phase 2: Automation History API (2 weeks)** 🟡

**Goal:** Replace `refreshAutomationHistory()` with Next.js API

#### **Week 1: Read & Build Logic**

**Tasks:**
1. ✅ Implement `getAutomationEnabledModules()`
2. ✅ Implement `getAutomationRunsFromSheet()`
3. ✅ Port `buildAutomationHistoryPatch()` logic
4. ✅ Port `getLatestRunsByEnvironment()` logic

**Deliverables:**
```typescript
// src/lib/automation-processor.ts (NEW)
export async function buildAutomationPatches(
  modules: Module[],
  runs: AutomationRun[]
): Promise<HistoryPatch[]> {

  // Group runs by key: environment|channel|contractKey
  const runsByKey = new Map<string, AutomationRun>();

  runs.forEach(run => {
    const key = `${run.environment}|${run.channel}|${run.contractKey}`;
    const existing = runsByKey.get(key);

    if (!existing || run.timestamp > existing.timestamp) {
      runsByKey.set(key, run);
    }
  });

  // Build patches for each module
  return modules.map(module => {
    const webRuns = getModuleRuns(module, runsByKey, 'web');
    const apiRuns = getModuleRuns(module, runsByKey, 'api');

    return {
      timestamp: new Date(),
      project: module.project,
      module: module.module,
      submodule: module.submodule,
      values: {
        // Web automation
        webAutomationPassed: webRuns.latest?.passed,
        webAutomationFailed: webRuns.latest?.failed,
        webAutomationPassRate: webRuns.latest?.passRate,
        webAutomationStatus: webRuns.latest?.status || 'No Run',

        webDevPassed: webRuns.dev?.passed,
        webDevFailed: webRuns.dev?.failed,
        // ... more fields

        // API automation (same structure)
        // ...
      }
    };
  });
}
```

---

#### **Week 2: Write & Integration**

**Tasks:**
1. ✅ Implement `upsertHistorySheet()` - Smart merge
2. ✅ Add locking mechanism (prevent concurrent updates)
3. ✅ API route implementation
4. ✅ Testing & validation

**Deliverables:**
```typescript
// src/app/api/refresh-automation/route.ts (NEW)
export async function POST(request: Request) {
  const startTime = Date.now();

  try {
    // 1. Get automation-enabled modules
    const modules = await getAutomationEnabledModules();

    if (modules.length === 0) {
      return Response.json({
        message: 'No modules with Web/API Enabled',
        updated: 0
      });
    }

    // 2. Get latest automation runs
    const runs = await getAutomationRunsFromSheet();

    // 3. Build patches
    const patches = await buildAutomationPatches(modules, runs);

    // 4. Upsert to History (with lock)
    const result = await upsertHistorySheet(patches);

    const duration = Date.now() - startTime;

    return Response.json({
      success: true,
      modules: modules.length,
      updated: result.updated,
      appended: result.appended,
      duration: `${(duration / 1000).toFixed(1)}s`
    });

  } catch (error) {
    console.error('Automation refresh failed:', error);

    return Response.json({
      success: false,
      error: error.message
    }, { status: 500 });
  }
}
```

**Testing:**
```bash
# Test automation refresh
curl -X POST https://qa-platform.inadigital.co.id/api/refresh-automation

# Expected response:
{
  "success": true,
  "modules": 15,
  "updated": 12,
  "appended": 3,
  "duration": "3.2s"
}
```

---

### **Phase 3: Jenkins Webhook Integration (2 weeks)** 🟢

**Goal:** Real-time automation updates from Jenkins

#### **Week 1: Webhook Endpoint**

**Tasks:**
1. ✅ Implement webhook authentication
2. ✅ Implement `appendAutomationRun()`
3. ✅ Add background job trigger
4. ✅ Error handling & retry

**Deliverables:**
```typescript
// src/app/api/automation-webhook/route.ts (NEW)
export async function POST(request: Request) {
  const startTime = Date.now();

  try {
    // 1. Validate Jenkins token
    const token = request.headers.get('X-Jenkins-Token');

    if (!validateJenkinsToken(token)) {
      return Response.json({
        error: 'Unauthorized'
      }, { status: 401 });
    }

    // 2. Parse payload
    const payload = await request.json();

    // 3. Validate payload structure
    if (!isValidAutomationPayload(payload)) {
      return Response.json({
        error: 'Invalid payload'
      }, { status: 400 });
    }

    // 4. Append to Automation Runs sheet
    await appendAutomationRunToSheet(payload);

    // 5. Trigger background refresh (async, don't wait)
    fetch('https://qa-platform.inadigital.co.id/api/refresh-automation', {
      method: 'POST'
    }).catch(err => console.error('Background refresh failed:', err));

    const duration = Date.now() - startTime;

    return Response.json({
      success: true,
      duration: `${duration}ms`,
      message: 'Automation run recorded, history refresh triggered'
    });

  } catch (error) {
    console.error('Webhook failed:', error);

    return Response.json({
      success: false,
      error: error.message
    }, { status: 500 });
  }
}

// Helper: Validate Jenkins token
function validateJenkinsToken(token: string | null): boolean {
  if (!token) return false;

  const validToken = process.env.JENKINS_WEBHOOK_TOKEN;
  if (!validToken) {
    console.warn('JENKINS_WEBHOOK_TOKEN not configured!');
    return true; // Allow for now (add token later)
  }

  return token === validToken;
}
```

---

#### **Week 2: Jenkins Pipeline Update**

**Tasks:**
1. ✅ Update `qa-library/vars/dashboardUtils.groovy`
2. ✅ Replace Python script with curl webhook call
3. ✅ Add token authentication
4. ✅ Testing & rollout

**Deliverables:**
```groovy
// qa-library/vars/dashboardUtils.groovy (MODIFIED)
def publishAutomationResult(String testType, String buildResult, String reportUrl, String commitId, def env, def params) {
    def jobInfo = jobNameParser.parseJobName(env.JOB_NAME)
    def stats = getStatsFromEnv(testType, env)
    def channel = normalizeTestType(testType)

    if (channel != 'web' && channel != 'api') {
        echo "QA Dashboard publish skipped: unsupported testType=${testType}"
        return [skipped: true, reason: 'unsupported_type']
    }

    def payload = [
        timestamp  : new Date().format("yyyy-MM-dd'T'HH:mm:ssXXX", TimeZone.getTimeZone('Asia/Jakarta')),
        source     : 'jenkins',
        channel    : channel,
        environment: jobInfo.env ?: '',
        project    : jobInfo.project ?: '',
        module     : jobInfo.moduleId ?: '',
        submodule  : jobInfo.submodule ?: '',
        contractKey: firstNonBlank(readParam(params, 'DASHBOARD_AUTOMATION_CONTRACT'), env.DASHBOARD_AUTOMATION_CONTRACT),
        tag        : env.EFFECTIVE_QA_SERVICE ?: '',
        jobName    : env.JOB_NAME ?: '',
        buildNumber: env.BUILD_NUMBER ?: '',
        buildUrl   : env.BUILD_URL ?: '',
        reportUrl  : reportUrl ?: '',
        commitId   : commitId ?: '',
        buildStatus: buildResult ?: 'SUCCESS',
        status     : deriveStatus(stats, buildResult),
        total      : stats.total,
        passed     : stats.passed,
        failed     : stats.failed,
        broken     : stats.broken,
        skipped    : stats.skipped,
        flaky      : stats.flaky
    ]

    // NEW: Call Next.js webhook instead of Python script
    return callNextJsWebhook(payload, env)
}

// NEW: Call Next.js webhook
private Map callNextJsWebhook(Map payload, def env) {
    def webhookUrl = 'https://qa-platform.inadigital.co.id/api/automation-webhook'
    def webhookToken = env.JENKINS_WEBHOOK_TOKEN ?: ''

    // Write payload to temp file
    writeFile file: 'qa-dashboard-payload.json', text: JsonOutput.toJson(payload)

    try {
        def response = sh(
            script: """
                set +x
                curl -X POST '${webhookUrl}' \\
                    -H 'Content-Type: application/json' \\
                    -H 'X-Jenkins-Token: ${webhookToken}' \\
                    -d @qa-dashboard-payload.json \\
                    --max-time 30
            """,
            returnStdout: true
        ).trim()

        echo "QA Dashboard webhook response: ${response}"

        // Cleanup
        sh 'rm -f qa-dashboard-payload.json'

        return [skipped: false, response: response]

    } catch (Exception e) {
        echo "QA Dashboard webhook failed: ${e.getMessage()}"

        // Cleanup
        sh 'rm -f qa-dashboard-payload.json'

        return [skipped: false, error: e.getMessage()]
    }
}
```

**Testing:**
```bash
# Trigger test automation job
# Jenkins should call webhook automatically

# Verify webhook received
kubectl logs -f deployment/qa-platform -n qa-prod | grep "Webhook"

# Check Automation Runs sheet
# Should have new row with latest run
```

---

### **Phase 4: Scheduled Refresh (1 week)** ⏰

**Goal:** Replace Apps Script time triggers

**Option A: Jenkins Cron Job (Recommended)**

Create new Jenkins job: `qa-dashboard-refresh-cron`

```groovy
pipeline {
    agent any

    triggers {
        cron('*/10 * * * *')  // Every 10 minutes
    }

    stages {
        stage('Refresh Dashboard') {
            steps {
                script {
                    def response = sh(
                        script: """
                            curl -X POST https://qa-platform.inadigital.co.id/api/refresh-dashboard \\
                                -H 'Content-Type: application/json' \\
                                --max-time 120
                        """,
                        returnStdout: true
                    ).trim()

                    echo "Dashboard refresh response: ${response}"

                    // Parse JSON response
                    def result = readJSON text: response

                    if (!result.success) {
                        error("Dashboard refresh failed: ${result.error}")
                    }
                }
            }
        }

        stage('Refresh Automation') {
            steps {
                script {
                    def response = sh(
                        script: """
                            curl -X POST https://qa-platform.inadigital.co.id/api/refresh-automation \\
                                -H 'Content-Type: application/json' \\
                                --max-time 60
                        """,
                        returnStdout: true
                    ).trim()

                    echo "Automation refresh response: ${response}"
                }
            }
        }
    }

    post {
        failure {
            // Send alert on failure
            echo "Dashboard refresh failed! Check logs."
        }
    }
}
```

**Option B: Kubernetes CronJob**

```yaml
# k8s/cronjob-refresh-dashboard.yaml
apiVersion: batch/v1
kind: CronJob
metadata:
  name: qa-dashboard-refresh
  namespace: qa-prod
spec:
  schedule: "*/10 * * * *"  # Every 10 minutes
  successfulJobsHistoryLimit: 3
  failedJobsHistoryLimit: 1
  jobTemplate:
    spec:
      template:
        spec:
          containers:
          - name: refresh
            image: curlimages/curl:latest
            command:
            - /bin/sh
            - -c
            - |
              echo "Refreshing dashboard..."
              curl -X POST https://qa-platform.inadigital.co.id/api/refresh-dashboard -f

              echo "Refreshing automation..."
              curl -X POST https://qa-platform.inadigital.co.id/api/refresh-automation -f

              echo "Refresh completed"
          restartPolicy: OnFailure
```

**Deploy:**
```bash
kubectl apply -f k8s/cronjob-refresh-dashboard.yaml

# Verify
kubectl get cronjobs -n qa-prod
kubectl get jobs -n qa-prod
```

---

## 🗓️ Timeline Summary

| Phase | Duration | Deliverables |
|-------|----------|--------------|
| **Phase 1: Core Refresh API** | 4 weeks | `/api/refresh-dashboard` fully working |
| **Phase 2: Automation API** | 2 weeks | `/api/refresh-automation` functional |
| **Phase 3: Jenkins Webhook** | 2 weeks | Real-time automation updates |
| **Phase 4: Scheduled Refresh** | 1 week | Replace Apps Script triggers |
| **Testing & Rollout** | 1 week | Production deployment |
| **TOTAL** | **10 weeks** | Complete migration |

---

## ✅ Success Criteria

### **Phase 1 Complete**
- [ ] `/api/refresh-dashboard` responds in <15s for 20 modules
- [ ] All Dashboard tabs updated correctly
- [ ] History automation columns preserved (no data loss)
- [ ] Zero errors in production logs
- [ ] A/B test: Apps Script vs Next.js results match 100%

### **Phase 2 Complete**
- [ ] `/api/refresh-automation` responds in <5s
- [ ] Automation data correctly materialized to History
- [ ] All environments (dev/stg/prod) tracked correctly
- [ ] Zero data inconsistencies

### **Phase 3 Complete**
- [ ] Jenkins webhook delivers data in real-time (<30s delay)
- [ ] Automation Runs sheet auto-updated
- [ ] History sheet auto-refreshed after webhook
- [ ] 100% webhook success rate

### **Phase 4 Complete**
- [ ] Scheduled refresh runs every 10 minutes
- [ ] Apps Script triggers disabled
- [ ] Zero downtime during migration
- [ ] Team fully trained on new system

---

## 🚨 Risk Management

### **Risk 1: Data Loss During Migration**

**Mitigation:**
- ✅ Keep Apps Script running during Phase 1-3
- ✅ Parallel testing (both systems run simultaneously)
- ✅ History preservation logic (never overwrite automation columns)
- ✅ Daily backups of Dashboard spreadsheet
- ✅ Rollback plan (revert to Apps Script if needed)

**Rollback Procedure:**
```bash
# If Next.js API fails, Apps Script is still running
# No action needed - system continues working

# To manually trigger Apps Script:
# Dashboard → QA Dashboard Menu → 🔄 Refresh Dashboard
```

---

### **Risk 2: Performance Degradation**

**Mitigation:**
- ✅ Load testing before production deployment
- ✅ Monitoring & alerting setup
- ✅ Gradual rollout (10% → 50% → 100%)
- ✅ Circuit breaker pattern (fallback to Apps Script)

**Load Test Plan:**
```bash
# Test with 20 modules
ab -n 10 -c 1 https://qa-platform.inadigital.co.id/api/refresh-dashboard

# Expected: <15s per request, 100% success rate
```

---

### **Risk 3: Google Sheets API Quota**

**Current Quotas:**
- 500 requests per 100 seconds per user
- 100 requests per 100 seconds per project

**Mitigation:**
- ✅ Use batch operations (reduce API calls by 10x)
- ✅ Service Account (separate quota pool)
- ✅ Implement exponential backoff & retry
- ✅ Cache frequently accessed data (Redis/Vercel KV)

**Quota Usage Calculation:**
```
Refresh Dashboard (20 modules):
- Config read: 1 request
- QATM batchGet (20 modules × 1 batch): 20 requests
- Dashboard writes (7 tabs × 1 update): 7 requests
Total: 28 requests per refresh

Every 10 minutes = 6 refreshes/hour
= 168 requests/hour
= Well within quota (500/100s = 18,000/hour)
```

---

## 📊 Monitoring & Alerts

### **Key Metrics to Track**

```typescript
// src/lib/monitoring.ts (NEW)
export interface RefreshMetrics {
  duration: number;        // Total time in ms
  modulesProcessed: number;
  errors: number;
  timestamp: Date;
}

export async function logMetrics(metrics: RefreshMetrics) {
  console.log(JSON.stringify({
    type: 'dashboard_refresh',
    ...metrics
  }));

  // Future: Send to monitoring service (Datadog, New Relic, etc.)
}
```

**Alerts to Configure:**
1. 🚨 Refresh duration >30s (warning)
2. 🚨 Refresh duration >60s (critical)
3. 🚨 Error rate >1% (warning)
4. 🚨 Error rate >5% (critical)
5. 🚨 Webhook failure (immediate)

---

## 🎓 Team Training

### **Required Skills**

| Role | Current (Apps Script) | New (Next.js) | Training Needed |
|------|----------------------|---------------|-----------------|
| **Backend Dev** | JavaScript basics | TypeScript + async/await | 2 weeks |
| **QA Lead** | Apps Script menu | API calls via curl | 1 week |
| **DevOps** | Apps Script triggers | Jenkins cron / K8s CronJob | 1 week |

### **Training Plan**

**Week 1: TypeScript Basics**
- Variables, types, interfaces
- Async/await, Promises
- Error handling

**Week 2: Google Sheets API**
- Service Account auth
- Read operations (get, batchGet)
- Write operations (update, batchUpdate)

**Week 3: Next.js API Routes**
- Route handlers (POST /api/*)
- Request/Response handling
- Environment variables

**Week 4: Testing & Debugging**
- Local testing with `npm run dev`
- Production testing with curl
- Reading logs (kubectl logs)

---

## 💰 Cost Analysis

| Component | Current (Apps Script) | New (Next.js) | Difference |
|-----------|----------------------|---------------|------------|
| **Apps Script** | FREE | FREE (keep for email) | $0 |
| **Next.js Hosting** | N/A | FREE (existing K8s) | $0 |
| **Google Sheets API** | FREE (within quota) | FREE (within quota) | $0 |
| **Jenkins** | FREE (existing) | FREE (existing) | $0 |
| **Total** | **$0/month** | **$0/month** | **$0** |

**Conclusion:** Zero cost increase! 🎉

---

## 📝 Documentation Updates

### **Files to Create**

1. ✅ `src/lib/qatm-fetcher.ts` - QATM data fetching logic
2. ✅ `src/lib/dashboard-writer.ts` - Dashboard write operations
3. ✅ `src/lib/automation-processor.ts` - Automation history logic
4. ✅ `src/app/api/refresh-dashboard/route.ts` - Main refresh API
5. ✅ `src/app/api/refresh-automation/route.ts` - Automation refresh API
6. ✅ `src/app/api/automation-webhook/route.ts` - Jenkins webhook handler
7. ✅ `k8s/cronjob-refresh-dashboard.yaml` - K8s cron job config

### **Files to Update**

1. ✅ `qa-library/vars/dashboardUtils.groovy` - Replace Python with webhook
2. ✅ `README.md` - Add migration documentation
3. ✅ `.env.local` - Add `JENKINS_WEBHOOK_TOKEN`

---

## 🎯 Decision Points

### **Week 4: Go/No-Go Decision**

**Criteria for GO:**
- ✅ Phase 1 APIs fully tested
- ✅ Performance meets targets (<15s)
- ✅ Zero data loss in testing
- ✅ Team trained on new system

**Criteria for NO-GO:**
- ❌ Performance worse than Apps Script
- ❌ Data inconsistencies detected
- ❌ Team not confident in new system

**If NO-GO:**
- Continue using Apps Script
- Reassess approach
- Re-plan timeline

---

## 📞 Support & Resources

### **Code References**

**Existing (Working):**
- `src/lib/google-sheets.ts` - Google Sheets API client
- `src/lib/google-workspace.ts` - Gmail API client
- `src/app/api/bugs/route.ts` - Example read-only API
- `src/app/api/history/route.ts` - Example read-only API

**Apps Script (Reference):**
- `projects/qa-dashboard/src/MasterDashboard.js` - Current implementation
- `projects/qa-dashboard/src/BugsTab.js` - Bug aggregation logic
- `projects/qa-dashboard/src/VAPTTab.js` - VAPT aggregation logic

### **Key Contacts**

- **Tech Lead:** For architecture decisions
- **QA Lead:** For testing & validation
- **DevOps:** For Jenkins & K8s setup

---

## 🎉 Summary

**Current State:**
- ✅ Next.js dashboard **live** (read-only)
- ❌ Apps Script backend (slow, 60-90s refresh)

**Target State:**
- ✅ Next.js dashboard (no changes)
- ✅ Next.js backend APIs (10x faster)
- ✅ Real-time webhook integration
- ✅ Jenkins cron for scheduled refresh
- ⚠️ Apps Script (keep for email only)

**Timeline:** 10 weeks (with 1 week buffer)

**Cost:** $0 (no additional infrastructure)

**Risk:** Low (parallel running, easy rollback)

**Impact:** 🚀 **10x faster, unlimited scale**

---

**Ready to start? Let's migrate Phase 1!** 🚀
