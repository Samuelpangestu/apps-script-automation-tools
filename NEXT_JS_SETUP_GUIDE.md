# Next.js + QA Dashboard Integration Guide

## Phase 1: Next.js Frontend dengan Apps Script Backend

### Architecture

```
┌─────────────────────────────────────────────────────────┐
│                    Next.js Frontend                      │
│  - Modern UI (Tailwind + shadcn/ui)                     │
│  - Fast navigation                                       │
│  - Real-time updates                                     │
│  - Mobile responsive                                     │
└─────────────────┬───────────────────────────────────────┘
                  │ HTTPS API Calls
                  ↓
┌─────────────────────────────────────────────────────────┐
│              Apps Script Web App (API)                   │
│  - Existing dashboard functions                          │
│  - Data validation                                       │
│  - Business logic                                        │
└─────────────────┬───────────────────────────────────────┘
                  │ Sheets API
                  ↓
┌─────────────────────────────────────────────────────────┐
│                   Google Sheets                          │
│  - QA Dashboard data                                     │
│  - Bug reports, VAPT, KPI, Teams                        │
└─────────────────────────────────────────────────────────┘
```

---

## Option 1: Via Apps Script Web App (RECOMMENDED)

### Step 1: Update Apps Script untuk expose API

File: `projects/qa-dashboard/src/WebAppAPI.js`

```javascript
/**
 * WebAppAPI.js - REST API for Next.js frontend
 *
 * This file exposes existing dashboard functions as API endpoints
 * that can be called from Next.js
 */

/**
 * Main doGet handler - handles GET requests from Next.js
 */
function doGet(e) {
  try {
    const action = e.parameter.action;
    const params = e.parameter;

    let result;
    switch (action) {
      case 'getBugsData':
        result = getBugsDataAPI(params);
        break;
      case 'getVAPTData':
        result = getVAPTDataAPI(params);
        break;
      case 'getKPIData':
        result = getKPIDataAPI(params);
        break;
      case 'getTeamsData':
        result = getTeamsDataAPI(params);
        break;
      default:
        result = { error: 'Unknown action' };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        error: error.message,
        stack: error.stack
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Main doPost handler - handles POST requests from Next.js
 */
function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;
    const params = postData.params;

    let result;
    switch (action) {
      case 'updateBugStatus':
        result = updateBugStatusAPI(params);
        break;
      case 'addComment':
        result = addCommentAPI(params);
        break;
      case 'updateVAPTFinding':
        result = updateVAPTFindingAPI(params);
        break;
      default:
        result = { error: 'Unknown action' };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        error: error.message,
        stack: error.stack
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Get Bugs Data API
 */
function getBugsDataAPI(params) {
  const ss = SpreadsheetApp.openById(DASHBOARD_SPREADSHEET_ID);
  const configSheet = ss.getSheetByName('Config');
  const data = configSheet.getDataRange().getValues();

  const bugs = [];
  for (let i = 3; i < data.length; i++) {
    if (!data[i][0]) continue; // Skip inactive

    const qatmId = data[i][6];
    if (!qatmId) continue;

    try {
      const qatm = SpreadsheetApp.openById(qatmId);
      const bugSheet = qatm.getSheetByName('Bug Report');
      if (!bugSheet) continue;

      const bugData = bugSheet.getDataRange().getValues();

      // Parse bugs (skip headers row 1-2)
      for (let j = 2; j < bugData.length; j++) {
        if (!bugData[j][0]) continue; // Skip if no bug ID

        bugs.push({
          id: bugData[j][0],
          project: data[i][2],
          module: data[i][3],
          title: bugData[j][1],
          status: bugData[j][2],
          priority: bugData[j][3],
          severity: bugData[j][4],
          assignee: bugData[j][5],
          reporter: bugData[j][6],
          createdDate: bugData[j][7],
          resolvedDate: bugData[j][8],
          description: bugData[j][9],
        });
      }
    } catch (e) {
      Logger.log('Error reading QATM: ' + qatmId + ' - ' + e.message);
    }
  }

  return {
    success: true,
    data: bugs,
    total: bugs.length,
    timestamp: new Date().toISOString()
  };
}

/**
 * Get VAPT Data API
 */
function getVAPTDataAPI(params) {
  const ss = SpreadsheetApp.openById(DASHBOARD_SPREADSHEET_ID);
  const configSheet = ss.getSheetByName('Config');
  const data = configSheet.getDataRange().getValues();

  const findings = [];
  for (let i = 3; i < data.length; i++) {
    if (!data[i][0]) continue;

    const qatmId = data[i][6];
    if (!qatmId) continue;

    try {
      const qatm = SpreadsheetApp.openById(qatmId);
      const vaptSheet = qatm.getSheetByName('VAPT - Detail Finding');
      if (!vaptSheet) continue;

      const vaptData = vaptSheet.getDataRange().getValues();

      // Parse findings (skip headers row 1-2)
      for (let j = 2; j < vaptData.length; j++) {
        if (!vaptData[j][0]) continue; // Skip if no finding ID

        findings.push({
          id: vaptData[j][0],
          project: data[i][2],
          module: data[i][3],
          app: vaptData[j][1],
          statusFix: vaptData[j][4],
          statusReVAPT: vaptData[j][5],
          risk: vaptData[j][6],
          adjustedRisk: vaptData[j][7],
          findingName: vaptData[j][8],
          reportDate: vaptData[j][14],
          alreadyInProd: vaptData[j][26],
        });
      }
    } catch (e) {
      Logger.log('Error reading VAPT: ' + qatmId + ' - ' + e.message);
    }
  }

  return {
    success: true,
    data: findings,
    total: findings.length,
    timestamp: new Date().toISOString()
  };
}

/**
 * Update Bug Status API
 */
function updateBugStatusAPI(params) {
  const { qatmId, bugId, newStatus } = params;

  try {
    const qatm = SpreadsheetApp.openById(qatmId);
    const bugSheet = qatm.getSheetByName('Bug Report');
    const data = bugSheet.getDataRange().getValues();

    // Find bug row
    for (let i = 2; i < data.length; i++) {
      if (data[i][0] === bugId) {
        bugSheet.getRange(i + 1, 3).setValue(newStatus); // Col C = Status

        return {
          success: true,
          message: 'Bug status updated',
          bugId: bugId,
          newStatus: newStatus
        };
      }
    }

    return {
      success: false,
      error: 'Bug not found'
    };

  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

// Add more API functions as needed...
```

### Step 2: Deploy Apps Script as Web App

1. Open Apps Script Editor
2. Click **Deploy** → **New deployment**
3. Type: **Web app**
4. Execute as: **Me**
5. Who has access: **Anyone** (or specific domain)
6. Click **Deploy**
7. Copy Web App URL: `https://script.google.com/macros/s/[DEPLOYMENT_ID]/exec`

---

## Step 3: Setup Next.js Project

### Create Next.js app

```bash
npx create-next-app@latest qa-engineering-platform
# ✔ Would you like to use TypeScript? Yes
# ✔ Would you like to use ESLint? Yes
# ✔ Would you like to use Tailwind CSS? Yes
# ✔ Would you like to use `src/` directory? Yes
# ✔ Would you like to use App Router? Yes
# ✔ Would you like to customize the default import alias? No

cd qa-engineering-platform
```

### Install dependencies

```bash
npm install @tanstack/react-query axios zustand
npm install -D @types/node
```

### Project Structure

```
qa-engineering-platform/
├── src/
│   ├── app/
│   │   ├── (dashboard)/
│   │   │   ├── layout.tsx
│   │   │   ├── page.tsx              # Dashboard home
│   │   │   ├── bugs/
│   │   │   │   └── page.tsx          # Bugs dashboard
│   │   │   ├── vapt/
│   │   │   │   └── page.tsx          # VAPT dashboard
│   │   │   └── kpi/
│   │   │       └── page.tsx          # KPI dashboard
│   │   └── api/
│   │       └── dashboard/
│   │           ├── bugs/route.ts     # Proxy to Apps Script
│   │           └── vapt/route.ts
│   ├── components/
│   │   ├── ui/                       # shadcn components
│   │   ├── dashboard/
│   │   │   ├── BugsTable.tsx
│   │   │   ├── VAPTChart.tsx
│   │   │   └── KPICards.tsx
│   │   └── layout/
│   │       ├── Header.tsx
│   │       └── Sidebar.tsx
│   └── lib/
│       ├── api/
│       │   └── apps-script.ts        # Apps Script client
│       ├── utils/
│       └── constants.ts
├── .env.local
└── package.json
```

---

## Step 4: Create Apps Script API Client

File: `src/lib/api/apps-script.ts`

```typescript
/**
 * Apps Script API Client
 *
 * This client communicates with the Apps Script Web App
 * to fetch and update data from Google Sheets
 */

const APPS_SCRIPT_WEB_APP_URL = process.env.NEXT_PUBLIC_APPS_SCRIPT_URL!;

export interface Bug {
  id: string;
  project: string;
  module: string;
  title: string;
  status: string;
  priority: string;
  severity: string;
  assignee: string;
  reporter: string;
  createdDate: string;
  resolvedDate?: string;
  description: string;
}

export interface VAPTFinding {
  id: string;
  project: string;
  module: string;
  app: string;
  statusFix: string;
  statusReVAPT: string;
  risk: string;
  adjustedRisk: string;
  findingName: string;
  reportDate: string;
  alreadyInProd: string;
}

export interface ApiResponse<T> {
  success: boolean;
  data?: T;
  error?: string;
  timestamp: string;
}

/**
 * Fetch bugs data from Apps Script
 */
export async function getBugs(): Promise<Bug[]> {
  const url = `${APPS_SCRIPT_WEB_APP_URL}?action=getBugsData`;

  const response = await fetch(url, {
    method: 'GET',
    headers: {
      'Content-Type': 'application/json',
    },
    // Next.js cache options
    next: {
      revalidate: 60, // Revalidate every 60 seconds
    },
  });

  if (!response.ok) {
    throw new Error('Failed to fetch bugs');
  }

  const result: ApiResponse<Bug[]> = await response.json();

  if (!result.success) {
    throw new Error(result.error || 'Unknown error');
  }

  return result.data || [];
}

/**
 * Fetch VAPT findings from Apps Script
 */
export async function getVAPTFindings(): Promise<VAPTFinding[]> {
  const url = `${APPS_SCRIPT_WEB_APP_URL}?action=getVAPTData`;

  const response = await fetch(url, {
    method: 'GET',
    next: {
      revalidate: 60,
    },
  });

  if (!response.ok) {
    throw new Error('Failed to fetch VAPT findings');
  }

  const result: ApiResponse<VAPTFinding[]> = await response.json();

  if (!result.success) {
    throw new Error(result.error || 'Unknown error');
  }

  return result.data || [];
}

/**
 * Update bug status
 */
export async function updateBugStatus(
  qatmId: string,
  bugId: string,
  newStatus: string
): Promise<void> {
  const response = await fetch(APPS_SCRIPT_WEB_APP_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      action: 'updateBugStatus',
      params: {
        qatmId,
        bugId,
        newStatus,
      },
    }),
  });

  if (!response.ok) {
    throw new Error('Failed to update bug status');
  }

  const result: ApiResponse<any> = await response.json();

  if (!result.success) {
    throw new Error(result.error || 'Unknown error');
  }
}
```

---

## Step 5: Create Next.js API Routes (Optional Proxy)

File: `src/app/api/dashboard/bugs/route.ts`

```typescript
/**
 * Next.js API Route - Bugs
 *
 * This is optional - acts as a proxy to Apps Script
 * Useful for:
 * - Adding authentication
 * - Rate limiting
 * - Caching
 * - Logging
 */

import { NextRequest, NextResponse } from 'next/server';

const APPS_SCRIPT_URL = process.env.APPS_SCRIPT_URL!;

export async function GET(request: NextRequest) {
  try {
    // You can add auth check here
    // const session = await getServerSession();
    // if (!session) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const response = await fetch(`${APPS_SCRIPT_URL}?action=getBugsData`, {
      method: 'GET',
      next: { revalidate: 60 }, // Cache for 60 seconds
    });

    const data = await response.json();

    return NextResponse.json(data);
  } catch (error) {
    console.error('Error fetching bugs:', error);
    return NextResponse.json(
      { error: 'Failed to fetch bugs' },
      { status: 500 }
    );
  }
}

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();

    const response = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        action: 'updateBugStatus',
        params: body,
      }),
    });

    const data = await response.json();

    return NextResponse.json(data);
  } catch (error) {
    console.error('Error updating bug:', error);
    return NextResponse.json(
      { error: 'Failed to update bug' },
      { status: 500 }
    );
  }
}
```

---

## Step 6: Create Dashboard Components

File: `src/app/(dashboard)/bugs/page.tsx`

```typescript
'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { getBugs, updateBugStatus } from '@/lib/api/apps-script';

export default function BugsPage() {
  const queryClient = useQueryClient();

  // Fetch bugs using React Query
  const { data: bugs, isLoading, error } = useQuery({
    queryKey: ['bugs'],
    queryFn: getBugs,
    refetchInterval: 60000, // Auto-refresh every 60s
  });

  // Mutation for updating bug status
  const updateStatusMutation = useMutation({
    mutationFn: ({ qatmId, bugId, status }: {
      qatmId: string;
      bugId: string;
      status: string;
    }) => updateBugStatus(qatmId, bugId, status),
    onSuccess: () => {
      // Invalidate and refetch bugs
      queryClient.invalidateQueries({ queryKey: ['bugs'] });
    },
  });

  if (isLoading) {
    return (
      <div className="flex items-center justify-center h-screen">
        <div className="animate-spin rounded-full h-32 w-32 border-b-2 border-gray-900" />
      </div>
    );
  }

  if (error) {
    return (
      <div className="p-8 text-center">
        <p className="text-red-600">Error loading bugs: {error.message}</p>
      </div>
    );
  }

  return (
    <div className="p-8">
      <h1 className="text-3xl font-bold mb-6">Bug Dashboard</h1>

      <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-8">
        <div className="bg-white p-6 rounded-lg shadow">
          <h3 className="text-gray-500 text-sm">Total Bugs</h3>
          <p className="text-3xl font-bold">{bugs?.length || 0}</p>
        </div>
        <div className="bg-white p-6 rounded-lg shadow">
          <h3 className="text-gray-500 text-sm">Open</h3>
          <p className="text-3xl font-bold text-red-600">
            {bugs?.filter(b => b.status === 'Open').length || 0}
          </p>
        </div>
        <div className="bg-white p-6 rounded-lg shadow">
          <h3 className="text-gray-500 text-sm">In Progress</h3>
          <p className="text-3xl font-bold text-yellow-600">
            {bugs?.filter(b => b.status === 'In Progress').length || 0}
          </p>
        </div>
        <div className="bg-white p-6 rounded-lg shadow">
          <h3 className="text-gray-500 text-sm">Resolved</h3>
          <p className="text-3xl font-bold text-green-600">
            {bugs?.filter(b => b.status === 'Resolved').length || 0}
          </p>
        </div>
      </div>

      <div className="bg-white rounded-lg shadow overflow-hidden">
        <table className="min-w-full divide-y divide-gray-200">
          <thead className="bg-gray-50">
            <tr>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase">
                ID
              </th>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase">
                Project
              </th>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase">
                Title
              </th>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase">
                Status
              </th>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase">
                Priority
              </th>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase">
                Assignee
              </th>
            </tr>
          </thead>
          <tbody className="bg-white divide-y divide-gray-200">
            {bugs?.map((bug) => (
              <tr key={bug.id} className="hover:bg-gray-50">
                <td className="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">
                  {bug.id}
                </td>
                <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                  {bug.project} - {bug.module}
                </td>
                <td className="px-6 py-4 text-sm text-gray-900">
                  {bug.title}
                </td>
                <td className="px-6 py-4 whitespace-nowrap">
                  <span className={`px-2 py-1 text-xs rounded-full ${
                    bug.status === 'Open' ? 'bg-red-100 text-red-800' :
                    bug.status === 'In Progress' ? 'bg-yellow-100 text-yellow-800' :
                    'bg-green-100 text-green-800'
                  }`}>
                    {bug.status}
                  </span>
                </td>
                <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                  {bug.priority}
                </td>
                <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                  {bug.assignee}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
```

---

## Step 7: Setup Environment Variables

File: `.env.local`

```bash
# Apps Script Web App URL
NEXT_PUBLIC_APPS_SCRIPT_URL=https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec

# Optional: If using server-side API routes
APPS_SCRIPT_URL=https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec
```

---

## Step 8: Run Development Server

```bash
npm run dev
```

Visit: `http://localhost:3000/bugs`

---

## Performance Optimization

### 1. Add Caching

```typescript
// In apps-script.ts

import { cache } from 'react';

export const getBugs = cache(async () => {
  // ... fetch logic
});
```

### 2. Use Server Components

```typescript
// src/app/(dashboard)/bugs/page.tsx

import { getBugs } from '@/lib/api/apps-script';

// Server Component (faster initial load)
export default async function BugsPage() {
  const bugs = await getBugs();

  return (
    <BugsTable bugs={bugs} />
  );
}
```

### 3. Add Loading States

```typescript
// src/app/(dashboard)/bugs/loading.tsx

export default function Loading() {
  return (
    <div className="p-8">
      <div className="animate-pulse">
        <div className="h-8 bg-gray-200 rounded w-1/4 mb-6"></div>
        <div className="grid grid-cols-4 gap-4 mb-8">
          {[1, 2, 3, 4].map(i => (
            <div key={i} className="h-24 bg-gray-200 rounded"></div>
          ))}
        </div>
        <div className="h-96 bg-gray-200 rounded"></div>
      </div>
    </div>
  );
}
```

---

## Deployment

### Deploy to Vercel

```bash
# Install Vercel CLI
npm i -g vercel

# Deploy
vercel

# Production deployment
vercel --prod
```

### Environment Variables in Vercel

1. Go to Vercel Dashboard
2. Select project
3. Settings → Environment Variables
4. Add `NEXT_PUBLIC_APPS_SCRIPT_URL`

---

## Testing

```bash
# Run tests
npm run test

# E2E tests with Playwright
npm run test:e2e
```

---

## Summary

✅ **What You Get:**
- Modern, fast Next.js frontend
- Direct connection to existing Google Sheets data
- No infrastructure change needed
- Apps Script logic intact
- Mobile responsive
- Real-time updates
- Easy to extend

✅ **Timeline:**
- Day 1: Setup Next.js + Apps Script API
- Day 2-3: Build Bug Dashboard
- Day 4-5: Build VAPT Dashboard
- Day 6-7: Styling + optimization
- Day 8: Testing + deployment

✅ **Cost:**
- Vercel Free tier (100GB bandwidth/month)
- Google Apps Script (Free)
- Total: $0/month

---

## Next Steps

1. ✅ Create Apps Script API endpoints
2. ⏳ Setup Next.js project
3. ⏳ Build first dashboard (Bugs)
4. ⏳ Deploy to Vercel
5. ⏳ Get feedback
6. ⏳ Iterate & improve

---

**Ready to start building?** 🚀
