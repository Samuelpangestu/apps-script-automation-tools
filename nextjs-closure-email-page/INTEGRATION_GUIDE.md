# Closure Email Integration Guide

## 📁 File Locations

Copy files ke Next.js project:

```
qa-platform/
├── src/
│   ├── app/
│   │   └── (dashboard)/
│   │       └── closure-email/
│   │           └── page.tsx          ← Copy: page.tsx
│   └── lib/
│       └── api/
│           └── closure-email.ts      ← Copy: closure-email.ts
```

---

## 🔧 Setup Steps

### 1. Install Dependencies (if not already installed)

```bash
cd qa-platform
npm install @tanstack/react-query
```

### 2. Update Navigation

Add closure email link to sidebar/navigation:

**File: `src/components/layout/Sidebar.tsx`** (or wherever nav is)

```tsx
const navItems = [
  { name: 'Dashboard', href: '/', icon: HomeIcon },
  { name: 'Bugs', href: '/bugs', icon: BugIcon },
  { name: 'VAPT', href: '/vapt', icon: ShieldIcon },
  { name: 'KPI', href: '/kpi', icon: ChartIcon },
  // Add this:
  { name: 'Closure Email', href: '/closure-email', icon: MailIcon },
];
```

### 3. Environment Variables

**Already setup in `.env.local`:**
```bash
NEXT_PUBLIC_APPS_SCRIPT_URL=https://script.google.com/macros/s/AKfycbyyVuslzjevRSZ3x37MXP-QWYVJcTqd4LgO8xWIZcaA/exec
```

✅ No changes needed!

### 4. Test Locally

```bash
# Development
npm run dev

# Visit
http://localhost:3000/closure-email
```

### 5. Deploy to Production

```bash
# Using Vercel
vercel --prod

# Or GitHub push (if auto-deploy enabled)
git add .
git commit -m "feat: Add closure email page"
git push origin main
```

---

## 🧪 Testing Checklist

### Local Testing

- [ ] Page loads without errors
- [ ] External test reports fetch successfully
- [ ] Filters work (project, status)
- [ ] Checkbox selection works
- [ ] Email form modal opens
- [ ] "Create Draft" button creates Gmail draft
- [ ] "Send Now" button sends email
- [ ] Error handling works (try invalid email)

### Production Testing

- [ ] Page accessible at https://qa-platform.inadigital.co.id/closure-email
- [ ] API calls work from production
- [ ] Email sent successfully
- [ ] PDF attachment received
- [ ] Gmail draft created in departemen.qa inbox

---

## 🎨 Customization

### Change Colors

```tsx
// In page.tsx, change Tailwind classes:

// Primary color (currently blue-600)
className="bg-blue-600 hover:bg-blue-700"
// Change to:
className="bg-indigo-600 hover:bg-indigo-700"

// Success color (currently green-600)
className="bg-green-600"
// etc.
```

### Add More Filters

```tsx
// In page.tsx, add filter state:
const [moduleFilter, setModuleFilter] = useState('');

// Add to query:
queryFn: () => getExternalTestReports({
  project: projectFilter,
  module: moduleFilter,  // Add this
  status: statusFilter
}),
```

### Custom Email Templates

```tsx
// Instead of auto-generated, provide custom body:
<textarea
  value={customBody}
  onChange={(e) => setCustomBody(e.target.value)}
  placeholder="Enter custom email body (HTML supported)"
/>
```

---

## 🔍 Troubleshooting

### "Failed to fetch external test reports"

**Check:**
1. Apps Script deployed? `clasp deployments`
2. ENV variable correct? Check `.env.local`
3. Authorization done? Open Apps Script URL in browser first

**Fix:**
```bash
# Re-deploy Apps Script
cd projects/qa-dashboard
clasp push
```

### "External Test Report tab not found"

**Cause:** QATM belum punya tab External Test Report

**Fix:** Dari Dashboard Apps Script:
1. Menu → Broadcast Fixes
2. Click "External QA: Create Report Tab"
3. Wait for broadcast complete
4. Refresh page

### CORS Errors

**Fix:** Apps Script deployment settings:
1. Open Apps Script Editor
2. Deploy → Manage deployments
3. Edit deployment
4. "Who has access": Change to **Anyone**
5. Save

### Email not sending

**Check:**
1. Apps Script authorized dengan akun departemen.qa?
2. Gmail quota exceeded? (limit ~100 emails/day for free)
3. Check Apps Script logs: Executions tab

**Fix:** Use "Create Draft" instead of "Send Now" if quota exceeded

---

## 📊 Performance Optimization

### Add Loading Skeleton

```tsx
// Create loading.tsx in same folder
export default function Loading() {
  return (
    <div className="p-8 animate-pulse">
      <div className="h-8 bg-gray-200 rounded w-1/3 mb-6"></div>
      <div className="h-96 bg-gray-200 rounded"></div>
    </div>
  );
}
```

### Add Client-side Caching

```tsx
// In lib/api/closure-email.ts
import { cache } from 'react';

export const getExternalTestReports = cache(async (params) => {
  // ... existing code
});
```

### Debounce Filters

```tsx
import { useDebouncedValue } from '@/hooks/useDebouncedValue';

// In component:
const debouncedProject = useDebouncedValue(projectFilter, 300);

useQuery({
  queryKey: ['external-reports', debouncedProject, statusFilter],
  // ...
});
```

---

## 🚀 Next Steps

1. **Add Bulk Actions:**
   - Select all ready for closure
   - Send to multiple recipients
   - Schedule emails

2. **Email Templates:**
   - Save custom templates
   - Template library
   - Variables/placeholders

3. **History Tracking:**
   - Log sent emails
   - View email history
   - Resend functionality

4. **Notifications:**
   - Slack/Teams notification when email sent
   - Email delivery status tracking

---

## ✅ Done!

Feature sudah ready untuk production. Tinggal:

1. Copy 2 files ke Next.js project
2. Test locally
3. Deploy to production
4. Access https://qa-platform.inadigital.co.id/closure-email

**Questions?** Check `API_CLOSURE_EMAIL.md` untuk API reference lengkap.
