/**
 * Closure Email API Client
 * Location: lib/api/closure-email.ts
 *
 * Client untuk fetch External Test Reports dan send closure emails
 */

const APPS_SCRIPT_URL = process.env.NEXT_PUBLIC_APPS_SCRIPT_URL!;

// ═══════════════════════════════════════════════════════════════════════
// Type Definitions
// ═══════════════════════════════════════════════════════════════════════

export interface ExternalTestReport {
  moduleId: string;
  project: string;
  module: string;
  submodule: string;
  picQA: string;
  isExternalQA: boolean;
  externalTeam: string;
  statusReview: string;
  functionalEvidenceUrl: string;
  functionalReviewStatus: string;
  performanceEvidenceUrl: string;
  performanceReviewStatus: string;
  vaptEvidenceUrl: string;
  vaptReviewStatus: string;
  overallStatus: string;
  reviewer: string;
  reviewDate: string;
  notes: string;
  // VAPT Findings from QATM Summary (row 65-83)
  vaptTotal: number;
  vaptCritical: number;
  vaptHigh: number;
  vaptMedium: number;
  vaptLow: number;
  vaptInformational: number;
  vaptTodo: number;
  vaptOnProgress: number;
  vaptDone: number;
  vaptOpen: number;
  vaptClosed: number;
  vaptBlockerCount: number;       // Critical + High + Medium
  vaptNonBlockerCount: number;    // Low + Informational
  vaptApprovalStatus: string;     // Computed: ⚠️ NEEDS ATTENTION | ✅ APPROVED WITH NOTES | ✅ FULLY APPROVED
}

export interface GetReportsParams {
  project?: string;
  module?: string;
  status?: string;
}

export interface SendEmailParams {
  moduleId: string;
  emailTo: string;
  emailCc?: string;
  emailSubject: string;
  emailBody?: string;
  attachPDF?: boolean;
}

export interface ApiResponse<T> {
  success: boolean;
  data?: T;
  error?: string;
  message?: string;
  timestamp: string;
}

export interface SendEmailResponse {
  success: boolean;
  message: string;
  sentTo: string;
  sentCc: string;
  timestamp: string;
}

export interface CreateDraftResponse {
  success: boolean;
  message: string;
  draftId: string;
  recipient: string;
  subject: string;
  timestamp: string;
}

// ═══════════════════════════════════════════════════════════════════════
// API Functions
// ═══════════════════════════════════════════════════════════════════════

/**
 * Get External Test Reports dengan filter optional
 */
export async function getExternalTestReports(
  params: GetReportsParams = {}
): Promise<ExternalTestReport[]> {
  const url = new URL(APPS_SCRIPT_URL);
  url.searchParams.set('action', 'getExternalTestReports');

  if (params.project) {
    url.searchParams.set('project', params.project);
  }
  if (params.module) {
    url.searchParams.set('module', params.module);
  }
  if (params.status) {
    url.searchParams.set('status', params.status);
  }

  const response = await fetch(url.toString(), {
    method: 'GET',
    headers: {
      'Content-Type': 'application/json',
    },
    // Cache for 30 seconds (data relatively static)
    next: {
      revalidate: 30,
    },
  });

  if (!response.ok) {
    throw new Error(`Failed to fetch external test reports: ${response.statusText}`);
  }

  const result: ApiResponse<ExternalTestReport[]> = await response.json();

  if (!result.success) {
    throw new Error(result.error || 'Failed to fetch external test reports');
  }

  return result.data || [];
}

/**
 * Send closure email langsung dengan PDF attachment
 */
export async function sendClosureEmail(
  params: SendEmailParams
): Promise<SendEmailResponse> {
  const response = await fetch(APPS_SCRIPT_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      action: 'sendClosureEmail',
      moduleId: params.moduleId,
      emailTo: params.emailTo,
      emailCc: params.emailCc || '',
      emailSubject: params.emailSubject,
      emailBody: params.emailBody || '',
      attachPDF: params.attachPDF !== false, // default true
    }),
  });

  if (!response.ok) {
    throw new Error(`Failed to send email: ${response.statusText}`);
  }

  const result: SendEmailResponse = await response.json();

  if (!result.success) {
    throw new Error(result.message || 'Failed to send email');
  }

  return result;
}

/**
 * Create Gmail draft (tidak langsung send)
 */
export async function createClosureEmailDraft(
  params: SendEmailParams
): Promise<CreateDraftResponse> {
  const response = await fetch(APPS_SCRIPT_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      action: 'createClosureEmailDraft',
      moduleId: params.moduleId,
      emailTo: params.emailTo,
      emailCc: params.emailCc || '',
      emailSubject: params.emailSubject,
      emailBody: params.emailBody || '',
      attachPDF: params.attachPDF !== false,
    }),
  });

  if (!response.ok) {
    throw new Error(`Failed to create draft: ${response.statusText}`);
  }

  const result: CreateDraftResponse = await response.json();

  if (!result.success) {
    throw new Error(result.message || 'Failed to create draft');
  }

  return result;
}

/**
 * Helper: Generate default email subject
 */
export function generateDefaultSubject(
  project: string,
  module: string,
  submodule?: string
): string {
  let subject = `Test Closure Report - ${project}`;

  if (module && module !== '-') {
    subject += ` - ${module}`;
  }

  if (submodule) {
    subject += ` - ${submodule}`;
  }

  return subject;
}

/**
 * Helper: Validate email address
 */
export function isValidEmail(email: string): boolean {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}
