'use client';

/**
 * Closure Email Page
 * Location: app/(dashboard)/closure-email/page.tsx
 *
 * Feature: Send test closure email dengan PDF attachment
 * Data source: External Test Report dari Apps Script API
 */

import { useState } from 'react';
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import {
  getExternalTestReports,
  sendClosureEmail,
  createClosureEmailDraft
} from '@/lib/api/closure-email';

export default function ClosureEmailPage() {
  const queryClient = useQueryClient();

  // Filters
  const [projectFilter, setProjectFilter] = useState('');
  const [statusFilter, setStatusFilter] = useState('');
  const [selectedModules, setSelectedModules] = useState<string[]>([]);

  // Email form
  const [emailTo, setEmailTo] = useState('');
  const [emailCc, setEmailCc] = useState('');
  const [emailSubject, setEmailSubject] = useState('Test Closure Report');
  const [customBody, setCustomBody] = useState('');
  const [showEmailForm, setShowEmailForm] = useState(false);

  // Fetch external test reports
  const { data: reports, isLoading, error } = useQuery({
    queryKey: ['external-reports', projectFilter, statusFilter],
    queryFn: () => getExternalTestReports({ project: projectFilter, status: statusFilter }),
    refetchInterval: 60000, // Auto-refresh every 60s
  });

  // Send email mutation
  const sendEmailMutation = useMutation({
    mutationFn: sendClosureEmail,
    onSuccess: () => {
      alert('Email sent successfully!');
      setShowEmailForm(false);
      setSelectedModules([]);
      setEmailTo('');
      setEmailCc('');
    },
    onError: (error: any) => {
      alert('Failed to send email: ' + error.message);
    },
  });

  // Create draft mutation
  const createDraftMutation = useMutation({
    mutationFn: createClosureEmailDraft,
    onSuccess: (data) => {
      alert(`Draft created! Check Gmail drafts (ID: ${data.draftId})`);
      window.open('https://mail.google.com/mail/#drafts', '_blank');
      setShowEmailForm(false);
      setSelectedModules([]);
    },
    onError: (error: any) => {
      alert('Failed to create draft: ' + error.message);
    },
  });

  // Toggle module selection
  const toggleModule = (moduleId: string) => {
    setSelectedModules(prev =>
      prev.includes(moduleId)
        ? prev.filter(id => id !== moduleId)
        : [...prev, moduleId]
    );
  };

  // Handle send email
  const handleSendEmail = () => {
    if (!emailTo) {
      alert('Please enter recipient email');
      return;
    }
    if (selectedModules.length === 0) {
      alert('Please select at least one module');
      return;
    }

    // Send email for each selected module
    selectedModules.forEach(moduleId => {
      sendEmailMutation.mutate({
        moduleId,
        emailTo,
        emailCc,
        emailSubject,
        emailBody: customBody,
        attachPDF: true,
      });
    });
  };

  // Handle create draft
  const handleCreateDraft = () => {
    if (!emailTo) {
      alert('Please enter recipient email');
      return;
    }
    if (selectedModules.length === 0) {
      alert('Please select at least one module');
      return;
    }

    // Create draft for each selected module
    selectedModules.forEach(moduleId => {
      createDraftMutation.mutate({
        moduleId,
        emailTo,
        emailCc,
        emailSubject,
        emailBody: customBody,
        attachPDF: true,
      });
    });
  };

  // Get unique projects for filter
  const projects = Array.from(new Set(reports?.map(r => r.project) || []));

  if (isLoading) {
    return (
      <div className="flex items-center justify-center h-screen">
        <div className="animate-spin rounded-full h-32 w-32 border-b-2 border-blue-600" />
      </div>
    );
  }

  if (error) {
    return (
      <div className="p-8">
        <div className="bg-red-50 border border-red-200 rounded-lg p-4">
          <p className="text-red-800">Error loading external test reports: {(error as Error).message}</p>
        </div>
      </div>
    );
  }

  return (
    <div className="p-8 max-w-7xl mx-auto">
      {/* Header */}
      <div className="mb-8">
        <h1 className="text-3xl font-bold text-gray-900">Test Closure Email</h1>
        <p className="text-gray-600 mt-2">
          Send closure email dengan PDF attachment berdasarkan External Test Report
        </p>
      </div>

      {/* Filters */}
      <div className="bg-white rounded-lg shadow p-6 mb-6">
        <h2 className="text-lg font-semibold mb-4">Filter</h2>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Project
            </label>
            <select
              value={projectFilter}
              onChange={(e) => setProjectFilter(e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
            >
              <option value="">All Projects</option>
              {projects.map(project => (
                <option key={project} value={project}>{project}</option>
              ))}
            </select>
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Status
            </label>
            <select
              value={statusFilter}
              onChange={(e) => setStatusFilter(e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
            >
              <option value="">All Status</option>
              <option value="Ready for Closure">Ready for Closure</option>
              <option value="Approved">Approved</option>
              <option value="In Review">In Review</option>
              <option value="Not Started">Not Started</option>
            </select>
          </div>
          <div className="flex items-end">
            <button
              onClick={() => setShowEmailForm(true)}
              disabled={selectedModules.length === 0}
              className="w-full px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 disabled:bg-gray-400 disabled:cursor-not-allowed transition"
            >
              Prepare Email ({selectedModules.length})
            </button>
          </div>
        </div>
      </div>

      {/* Summary Cards */}
      <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-6">
        <div className="bg-white p-6 rounded-lg shadow">
          <h3 className="text-gray-500 text-sm">Total Modules</h3>
          <p className="text-3xl font-bold">{reports?.length || 0}</p>
        </div>
        <div className="bg-white p-6 rounded-lg shadow">
          <h3 className="text-gray-500 text-sm">Ready for Closure</h3>
          <p className="text-3xl font-bold text-green-600">
            {reports?.filter(r => r.overallStatus === 'Ready for Closure').length || 0}
          </p>
        </div>
        <div className="bg-white p-6 rounded-lg shadow">
          <h3 className="text-gray-500 text-sm">In Review</h3>
          <p className="text-3xl font-bold text-yellow-600">
            {reports?.filter(r => r.overallStatus === 'In Review').length || 0}
          </p>
        </div>
        <div className="bg-white p-6 rounded-lg shadow">
          <h3 className="text-gray-500 text-sm">Selected</h3>
          <p className="text-3xl font-bold text-blue-600">
            {selectedModules.length}
          </p>
        </div>
      </div>

      {/* External Test Reports Table */}
      <div className="bg-white rounded-lg shadow overflow-hidden mb-6">
        <table className="min-w-full divide-y divide-gray-200">
          <thead className="bg-gray-50">
            <tr>
              <th className="px-6 py-3 text-left">
                <input
                  type="checkbox"
                  checked={selectedModules.length === reports?.length && reports?.length > 0}
                  onChange={(e) => {
                    if (e.target.checked) {
                      setSelectedModules(reports?.map(r => r.moduleId) || []);
                    } else {
                      setSelectedModules([]);
                    }
                  }}
                  className="rounded border-gray-300"
                />
              </th>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase">
                Project / Module
              </th>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase">
                External Team
              </th>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase">
                Overall Status
              </th>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase">
                Reviewer
              </th>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase">
                Review Date
              </th>
            </tr>
          </thead>
          <tbody className="bg-white divide-y divide-gray-200">
            {reports?.map((report) => (
              <tr key={report.moduleId} className="hover:bg-gray-50">
                <td className="px-6 py-4 whitespace-nowrap">
                  <input
                    type="checkbox"
                    checked={selectedModules.includes(report.moduleId)}
                    onChange={() => toggleModule(report.moduleId)}
                    className="rounded border-gray-300"
                  />
                </td>
                <td className="px-6 py-4">
                  <div className="text-sm font-medium text-gray-900">
                    {report.project}
                  </div>
                  <div className="text-sm text-gray-500">
                    {report.module} - {report.submodule}
                  </div>
                </td>
                <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                  {report.externalTeam || '-'}
                </td>
                <td className="px-6 py-4 whitespace-nowrap">
                  <span className={`px-2 py-1 text-xs rounded-full ${
                    report.overallStatus === 'Approved' || report.overallStatus === 'Ready for Closure'
                      ? 'bg-green-100 text-green-800' :
                    report.overallStatus === 'In Review'
                      ? 'bg-yellow-100 text-yellow-800' :
                    report.overallStatus === 'Rejected'
                      ? 'bg-red-100 text-red-800' :
                    'bg-gray-100 text-gray-800'
                  }`}>
                    {report.overallStatus || 'Not Started'}
                  </span>
                </td>
                <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                  {report.reviewer || '-'}
                </td>
                <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                  {report.reviewDate || '-'}
                </td>
              </tr>
            ))}
          </tbody>
        </table>

        {reports?.length === 0 && (
          <div className="text-center py-12 text-gray-500">
            No external test reports found. Try adjusting filters.
          </div>
        )}
      </div>

      {/* Email Form Modal */}
      {showEmailForm && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
          <div className="bg-white rounded-lg shadow-xl p-8 max-w-2xl w-full mx-4 max-h-[90vh] overflow-y-auto">
            <div className="flex justify-between items-center mb-6">
              <h2 className="text-2xl font-bold">Prepare Closure Email</h2>
              <button
                onClick={() => setShowEmailForm(false)}
                className="text-gray-400 hover:text-gray-600"
              >
                <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                </svg>
              </button>
            </div>

            <div className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  To: <span className="text-red-500">*</span>
                </label>
                <input
                  type="email"
                  value={emailTo}
                  onChange={(e) => setEmailTo(e.target.value)}
                  placeholder="stakeholder@example.com"
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                  required
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Cc: (optional)
                </label>
                <input
                  type="email"
                  value={emailCc}
                  onChange={(e) => setEmailCc(e.target.value)}
                  placeholder="manager@example.com"
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Subject: <span className="text-red-500">*</span>
                </label>
                <input
                  type="text"
                  value={emailSubject}
                  onChange={(e) => setEmailSubject(e.target.value)}
                  placeholder="Test Closure Report - Project X"
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                  required
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Custom Body (optional)
                </label>
                <textarea
                  value={customBody}
                  onChange={(e) => setCustomBody(e.target.value)}
                  placeholder="Leave empty to use auto-generated template"
                  rows={4}
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                />
                <p className="text-sm text-gray-500 mt-1">
                  Kosongkan untuk menggunakan template otomatis dengan data dari External Test Report
                </p>
              </div>

              <div className="bg-blue-50 border border-blue-200 rounded-lg p-4">
                <h4 className="font-semibold text-blue-900 mb-2">Selected Modules ({selectedModules.length}):</h4>
                <ul className="text-sm text-blue-800">
                  {reports
                    ?.filter(r => selectedModules.includes(r.moduleId))
                    .map(r => (
                      <li key={r.moduleId}>
                        {r.project} - {r.module} - {r.submodule}
                      </li>
                    ))}
                </ul>
              </div>

              <div className="flex gap-3 pt-4">
                <button
                  onClick={handleCreateDraft}
                  disabled={createDraftMutation.isPending}
                  className="flex-1 px-6 py-3 bg-blue-600 text-white rounded-md hover:bg-blue-700 disabled:bg-gray-400 disabled:cursor-not-allowed transition font-medium"
                >
                  {createDraftMutation.isPending ? 'Creating...' : '📝 Create Draft'}
                </button>
                <button
                  onClick={handleSendEmail}
                  disabled={sendEmailMutation.isPending}
                  className="flex-1 px-6 py-3 bg-green-600 text-white rounded-md hover:bg-green-700 disabled:bg-gray-400 disabled:cursor-not-allowed transition font-medium"
                >
                  {sendEmailMutation.isPending ? 'Sending...' : '✉️ Send Now'}
                </button>
              </div>

              <p className="text-xs text-gray-500 text-center">
                💡 Tip: Gunakan "Create Draft" untuk review email di Gmail sebelum send
              </p>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
