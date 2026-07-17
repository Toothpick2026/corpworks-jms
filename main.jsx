import React, { useState, useEffect } from 'react';
import { LayoutDashboard, Calendar, FileText, DollarSign, Truck, Cloud, LogOut, Check, AlertCircle, Loader, X, Trash2, Pencil, Plus } from 'lucide-react';

// ============================================================================
// CONSTANTS & HELPERS
// ============================================================================
const GST_RATE = 0.15; // NZ GST

// Escape user text before injecting into generated HTML files
const esc = (val) =>
  String(val ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');

// Pure total calculation from a quote object (does NOT read component state)
const computeQuoteTotals = (quote) => {
  const itemsTotal = (quote.lineItems || []).reduce(
    (sum, item) => sum + (Number(item.qty) || 0) * (Number(item.rate) || 0),
    0
  );
  const freight = Number(quote.freight) || 0;
  const travel = (Number(quote.travelHours) || 0) * (Number(quote.travelRate) || 0);
  const subtotal = itemsTotal + freight + travel;
  const gst = subtotal * GST_RATE;
  const total = subtotal + gst;
  return { itemsTotal, freight, travel, subtotal, gst, total };
};

const emptyQuote = () => ({
  quoteNum: '',
  client: '',
  endClient: '',
  date: new Date().toISOString().split('T')[0],
  preparedBy: 'Monique',
  jobType: 'Install',
  notes: '',
  lineItems: [],
  travelHours: 0,
  travelRate: 0,
  freight: 0
});

// ============================================================================
// ONEDRIVE INTEGRATION MODULE
// NOTE: Browser OAuth must NOT contain a client secret. This module now uses
// the PKCE public-client flow (no secret). Register the app in Azure as a SPA
// and set VITE_ONEDRIVE_CLIENT_ID / VITE_ONEDRIVE_REDIRECT_URI env vars.
// ============================================================================
const b64url = (bytes) =>
  btoa(String.fromCharCode(...new Uint8Array(bytes)))
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');

const OneDriveManager = {
  clientId: import.meta.env.VITE_ONEDRIVE_CLIENT_ID || 'b9c9a198-ecad-4041-9f15-ce32f34c0567',
  // Pin sign-in to the Corpworks tenant (not /common) so Azure can locate the app.
  tenant: import.meta.env.VITE_ONEDRIVE_TENANT || 'd5d1b2b1-9dd1-4628-b4a8-2b579a657f13',
  redirectUri: import.meta.env.VITE_ONEDRIVE_REDIRECT_URI || `${window.location.origin}/`,
  scopes: ['Files.ReadWrite', 'User.Read', 'offline_access'],

  async createPkceChallenge() {
    const verifierBytes = crypto.getRandomValues(new Uint8Array(32));
    const verifier = b64url(verifierBytes);
    const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier));
    const challenge = b64url(digest);
    sessionStorage.setItem('pkce_verifier', verifier);
    return challenge;
  },

  async getAuthUrl() {
    const challenge = await this.createPkceChallenge();
    const params = new URLSearchParams({
      client_id: this.clientId,
      redirect_uri: this.redirectUri,
      response_type: 'code',
      scope: this.scopes.join(' '),
      response_mode: 'query',
      code_challenge: challenge,
      code_challenge_method: 'S256'
    });
    return `https://login.microsoftonline.com/${this.tenant}/oauth2/v2.0/authorize?${params}`;
  },

  async exchangeCodeForToken(code) {
    try {
      const verifier = sessionStorage.getItem('pkce_verifier') || '';
      const response = await fetch(`https://login.microsoftonline.com/${this.tenant}/oauth2/v2.0/token`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id: this.clientId,
          code: code,
          redirect_uri: this.redirectUri,
          grant_type: 'authorization_code',
          code_verifier: verifier
        })
      });
      return await response.json();
    } catch (error) {
      console.error('Token exchange failed:', error);
      return null;
    }
  },

  async uploadFile(accessToken, fileName, fileContent, folderPath = 'Corpworks/Job Sheets') {
    try {
      const encodedPath = encodeURIComponent(`/${folderPath}/${fileName}`);
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/drive/root:${encodedPath}:/content`,
        {
          method: 'PUT',
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/octet-stream'
          },
          body: fileContent
        }
      );
      return response.ok;
    } catch (error) {
      console.error('OneDrive upload failed:', error);
      return false;
    }
  }
};

// ============================================================================
// MAIN APP COMPONENT
// ============================================================================
export default function App() {
  const [activeTab, setActiveTab] = useState('dashboard');
  const [oneDriveToken, setOneDriveToken] = useState(localStorage.getItem('onedrive_token'));
  const [uploadStatus, setUploadStatus] = useState(null);
  const [jobs, setJobs] = useState(() => {
    const saved = localStorage.getItem('corpworks_jobs');
    return saved ? JSON.parse(saved) : [];
  });
  const [quotes, setQuotes] = useState(() => {
    const saved = localStorage.getItem('corpworks_quotes');
    return saved ? JSON.parse(saved) : [];
  });
  const [viewMode, setViewMode] = useState('list');
  const [editingJob, setEditingJob] = useState(null); // job being edited in the modal
  const [currentQuote, setCurrentQuote] = useState(emptyQuote);

  // Check for OAuth callback
  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const code = params.get('code');
    if (code) {
      exchangeAuthCode(code);
    }
  }, []);

  // Persist jobs & quotes to localStorage
  useEffect(() => {
    localStorage.setItem('corpworks_jobs', JSON.stringify(jobs));
  }, [jobs]);
  useEffect(() => {
    localStorage.setItem('corpworks_quotes', JSON.stringify(quotes));
  }, [quotes]);

  // Auto-dismiss status messages
  useEffect(() => {
    if (uploadStatus && uploadStatus.type !== 'loading') {
      const timer = setTimeout(() => setUploadStatus(null), 4000);
      return () => clearTimeout(timer);
    }
  }, [uploadStatus]);

  const exchangeAuthCode = async (code) => {
    const tokenData = await OneDriveManager.exchangeCodeForToken(code);
    if (tokenData && tokenData.access_token) {
      setOneDriveToken(tokenData.access_token);
      localStorage.setItem('onedrive_token', tokenData.access_token);
      if (tokenData.refresh_token) localStorage.setItem('onedrive_refresh', tokenData.refresh_token);
      setUploadStatus({ type: 'success', message: '✓ OneDrive connected!' });
      window.history.replaceState({}, document.title, window.location.pathname);
    } else {
      setUploadStatus({ type: 'error', message: '✗ Failed to connect OneDrive' });
    }
  };

  const initiateOneDriveAuth = async () => {
    window.location.href = await OneDriveManager.getAuthUrl();
  };

  const handleSaveQuoteToOneDrive = async () => {
    if (!oneDriveToken) {
      setUploadStatus({ type: 'error', message: 'Please connect OneDrive first' });
      return;
    }
    if (!currentQuote.quoteNum) {
      setUploadStatus({ type: 'error', message: 'Please enter a quote number' });
      return;
    }
    setUploadStatus({ type: 'loading', message: 'Uploading quote to OneDrive...' });
    const quoteContent = generateQuoteHTML(currentQuote);
    const fileName = `Quote_${currentQuote.quoteNum}_${new Date().toISOString().split('T')[0]}.html`;
    const success = await OneDriveManager.uploadFile(
      oneDriveToken,
      fileName,
      new Blob([quoteContent], { type: 'text/html' }),
      'Corpworks/Quotes'
    );
    if (success) {
      setUploadStatus({ type: 'success', message: `✓ Quote saved to OneDrive: ${fileName}` });
      const newQuote = { ...currentQuote, id: Date.now(), savedAt: new Date().toISOString() };
      setQuotes([...quotes, newQuote]);
      clearQuoteForm();
    } else {
      setUploadStatus({ type: 'error', message: '✗ Failed to upload quote to OneDrive' });
    }
  };

  const handleSaveJobSheetToOneDrive = async (job) => {
    if (!oneDriveToken) {
      setUploadStatus({ type: 'error', message: 'Please connect OneDrive first' });
      return;
    }
    setUploadStatus({ type: 'loading', message: 'Uploading job sheet to OneDrive...' });
    const jobContent = generateJobSheetHTML(job);
    const fileName = `JobSheet_${job.jobNumber || job.id}_${new Date().toISOString().split('T')[0]}.html`;
    const success = await OneDriveManager.uploadFile(
      oneDriveToken,
      fileName,
      new Blob([jobContent], { type: 'text/html' }),
      'Corpworks/Job Sheets'
    );
    if (success) {
      setUploadStatus({ type: 'success', message: `✓ Job sheet saved to OneDrive: ${fileName}` });
      setJobs(jobs.map(j => j.id === job.id ? { ...j, syncedToOneDrive: true, syncedAt: new Date().toISOString() } : j));
    } else {
      setUploadStatus({ type: 'error', message: '✗ Failed to upload job sheet to OneDrive' });
    }
  };

  const handleDisconnectOneDrive = () => {
    setOneDriveToken(null);
    localStorage.removeItem('onedrive_token');
    localStorage.removeItem('onedrive_refresh');
    setUploadStatus({ type: 'info', message: 'OneDrive disconnected' });
  };

  // ---- Jobs ----------------------------------------------------------------
  const getNextJobNumber = () => {
    const next = (parseInt(localStorage.getItem('corpworks_job_counter') || '0', 10)) + 1;
    localStorage.setItem('corpworks_job_counter', String(next));
    return `JS-${String(next).padStart(3, '0')}`;
  };

  const handleAddJob = () => {
    const newJob = {
      id: Date.now(),
      jobNumber: getNextJobNumber(),
      client: '',
      status: 'New',
      assignedTo: '',
      notes: '',
      createdAt: new Date().toISOString()
    };
    setJobs([...jobs, newJob]);
    setEditingJob(newJob); // open editor immediately so it isn't left blank
  };

  const handleSaveJob = () => {
    if (!editingJob) return;
    setJobs(jobs.map(j => j.id === editingJob.id ? editingJob : j));
    setEditingJob(null);
  };

  const handleDeleteJob = (jobId) => {
    if (!window.confirm('Delete this job? This cannot be undone.')) return;
    setJobs(jobs.filter(j => j.id !== jobId));
    if (editingJob && editingJob.id === jobId) setEditingJob(null);
  };

  // ---- Quotes --------------------------------------------------------------
  const clearQuoteForm = () => setCurrentQuote(emptyQuote());

  const handleAddLineItem = () => {
    setCurrentQuote({
      ...currentQuote,
      lineItems: [...currentQuote.lineItems, { description: '', qty: 1, rate: 0 }]
    });
  };
  const handleUpdateLineItem = (index, field, value) => {
    const updated = currentQuote.lineItems.map((item, i) =>
      i === index
        ? { ...item, [field]: field === 'description' ? value : Math.max(0, parseFloat(value) || 0) }
        : item
    );
    setCurrentQuote({ ...currentQuote, lineItems: updated });
  };
  const handleRemoveLineItem = (index) => {
    setCurrentQuote({
      ...currentQuote,
      lineItems: currentQuote.lineItems.filter((_, i) => i !== index)
    });
  };

  const generateQuoteHTML = (quote) => {
    const t = computeQuoteTotals(quote);
    const itemsHtml = quote.lineItems.map(item =>
      `<tr><td>${esc(item.description)}</td><td>${Number(item.qty) || 0}</td><td>$${(Number(item.rate) || 0).toFixed(2)}</td><td>$${((Number(item.qty) || 0) * (Number(item.rate) || 0)).toFixed(2)}</td></tr>`
    ).join('');

    return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; margin: 30px; color: #333; }
    .header { background: linear-gradient(135deg, #ea580c 0%, #c2410c 100%); color: white; padding: 25px; border-radius: 5px; margin-bottom: 30px; }
    .header h1 { margin: 0; font-size: 28px; }
    .header p { margin: 5px 0 0 0; opacity: 0.9; }
    .info-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 30px; }
    .info-label { font-weight: bold; color: #ea580c; font-size: 12px; text-transform: uppercase; }
    .info-value { font-size: 16px; margin-top: 3px; }
    table { width: 100%; border-collapse: collapse; margin: 30px 0; }
    th { background-color: #f0f0f0; border: 1px solid #ddd; padding: 12px; text-align: left; font-weight: bold; }
    td { border: 1px solid #ddd; padding: 12px; }
    tr:nth-child(even) { background-color: #fafafa; }
    .summary { background: #f9f9f9; border-left: 4px solid #ea580c; padding: 20px; margin-top: 30px; }
    .summary div { margin: 4px 0; }
    .total-row { font-size: 18px; font-weight: bold; color: #ea580c; text-align: right; margin-top: 8px; }
    .footer { margin-top: 40px; padding-top: 20px; border-top: 1px solid #ddd; font-size: 12px; color: #666; }
  </style>
</head>
<body>
  <div class="header">
    <h1>Corpworks Limited</h1>
    <p>Quote #${esc(quote.quoteNum)} | ${esc(quote.date)}</p>
  </div>

  <div class="info-grid">
    <div><div class="info-label">Client</div><div class="info-value">${esc(quote.client)}</div></div>
    <div><div class="info-label">End Client / Site</div><div class="info-value">${esc(quote.endClient)}</div></div>
    <div><div class="info-label">Job Type</div><div class="info-value">${esc(quote.jobType)}</div></div>
    <div><div class="info-label">Prepared By</div><div class="info-value">${esc(quote.preparedBy)}</div></div>
  </div>
  ${quote.notes ? `<p><strong>Notes:</strong> ${esc(quote.notes)}</p>` : ''}

  <h3 style="color: #ea580c;">Line Items</h3>
  <table>
    <tr><th>Description</th><th style="width: 80px;">Qty</th><th style="width: 100px;">Rate</th><th style="width: 100px;">Total</th></tr>
    ${itemsHtml}
  </table>

  <div class="summary">
    <div>Subtotal: $${t.subtotal.toFixed(2)}</div>
    ${t.freight > 0 ? `<div>Freight: $${t.freight.toFixed(2)}</div>` : ''}
    ${t.travel > 0 ? `<div>Travel: $${t.travel.toFixed(2)}</div>` : ''}
    <div>GST (15%): $${t.gst.toFixed(2)}</div>
    <div class="total-row">Total (incl. GST): $${t.total.toFixed(2)}</div>
  </div>
  <div class="footer">
    <p>Generated: ${new Date().toLocaleString()}</p>
    <p>This quote is valid for 30 days from the date above. All prices in NZD.</p>
  </div>
</body>
</html>
    `;
  };

  const generateJobSheetHTML = (job) => {
    return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; margin: 30px; color: #333; }
    .header { background: linear-gradient(135deg, #ea580c 0%, #c2410c 100%); color: white; padding: 25px; border-radius: 5px; margin-bottom: 30px; }
    .header h1 { margin: 0; font-size: 28px; }
    .info-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 30px; }
    .info-label { font-weight: bold; color: #ea580c; font-size: 12px; text-transform: uppercase; }
    .info-value { font-size: 16px; margin-top: 3px; }
    .status { display: inline-block; padding: 6px 12px; border-radius: 20px; font-size: 14px; font-weight: bold; }
    .status.completed { background: #d4edda; color: #155724; }
    .status.in-progress { background: #cce5ff; color: #004085; }
    .status.new { background: #fff3cd; color: #856404; }
  </style>
</head>
<body>
  <div class="header">
    <h1>Corpworks Limited</h1>
    <p>Job Sheet | ${esc(job.jobNumber || job.id)}</p>
  </div>

  <div class="info-grid">
    <div><div class="info-label">Client</div><div class="info-value">${esc(job.client || 'N/A')}</div></div>
    <div><div class="info-label">Status</div><div><span class="status ${esc((job.status || 'New').toLowerCase().replace(' ', '-'))}">${esc(job.status || 'New')}</span></div></div>
    <div><div class="info-label">Assigned To</div><div class="info-value">${esc(job.assignedTo || 'Unassigned')}</div></div>
    <div><div class="info-label">Created</div><div class="info-value">${new Date(job.createdAt).toLocaleDateString()}</div></div>
  </div>
  ${job.notes ? `<div><strong>Notes:</strong> ${esc(job.notes)}</div>` : ''}

  <p style="color: #666; font-size: 12px; margin-top: 40px;">Generated: ${new Date().toLocaleString()}</p>
</body>
</html>
    `;
  };

  // ========================================================================
  // RENDER SECTIONS  (presentation only — RoofPro design system)
  // ========================================================================
  const statusPill = (status) => (
    <span className={`inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold ${
      status === 'Completed' ? 'bg-green-100 text-green-700' :
      status === 'In Progress' ? 'bg-blue-100 text-blue-700' :
      'bg-amber-100 text-amber-700'
    }`}>
      <span className="w-1.5 h-1.5 rounded-full bg-current" />
      {status}
    </span>
  );

  const renderJobModal = () => (
    <div className="fixed inset-0 bg-slate-900/50 flex items-start justify-center z-50 p-4 pt-16 overflow-auto" onClick={() => setEditingJob(null)}>
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg" onClick={(e) => e.stopPropagation()}>
        <div className="flex justify-between items-center border-b border-slate-200 p-5">
          <h3 className="text-lg font-bold text-slate-900">Edit {editingJob.jobNumber}</h3>
          <button onClick={() => setEditingJob(null)} className="text-slate-400 hover:text-slate-700"><X size={22} /></button>
        </div>
        <div className="p-5 space-y-4">
          <div>
            <label className="block text-sm font-semibold mb-1 text-slate-700">Client</label>
            <input
              type="text"
              value={editingJob.client}
              onChange={(e) => setEditingJob({ ...editingJob, client: e.target.value })}
              className="w-full border border-slate-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-orange-500 focus:border-orange-500"
              placeholder="Client name"
            />
          </div>
          <div className="grid grid-cols-2 gap-4">
            <div>
              <label className="block text-sm font-semibold mb-1 text-slate-700">Status</label>
              <select
                value={editingJob.status}
                onChange={(e) => setEditingJob({ ...editingJob, status: e.target.value })}
                className="w-full border border-slate-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-orange-500 focus:border-orange-500"
              >
                <option>New</option>
                <option>In Progress</option>
                <option>Completed</option>
              </select>
            </div>
            <div>
              <label className="block text-sm font-semibold mb-1 text-slate-700">Assigned To</label>
              <input
                type="text"
                value={editingJob.assignedTo}
                onChange={(e) => setEditingJob({ ...editingJob, assignedTo: e.target.value })}
                className="w-full border border-slate-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-orange-500 focus:border-orange-500"
                placeholder="Technician"
              />
            </div>
          </div>
          <div>
            <label className="block text-sm font-semibold mb-1 text-slate-700">Notes</label>
            <textarea
              value={editingJob.notes}
              onChange={(e) => setEditingJob({ ...editingJob, notes: e.target.value })}
              rows={3}
              className="w-full border border-slate-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-orange-500 focus:border-orange-500"
            />
          </div>
        </div>
        <div className="flex justify-between items-center border-t border-slate-200 p-5">
          <button
            onClick={() => handleDeleteJob(editingJob.id)}
            className="text-red-600 hover:text-red-800 font-semibold flex items-center gap-1"
          >
            <Trash2 size={16} /> Delete
          </button>
          <div className="flex gap-2">
            <button onClick={() => setEditingJob(null)} className="px-4 py-2 rounded-lg bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 font-semibold">Cancel</button>
            <button onClick={handleSaveJob} className="px-4 py-2 rounded-lg bg-orange-600 text-white hover:bg-orange-700 font-semibold">Save</button>
          </div>
        </div>
      </div>
    </div>
  );

  const renderKanban = () => {
    const columns = ['New', 'In Progress', 'Completed'];
    return (
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        {columns.map(col => (
          <div key={col} className="bg-slate-100 rounded-xl p-3">
            <h3 className="font-semibold text-slate-700 mb-3 flex items-center justify-between text-sm">
              {col}
              <span className="text-xs bg-white rounded-full px-2 py-0.5 text-slate-500">{jobs.filter(j => j.status === col).length}</span>
            </h3>
            <div className="space-y-2 min-h-[60px]">
              {jobs.filter(j => j.status === col).map(job => (
                <div
                  key={job.id}
                  onClick={() => setEditingJob(job)}
                  className="bg-white rounded-lg p-3 border border-slate-200 shadow-sm cursor-pointer hover:shadow-md hover:border-orange-200 transition"
                >
                  <p className="font-mono text-xs text-slate-400">{job.jobNumber}</p>
                  <p className="font-semibold text-slate-900">{job.client || '—'}</p>
                  <p className="text-sm text-slate-500">{job.assignedTo || 'Unassigned'}</p>
                </div>
              ))}
            </div>
          </div>
        ))}
      </div>
    );
  };

  const renderDashboard = () => {
    const count = (s) => jobs.filter(j => j.status === s).length;
    const active = jobs.filter(j => j.status !== 'Completed');
    const stats = [
      { label: 'Total jobs', value: jobs.length, hint: `${count('New')} new` },
      { label: 'In progress', value: count('In Progress'), hint: 'currently on site' },
      { label: 'Completed', value: count('Completed'), hint: 'all time' },
      { label: 'Quotes saved', value: quotes.length, hint: 'to OneDrive' },
    ];
    return (
      <div className="space-y-5">
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
          {stats.map(s => (
            <div key={s.label} className="bg-white border border-slate-200 rounded-xl p-5">
              <div className="text-sm text-slate-500 font-medium">{s.label}</div>
              <div className="text-3xl font-bold text-slate-900 mt-1">{s.value}</div>
              <div className="text-xs text-slate-400 mt-1">{s.hint}</div>
            </div>
          ))}
        </div>

        <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
          <div className="flex items-center justify-between px-5 py-4 border-b border-slate-200">
            <h3 className="font-semibold text-slate-900">Active jobs</h3>
            <button onClick={() => setActiveTab('tracker')} className="text-sm font-semibold text-white bg-orange-600 hover:bg-orange-700 px-3 py-1.5 rounded-lg">View tracker</button>
          </div>
          {active.length === 0 ? (
            <div className="text-center py-12 text-slate-400">
              <LayoutDashboard size={40} className="mx-auto mb-3 opacity-30" />
              <p className="font-semibold">No active jobs</p>
              <p className="text-sm">Create one from the Job Tracker</p>
            </div>
          ) : (
            <table className="w-full">
              <thead>
                <tr className="bg-slate-50">
                  <th className="text-left text-xs uppercase tracking-wide text-slate-500 font-semibold px-5 py-3">Job #</th>
                  <th className="text-left text-xs uppercase tracking-wide text-slate-500 font-semibold px-5 py-3">Client</th>
                  <th className="text-left text-xs uppercase tracking-wide text-slate-500 font-semibold px-5 py-3">Status</th>
                  <th className="text-left text-xs uppercase tracking-wide text-slate-500 font-semibold px-5 py-3">Assigned</th>
                </tr>
              </thead>
              <tbody>
                {active.map(job => (
                  <tr key={job.id} className="border-t border-slate-100 hover:bg-orange-50/40 cursor-pointer" onClick={() => setEditingJob(job)}>
                    <td className="px-5 py-3 font-mono text-sm text-slate-500">{job.jobNumber}</td>
                    <td className="px-5 py-3 font-medium">{job.client || '—'}</td>
                    <td className="px-5 py-3">{statusPill(job.status)}</td>
                    <td className="px-5 py-3 text-slate-500">{job.assignedTo || '—'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      </div>
    );
  };

  const renderJobTracker = () => (
    <div className="bg-white border border-slate-200 rounded-xl p-6">
      <div className="flex justify-between items-center mb-6">
        <h2 className="text-xl font-bold text-slate-900">Job Tracker</h2>
        <button onClick={handleAddJob} className="bg-orange-600 text-white px-4 py-2 rounded-lg hover:bg-orange-700 font-semibold flex items-center gap-1">
          <Plus size={18} /> New Job
        </button>
      </div>
      <div className="mb-5 flex gap-1 border-b border-slate-200">
        <button
          onClick={() => setViewMode('list')}
          className={`px-4 py-2 font-semibold text-sm transition border-b-2 -mb-px ${viewMode === 'list' ? 'text-orange-600 border-orange-600' : 'text-slate-500 border-transparent hover:text-slate-700'}`}
        >
          List View
        </button>
        <button
          onClick={() => setViewMode('kanban')}
          className={`px-4 py-2 font-semibold text-sm transition border-b-2 -mb-px ${viewMode === 'kanban' ? 'text-orange-600 border-orange-600' : 'text-slate-500 border-transparent hover:text-slate-700'}`}
        >
          Kanban
        </button>
      </div>

      {jobs.length === 0 ? (
        <div className="text-center py-12 text-slate-400">
          <FileText size={48} className="mx-auto mb-4 opacity-30" />
          <p className="font-semibold">No jobs yet</p>
          <p className="text-sm">Create one to get started</p>
        </div>
      ) : viewMode === 'list' ? (
        <div className="overflow-x-auto border border-slate-200 rounded-xl">
          <table className="w-full">
            <thead>
              <tr className="bg-slate-50">
                <th className="px-4 py-3 text-left text-xs uppercase tracking-wide font-semibold text-slate-500">Job #</th>
                <th className="px-4 py-3 text-left text-xs uppercase tracking-wide font-semibold text-slate-500">Client</th>
                <th className="px-4 py-3 text-left text-xs uppercase tracking-wide font-semibold text-slate-500">Status</th>
                <th className="px-4 py-3 text-left text-xs uppercase tracking-wide font-semibold text-slate-500">Assigned</th>
                <th className="px-4 py-3 text-left text-xs uppercase tracking-wide font-semibold text-slate-500">Actions</th>
              </tr>
            </thead>
            <tbody>
              {jobs.map(job => (
                <tr key={job.id} className="border-t border-slate-100 hover:bg-orange-50/40 cursor-pointer" onClick={() => setEditingJob(job)}>
                  <td className="px-4 py-3 font-mono text-sm text-slate-500">{job.jobNumber}</td>
                  <td className="px-4 py-3 font-medium">{job.client || '—'}</td>
                  <td className="px-4 py-3">{statusPill(job.status)}</td>
                  <td className="px-4 py-3 text-slate-500">{job.assignedTo || '—'}</td>
                  <td className="px-4 py-3">
                    <div className="flex gap-3" onClick={(e) => e.stopPropagation()}>
                      <button onClick={() => setEditingJob(job)} className="text-slate-600 hover:text-orange-600 text-sm font-semibold flex items-center gap-1">
                        <Pencil size={16} /> Edit
                      </button>
                      <button onClick={() => handleSaveJobSheetToOneDrive(job)} className="text-slate-600 hover:text-orange-600 text-sm font-semibold flex items-center gap-1 disabled:opacity-40" disabled={!oneDriveToken}>
                        <Cloud size={16} /> Save
                      </button>
                      <button onClick={() => handleDeleteJob(job.id)} className="text-red-600 hover:text-red-800 text-sm font-semibold flex items-center gap-1">
                        <Trash2 size={16} /> Delete
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : (
        renderKanban()
      )}
    </div>
  );

  const renderQuoteBuilder = () => {
    const t = computeQuoteTotals(currentQuote);
    return (
      <div className="bg-white border border-slate-200 rounded-xl p-6">
        <h2 className="text-xl font-bold mb-6 text-slate-900">Quote Builder</h2>
        <div className="grid grid-cols-2 gap-4 mb-5">
          <div>
            <label className="block text-sm font-semibold mb-2 text-slate-700">Quote Number</label>
            <input type="text" value={currentQuote.quoteNum} onChange={(e) => setCurrentQuote({ ...currentQuote, quoteNum: e.target.value })} className="w-full border border-slate-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-orange-500 focus:border-orange-500" placeholder="QT-001" />
          </div>
          <div>
            <label className="block text-sm font-semibold mb-2 text-slate-700">Date</label>
            <input type="date" value={currentQuote.date} onChange={(e) => setCurrentQuote({ ...currentQuote, date: e.target.value })} className="w-full border border-slate-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-orange-500 focus:border-orange-500" />
          </div>
        </div>
        <div className="grid grid-cols-2 gap-4 mb-5">
          <div>
            <label className="block text-sm font-semibold mb-2 text-slate-700">Client</label>
            <input type="text" value={currentQuote.client} onChange={(e) => setCurrentQuote({ ...currentQuote, client: e.target.value })} className="w-full border border-slate-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-orange-500 focus:border-orange-500" />
          </div>
          <div>
            <label className="block text-sm font-semibold mb-2 text-slate-700">End Client / Site</label>
            <input type="text" value={currentQuote.endClient} onChange={(e) => setCurrentQuote({ ...currentQuote, endClient: e.target.value })} className="w-full border border-slate-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-orange-500 focus:border-orange-500" />
          </div>
        </div>

        <div className="mb-5">
          <label className="block text-sm font-semibold mb-3 text-slate-700">Line Items</label>
          <div className="overflow-x-auto mb-4 border border-slate-200 rounded-xl">
            <table className="w-full">
              <thead>
                <tr className="bg-slate-50">
                  <th className="p-3 text-left font-semibold text-xs uppercase tracking-wide text-slate-500">Description</th>
                  <th className="p-3 text-left font-semibold text-xs uppercase tracking-wide text-slate-500 w-20">Qty</th>
                  <th className="p-3 text-left font-semibold text-xs uppercase tracking-wide text-slate-500 w-24">Rate $</th>
                  <th className="p-3 text-left font-semibold text-xs uppercase tracking-wide text-slate-500 w-24">Total</th>
                  <th className="p-3 text-left font-semibold text-xs uppercase tracking-wide text-slate-500 w-16">Action</th>
                </tr>
              </thead>
              <tbody>
                {currentQuote.lineItems.map((item, idx) => (
                  <tr key={idx} className="border-t border-slate-100">
                    <td className="p-3">
                      <input type="text" value={item.description} onChange={(e) => handleUpdateLineItem(idx, 'description', e.target.value)} className="w-full border border-slate-300 rounded px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-orange-500 focus:border-orange-500" />
                    </td>
                    <td className="p-3">
                      <input type="number" min="0" value={item.qty} onChange={(e) => handleUpdateLineItem(idx, 'qty', e.target.value)} className="w-full border border-slate-300 rounded px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-orange-500 focus:border-orange-500" />
                    </td>
                    <td className="p-3">
                      <input type="number" min="0" value={item.rate} onChange={(e) => handleUpdateLineItem(idx, 'rate', e.target.value)} className="w-full border border-slate-300 rounded px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-orange-500 focus:border-orange-500" />
                    </td>
                    <td className="p-3 text-sm text-right font-semibold">${((Number(item.qty) || 0) * (Number(item.rate) || 0)).toFixed(2)}</td>
                    <td className="p-3 text-center">
                      <button onClick={() => handleRemoveLineItem(idx)} className="text-red-600 hover:text-red-800 font-semibold text-sm">Remove</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <button onClick={handleAddLineItem} className="bg-white border border-slate-300 text-slate-700 px-4 py-2 rounded-lg hover:bg-slate-50 font-semibold text-sm">+ Add Line Item</button>
        </div>

        <div className="grid grid-cols-3 gap-4 mb-5">
          <div>
            <label className="block text-sm font-semibold mb-2 text-slate-700">Freight $</label>
            <input type="number" min="0" value={currentQuote.freight} onChange={(e) => setCurrentQuote({ ...currentQuote, freight: Math.max(0, parseFloat(e.target.value) || 0) })} className="w-full border border-slate-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-orange-500 focus:border-orange-500" />
          </div>
          <div>
            <label className="block text-sm font-semibold mb-2 text-slate-700">Travel Hours</label>
            <input type="number" min="0" value={currentQuote.travelHours} onChange={(e) => setCurrentQuote({ ...currentQuote, travelHours: Math.max(0, parseFloat(e.target.value) || 0) })} className="w-full border border-slate-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-orange-500 focus:border-orange-500" />
          </div>
          <div>
            <label className="block text-sm font-semibold mb-2 text-slate-700">Travel Rate $/hr</label>
            <input type="number" min="0" value={currentQuote.travelRate} onChange={(e) => setCurrentQuote({ ...currentQuote, travelRate: Math.max(0, parseFloat(e.target.value) || 0) })} className="w-full border border-slate-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-orange-500 focus:border-orange-500" />
          </div>
        </div>

        <div className="bg-orange-50 border border-orange-200 p-6 rounded-xl mb-6">
          <div className="text-right space-y-1">
            <p className="text-sm text-slate-600">Subtotal: <span className="font-bold">${t.subtotal.toFixed(2)}</span></p>
            {t.freight > 0 && <p className="text-sm text-slate-600">Freight: <span className="font-bold">${t.freight.toFixed(2)}</span></p>}
            {t.travel > 0 && <p className="text-sm text-slate-600">Travel: <span className="font-bold">${t.travel.toFixed(2)}</span></p>}
            <p className="text-sm text-slate-600">GST (15%): <span className="font-bold">${t.gst.toFixed(2)}</span></p>
            <p className="text-2xl font-bold text-orange-600">Total (incl. GST): ${t.total.toFixed(2)}</p>
          </div>
        </div>

        <div className="flex gap-3">
          <button onClick={handleSaveQuoteToOneDrive} disabled={!oneDriveToken} className="bg-orange-600 text-white px-6 py-3 rounded-lg hover:bg-orange-700 font-semibold flex items-center gap-2 disabled:opacity-50 disabled:cursor-not-allowed">
            <Cloud size={18} /> Save to OneDrive
          </button>
          <button onClick={clearQuoteForm} className="bg-white border border-slate-300 text-slate-700 px-6 py-3 rounded-lg hover:bg-slate-50 font-semibold">Clear</button>
        </div>
      </div>
    );
  };

  // Compact OneDrive banner shown at the top of the content area
  const renderOneDriveStatus = () => (
    <div className="bg-white border border-slate-200 rounded-xl p-4 mb-5">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-4">
          <div className={`p-3 rounded-full ${oneDriveToken ? 'bg-green-100' : 'bg-slate-100'}`}>
            <Cloud size={22} className={oneDriveToken ? 'text-green-600' : 'text-slate-500'} />
          </div>
          <div>
            <p className="font-bold text-slate-900">OneDrive Integration</p>
            <p className="text-sm text-slate-500">
              {oneDriveToken ? '✓ Connected and ready to sync' : 'Connect to enable automatic file uploads'}
            </p>
          </div>
        </div>
        {!oneDriveToken ? (
          <button onClick={initiateOneDriveAuth} className="bg-orange-600 text-white px-5 py-2.5 rounded-lg hover:bg-orange-700 font-semibold flex items-center gap-2 whitespace-nowrap">
            <Cloud size={18} /> Connect OneDrive
          </button>
        ) : (
          <button onClick={handleDisconnectOneDrive} className="bg-white border border-slate-300 text-slate-700 px-5 py-2.5 rounded-lg hover:bg-slate-50 font-semibold flex items-center gap-2 whitespace-nowrap">
            <LogOut size={18} /> Disconnect
          </button>
        )}
      </div>
      {uploadStatus && (
        <div className={`mt-4 p-4 rounded-lg flex items-center gap-3 ${
          uploadStatus.type === 'success' ? 'bg-green-100 text-green-800' :
          uploadStatus.type === 'error' ? 'bg-red-100 text-red-800' :
          uploadStatus.type === 'loading' ? 'bg-blue-100 text-blue-800' :
          'bg-slate-100 text-slate-800'
        }`}>
          {uploadStatus.type === 'loading' && <Loader size={18} className="animate-spin flex-shrink-0" />}
          {uploadStatus.type === 'success' && <Check size={18} className="flex-shrink-0" />}
          {uploadStatus.type === 'error' && <AlertCircle size={18} className="flex-shrink-0" />}
          <span className="font-semibold">{uploadStatus.message}</span>
        </div>
      )}
    </div>
  );

  // ---- Layout --------------------------------------------------------------
  const navItems = [
    { id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
    { id: 'tracker', label: 'Job Tracker', icon: FileText },
    { id: 'quoting', label: 'Quote Builder', icon: DollarSign },
    { id: 'calendar', label: 'Calendar', icon: Calendar },
    { id: 'van', label: 'Van View', icon: Truck },
  ];
  const titles = {
    dashboard: ['Dashboard', 'Your business at a glance'],
    tracker: ['Job Tracker', 'All jobs'],
    quoting: ['Quote Builder', 'Build and save quotes'],
    calendar: ['Calendar', 'Scheduling'],
    van: ['Van View', 'Van assignments'],
  };

  return (
    <div className="min-h-screen flex bg-slate-50 text-slate-900">
      {/* Sidebar */}
      <aside className="w-60 bg-slate-900 text-slate-300 flex flex-col fixed inset-y-0 left-0">
        <div className="flex items-center gap-3 px-5 py-5 border-b border-slate-800">
          <div className="w-9 h-9 rounded-lg bg-orange-600 text-white flex items-center justify-center font-extrabold text-lg">C</div>
          <div>
            <div className="text-white font-bold leading-tight">Corpworks JMS</div>
            <div className="text-xs text-slate-500">Job Management</div>
          </div>
        </div>
        <nav className="flex-1 p-3 space-y-1 overflow-auto">
          {navItems.map(n => {
            const Icon = n.icon;
            return (
              <button
                key={n.id}
                onClick={() => setActiveTab(n.id)}
                className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition ${
                  activeTab === n.id ? 'bg-orange-600 text-white' : 'text-slate-300 hover:bg-slate-800 hover:text-white'
                }`}
              >
                <Icon size={18} /> {n.label}
              </button>
            );
          })}
        </nav>
        <div className="px-5 py-4 border-t border-slate-800 text-xs text-slate-500 flex items-center gap-2">
          <Cloud size={14} className={oneDriveToken ? 'text-green-400' : 'text-slate-500'} />
          {oneDriveToken ? 'OneDrive connected' : 'OneDrive offline'}
        </div>
      </aside>

      {/* Main */}
      <div className="flex-1 ml-60 min-w-0">
        <header className="bg-white border-b border-slate-200 px-7 py-4 flex items-center justify-between sticky top-0 z-10">
          <div>
            <h1 className="text-xl font-bold text-slate-900">{titles[activeTab][0]}</h1>
            <p className="text-sm text-slate-500">{titles[activeTab][1]}</p>
          </div>
          <div className="flex items-center gap-4 text-sm">
            {oneDriveToken && (
              <div className="flex items-center gap-2 bg-green-50 text-green-700 px-3 py-1.5 rounded-lg font-medium">
                <Cloud size={16} /> OneDrive Ready
              </div>
            )}
            <span className="text-slate-500">{new Date().toLocaleDateString()}</span>
          </div>
        </header>

        <main className="px-7 py-6 max-w-6xl">
          {renderOneDriveStatus()}

          {activeTab === 'dashboard' && renderDashboard()}
          {activeTab === 'tracker' && renderJobTracker()}
          {activeTab === 'quoting' && renderQuoteBuilder()}
          {activeTab === 'calendar' && (
            <div className="bg-white border border-slate-200 rounded-xl p-10 text-center">
              <Calendar size={48} className="mx-auto mb-4 text-slate-300" />
              <h2 className="text-xl font-bold text-slate-900 mb-2">Calendar View</h2>
              <p className="text-slate-500">Calendar integration coming soon...</p>
            </div>
          )}
          {activeTab === 'van' && (
            <div className="bg-white border border-slate-200 rounded-xl p-10 text-center">
              <Truck size={48} className="mx-auto mb-4 text-slate-300" />
              <h2 className="text-xl font-bold text-slate-900 mb-2">Van View</h2>
              <p className="text-slate-500">Van assignments and tracking coming soon...</p>
            </div>
          )}
        </main>
      </div>

      {editingJob && renderJobModal()}
    </div>
  );
}
