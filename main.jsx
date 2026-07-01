import React, { useState, useEffect } from 'react';
import { Calendar, FileText, DollarSign, Truck, Cloud, LogOut, Download, Check, AlertCircle, Loader, X } from 'lucide-react';

// ============================================================================
// ONEDRIVE INTEGRATION MODULE
// ============================================================================

const OneDriveManager = {
  clientId: 'b9c9a198-ecad-4041-9f15-ce32f34c0567',
  redirectUri: 'https://corpworks-jms.netlify.app/auth/callback',
  scopes: ['Files.ReadWrite', 'User.Read'],
  
  getAuthUrl() {
    const params = new URLSearchParams({
      client_id: this.clientId,
      redirect_uri: this.redirectUri,
      response_type: 'code',
      scope: this.scopes.join(' '),
      response_mode: 'query'
    });
    return `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?${params}`;
  },

  async exchangeCodeForToken(code) {
    try {
      const response = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id: this.clientId,
          client_secret: 'a_n8Q~2q3RfnVsc2Z9qcs5rxoZ6FZ8Fj8JMuddek',
          code: code,
          redirect_uri: this.redirectUri,
          grant_type: 'authorization_code'
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
  const [activeTab, setActiveTab] = useState('tracker');
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
  const [currentQuote, setCurrentQuote] = useState({
    quoteNum: '',
    client: '',
    endClient: '',
    date: new Date().toISOString().split('T')[0],
    preparedBy: 'Monique',
    jobType: 'Install',
    notes: '',
    lineItems: [],
    travelHours: 0,
    freight: 0
  });

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
      setUploadStatus({ type: 'success', message: '✓ OneDrive connected!' });
      window.history.replaceState({}, document.title, window.location.pathname);
    } else {
      setUploadStatus({ type: 'error', message: '✗ Failed to connect OneDrive' });
    }
  };

  const initiateOneDriveAuth = () => {
    window.location.href = OneDriveManager.getAuthUrl();
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
    const fileName = `JobSheet_${job.id}_${new Date().toISOString().split('T')[0]}.html`;

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
    setUploadStatus({ type: 'info', message: 'OneDrive disconnected' });
  };

  const clearQuoteForm = () => {
    setCurrentQuote({
      quoteNum: '',
      client: '',
      endClient: '',
      date: new Date().toISOString().split('T')[0],
      preparedBy: 'Monique',
      jobType: 'Install',
      notes: '',
      lineItems: [],
      travelHours: 0,
      freight: 0
    });
  };

  const handleAddLineItem = () => {
    setCurrentQuote({
      ...currentQuote,
      lineItems: [...currentQuote.lineItems, { description: '', qty: 1, rate: 0 }]
    });
  };

  const handleUpdateLineItem = (index, field, value) => {
    const updated = [...currentQuote.lineItems];
    updated[index][field] = field === 'description' ? value : parseFloat(value) || 0;
    setCurrentQuote({ ...currentQuote, lineItems: updated });
  };

  const handleRemoveLineItem = (index) => {
    setCurrentQuote({
      ...currentQuote,
      lineItems: currentQuote.lineItems.filter((_, i) => i !== index)
    });
  };

  const calculateQuoteTotal = () => {
    const itemsTotal = currentQuote.lineItems.reduce((sum, item) => sum + (item.qty * item.rate), 0);
    return itemsTotal + (currentQuote.freight || 0);
  };

  const generateQuoteHTML = (quote) => {
    const itemsHtml = quote.lineItems.map(item => 
      `<tr><td>${item.description}</td><td>${item.qty}</td><td>$${item.rate.toFixed(2)}</td><td>$${(item.qty * item.rate).toFixed(2)}</td></tr>`
    ).join('');
    
    const total = calculateQuoteTotal();
    
    return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; margin: 30px; color: #333; }
    .header { background: linear-gradient(135deg, #0078d4 0%, #0063b1 100%); color: white; padding: 25px; border-radius: 5px; margin-bottom: 30px; }
    .header h1 { margin: 0; font-size: 28px; }
    .header p { margin: 5px 0 0 0; opacity: 0.9; }
    .info-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 30px; }
    .info-item { }
    .info-label { font-weight: bold; color: #0078d4; font-size: 12px; text-transform: uppercase; }
    .info-value { font-size: 16px; margin-top: 3px; }
    table { width: 100%; border-collapse: collapse; margin: 30px 0; }
    th { background-color: #f0f0f0; border: 1px solid #ddd; padding: 12px; text-align: left; font-weight: bold; }
    td { border: 1px solid #ddd; padding: 12px; }
    tr:nth-child(even) { background-color: #fafafa; }
    .summary { background: #f9f9f9; border-left: 4px solid #0078d4; padding: 20px; margin-top: 30px; }
    .total-row { font-size: 18px; font-weight: bold; color: #0078d4; text-align: right; }
    .footer { margin-top: 40px; padding-top: 20px; border-top: 1px solid #ddd; font-size: 12px; color: #666; }
  </style>
</head>
<body>
  <div class="header">
    <h1>Corpworks Limited</h1>
    <p>Quote #${quote.quoteNum} | ${quote.date}</p>
  </div>
  
  <div class="info-grid">
    <div class="info-item">
      <div class="info-label">Client</div>
      <div class="info-value">${quote.client}</div>
    </div>
    <div class="info-item">
      <div class="info-label">End Client / Site</div>
      <div class="info-value">${quote.endClient}</div>
    </div>
    <div class="info-item">
      <div class="info-label">Job Type</div>
      <div class="info-value">${quote.jobType}</div>
    </div>
    <div class="info-item">
      <div class="info-label">Prepared By</div>
      <div class="info-value">${quote.preparedBy}</div>
    </div>
  </div>

  ${quote.notes ? `<p><strong>Notes:</strong> ${quote.notes}</p>` : ''}
  
  <h3 style="color: #0078d4;">Line Items</h3>
  <table>
    <tr>
      <th>Description</th>
      <th style="width: 80px;">Qty</th>
      <th style="width: 100px;">Rate</th>
      <th style="width: 100px;">Total</th>
    </tr>
    ${itemsHtml}
  </table>
  
  <div class="summary">
    ${quote.lineItems.length > 0 ? `<div>Subtotal: $${quote.lineItems.reduce((sum, i) => sum + (i.qty * i.rate), 0).toFixed(2)}</div>` : ''}
    ${quote.freight > 0 ? `<div>Freight: $${quote.freight.toFixed(2)}</div>` : ''}
    <div class="total-row">Total: $${total.toFixed(2)}</div>
  </div>

  <div class="footer">
    <p>Generated: ${new Date().toLocaleString()}</p>
    <p>This quote is valid for 30 days from the date above.</p>
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
    .header { background: linear-gradient(135deg, #0078d4 0%, #0063b1 100%); color: white; padding: 25px; border-radius: 5px; margin-bottom: 30px; }
    .header h1 { margin: 0; font-size: 28px; }
    .info-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 30px; }
    .info-label { font-weight: bold; color: #0078d4; font-size: 12px; text-transform: uppercase; }
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
    <p>Job Sheet | Job #${job.id}</p>
  </div>
  
  <div class="info-grid">
    <div>
      <div class="info-label">Client</div>
      <div class="info-value">${job.client || 'N/A'}</div>
    </div>
    <div>
      <div class="info-label">Status</div>
      <div><span class="status ${job.status?.toLowerCase().replace(' ', '-')}">${job.status || 'New'}</span></div>
    </div>
    <div>
      <div class="info-label">Assigned To</div>
      <div class="info-value">${job.assignedTo || 'Unassigned'}</div>
    </div>
    <div>
      <div class="info-label">Created</div>
      <div class="info-value">${new Date(job.createdAt).toLocaleDateString()}</div>
    </div>
  </div>

  ${job.notes ? `<div><strong>Notes:</strong> ${job.notes}</div>` : ''}
  
  <p style="color: #666; font-size: 12px; margin-top: 40px;">Generated: ${new Date().toLocaleString()}</p>
</body>
</html>
    `;
  };

  // ========================================================================
  // RENDER SECTIONS
  // ========================================================================

  const renderJobTracker = () => (
    <div className="bg-white rounded-lg p-6 shadow-sm">
      <div className="flex justify-between items-center mb-6">
        <h2 className="text-2xl font-bold text-gray-900">Job Tracker</h2>
        <button
          onClick={() => {
            const jobNumber = `JS-${String(jobs.length + 1).padStart(3, '0')}`;
            const newJob = {
              id: Date.now(),
              jobNumber,
              client: '',
              status: 'New',
              assignedTo: '',
              notes: '',
              createdAt: new Date().toISOString()
            };
            setJobs([...jobs, newJob]);
          }}
          className="bg-blue-600 text-white px-4 py-2 rounded-lg hover:bg-blue-700 font-semibold"
        >
          + New Job
        </button>
      </div>

      <div className="mb-4 flex gap-2">
        <button
          onClick={() => setViewMode('list')}
          className={`px-4 py-2 rounded-lg font-semibold transition ${
            viewMode === 'list' ? 'bg-blue-600 text-white' : 'bg-gray-200 text-gray-700 hover:bg-gray-300'
          }`}
        >
          📋 List View
        </button>
        <button
          onClick={() => setViewMode('kanban')}
          className={`px-4 py-2 rounded-lg font-semibold transition ${
            viewMode === 'kanban' ? 'bg-blue-600 text-white' : 'bg-gray-200 text-gray-700 hover:bg-gray-300'
          }`}
        >
          🎯 Kanban
        </button>
      </div>

      {viewMode === 'list' && (
        <div className="overflow-x-auto">
          <table className="w-full">
            <thead>
              <tr className="bg-gray-100 border-b">
                <th className="p-3 text-left font-semibold">Job #</th>
                <th className="p-3 text-left font-semibold">Client</th>
                <th className="p-3 text-left font-semibold">Status</th>
                <th className="p-3 text-left font-semibold">Assigned</th>
                <th className="p-3 text-left font-semibold">Actions</th>
              </tr>
            </thead>
            <tbody>
              {jobs.map(job => (
                <tr key={job.id} className="border-b hover:bg-gray-50">
                  <td className="p-3 font-mono text-sm">{job.jobNumber}</td>
                  <td className="p-3">{job.client || '—'}</td>
                  <td className="p-3">
                    <span className={`px-3 py-1 rounded-full text-sm font-semibold ${
                      job.status === 'Completed' ? 'bg-green-100 text-green-800' :
                      job.status === 'In Progress' ? 'bg-blue-100 text-blue-800' :
                      'bg-yellow-100 text-yellow-800'
                    }`}>
                      {job.status}
                    </span>
                  </td>
                  <td className="p-3">{job.assignedTo || '—'}</td>
                  <td className="p-3">
                    <button
                      onClick={() => handleSaveJobSheetToOneDrive(job)}
                      className="text-blue-600 hover:text-blue-800 text-sm font-semibold flex items-center gap-1"
                      disabled={!oneDriveToken}
                    >
                      <Cloud size={16} /> Save
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {jobs.length === 0 && (
        <div className="text-center py-12 text-gray-500">
          <FileText size={48} className="mx-auto mb-4 opacity-30" />
          <p className="font-semibold">No jobs yet</p>
          <p className="text-sm">Create one to get started</p>
        </div>
      )}
    </div>
  );

  const renderQuoteBuilder = () => (
    <div className="bg-white rounded-lg p-6 shadow-sm">
      <h2 className="text-2xl font-bold mb-6 text-gray-900">Quote Builder</h2>

      <div className="grid grid-cols-2 gap-4 mb-6">
        <div>
          <label className="block text-sm font-bold mb-2 text-gray-700">Quote Number</label>
          <input
            type="text"
            value={currentQuote.quoteNum}
            onChange={(e) => setCurrentQuote({ ...currentQuote, quoteNum: e.target.value })}
            className="w-full border border-gray-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-blue-500"
            placeholder="QT-001"
          />
        </div>
        <div>
          <label className="block text-sm font-bold mb-2 text-gray-700">Date</label>
          <input
            type="date"
            value={currentQuote.date}
            onChange={(e) => setCurrentQuote({ ...currentQuote, date: e.target.value })}
            className="w-full border border-gray-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-blue-500"
          />
        </div>
      </div>

      <div className="grid grid-cols-2 gap-4 mb-6">
        <div>
          <label className="block text-sm font-bold mb-2 text-gray-700">Client</label>
          <input
            type="text"
            value={currentQuote.client}
            onChange={(e) => setCurrentQuote({ ...currentQuote, client: e.target.value })}
            className="w-full border border-gray-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-blue-500"
          />
        </div>
        <div>
          <label className="block text-sm font-bold mb-2 text-gray-700">End Client / Site</label>
          <input
            type="text"
            value={currentQuote.endClient}
            onChange={(e) => setCurrentQuote({ ...currentQuote, endClient: e.target.value })}
            className="w-full border border-gray-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-blue-500"
          />
        </div>
      </div>

      <div className="mb-6">
        <label className="block text-sm font-bold mb-3 text-gray-700">Line Items</label>
        <div className="overflow-x-auto mb-4 border border-gray-200 rounded-lg">
          <table className="w-full">
            <thead>
              <tr className="bg-gray-100">
                <th className="border-r p-3 text-left font-semibold text-sm">Description</th>
                <th className="border-r p-3 text-left font-semibold text-sm w-20">Qty</th>
                <th className="border-r p-3 text-left font-semibold text-sm w-24">Rate $</th>
                <th className="border-r p-3 text-left font-semibold text-sm w-24">Total</th>
                <th className="p-3 text-left font-semibold text-sm w-16">Action</th>
              </tr>
            </thead>
            <tbody>
              {currentQuote.lineItems.map((item, idx) => (
                <tr key={idx} className="border-t">
                  <td className="border-r p-3">
                    <input
                      type="text"
                      value={item.description}
                      onChange={(e) => handleUpdateLineItem(idx, 'description', e.target.value)}
                      className="w-full border border-gray-300 rounded px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                    />
                  </td>
                  <td className="border-r p-3">
                    <input
                      type="number"
                      value={item.qty}
                      onChange={(e) => handleUpdateLineItem(idx, 'qty', e.target.value)}
                      className="w-full border border-gray-300 rounded px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                    />
                  </td>
                  <td className="border-r p-3">
                    <input
                      type="number"
                      value={item.rate}
                      onChange={(e) => handleUpdateLineItem(idx, 'rate', e.target.value)}
                      className="w-full border border-gray-300 rounded px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                    />
                  </td>
                  <td className="border-r p-3 text-sm text-right font-semibold">${(item.qty * item.rate).toFixed(2)}</td>
                  <td className="p-3 text-center">
                    <button
                      onClick={() => handleRemoveLineItem(idx)}
                      className="text-red-600 hover:text-red-800 font-semibold text-sm"
                    >
                      Remove
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <button
          onClick={handleAddLineItem}
          className="bg-gray-600 text-white px-4 py-2 rounded-lg hover:bg-gray-700 font-semibold text-sm"
        >
          + Add Line Item
        </button>
      </div>

      <div className="bg-blue-50 border border-blue-200 p-6 rounded-lg mb-6">
        <div className="text-right space-y-2">
          <p className="text-sm text-gray-600">Subtotal: <span className="font-bold">${currentQuote.lineItems.reduce((sum, i) => sum + (i.qty * i.rate), 0).toFixed(2)}</span></p>
          {currentQuote.freight > 0 && <p className="text-sm text-gray-600">Freight: <span className="font-bold">${currentQuote.freight.toFixed(2)}</span></p>}
          <p className="text-2xl font-bold text-blue-600">Total: ${calculateQuoteTotal().toFixed(2)}</p>
        </div>
      </div>

      <div className="flex gap-3">
        <button
          onClick={handleSaveQuoteToOneDrive}
          disabled={!oneDriveToken}
          className="bg-blue-600 text-white px-6 py-3 rounded-lg hover:bg-blue-700 font-semibold flex items-center gap-2 disabled:opacity-50 disabled:cursor-not-allowed"
        >
          <Cloud size={18} /> Save to OneDrive
        </button>
        <button
          onClick={clearQuoteForm}
          className="bg-gray-600 text-white px-6 py-3 rounded-lg hover:bg-gray-700 font-semibold"
        >
          Clear
        </button>
      </div>
    </div>
  );

  const renderOneDriveStatus = () => (
    <div className="bg-gradient-to-r from-blue-50 to-blue-100 border border-blue-200 rounded-lg p-5 mb-6">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-4">
          <div className={`p-3 rounded-full ${oneDriveToken ? 'bg-green-200' : 'bg-gray-200'}`}>
            <Cloud size={24} className={oneDriveToken ? 'text-green-600' : 'text-gray-600'} />
          </div>
          <div>
            <p className="font-bold text-gray-900">OneDrive Integration</p>
            <p className="text-sm text-gray-700">
              {oneDriveToken ? '✓ Connected and ready to sync' : 'Connect to enable automatic file uploads'}
            </p>
          </div>
        </div>
        {!oneDriveToken && (
          <button
            onClick={initiateOneDriveAuth}
            className="bg-blue-600 text-white px-6 py-3 rounded-lg hover:bg-blue-700 font-semibold flex items-center gap-2 whitespace-nowrap"
          >
            <Cloud size={18} /> Connect OneDrive
          </button>
        )}
        {oneDriveToken && (
          <button
            onClick={handleDisconnectOneDrive}
            className="bg-gray-600 text-white px-6 py-3 rounded-lg hover:bg-gray-700 font-semibold flex items-center gap-2 whitespace-nowrap"
          >
            <LogOut size={18} /> Disconnect
          </button>
        )}
      </div>

      {uploadStatus && (
        <div className={`mt-4 p-4 rounded-lg flex items-center gap-3 ${
          uploadStatus.type === 'success' ? 'bg-green-100 text-green-800' :
          uploadStatus.type === 'error' ? 'bg-red-100 text-red-800' :
          uploadStatus.type === 'loading' ? 'bg-blue-100 text-blue-800' :
          'bg-gray-100 text-gray-800'
        }`}>
          {uploadStatus.type === 'loading' && <Loader size={18} className="animate-spin flex-shrink-0" />}
          {uploadStatus.type === 'success' && <Check size={18} className="flex-shrink-0" />}
          {uploadStatus.type === 'error' && <AlertCircle size={18} className="flex-shrink-0" />}
          <span className="font-semibold">{uploadStatus.message}</span>
        </div>
      )}
    </div>
  );

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Header */}
      <header className="bg-gradient-to-r from-blue-600 to-blue-800 text-white shadow-lg">
        <div className="max-w-7xl mx-auto px-6 py-6 flex justify-between items-center">
          <div>
            <h1 className="text-4xl font-bold">Corpworks JMS</h1>
            <p className="text-blue-100 text-sm">Job Management System + OneDrive</p>
          </div>
          <div className="flex items-center gap-4 text-sm">
            {oneDriveToken && (
              <div className="flex items-center gap-2 bg-blue-500 bg-opacity-30 px-3 py-2 rounded-lg">
                <Cloud size={18} className="text-green-300" />
                <span>OneDrive Ready</span>
              </div>
            )}
            <span>{new Date().toLocaleDateString()}</span>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-6 py-8">
        {/* OneDrive Status Bar */}
        {renderOneDriveStatus()}

        {/* Navigation Tabs */}
        <div className="flex gap-2 mb-8 flex-wrap">
          {[
            { id: 'tracker', label: '📊 Job Tracker', icon: FileText },
            { id: 'quoting', label: '💰 Quote Builder', icon: DollarSign },
            { id: 'calendar', label: '📅 Calendar', icon: Calendar },
            { id: 'van', label: '🚐 Van View', icon: Truck }
          ].map(tab => (
            <button
              key={tab.id}
              onClick={() => setActiveTab(tab.id)}
              className={`px-6 py-3 rounded-lg font-semibold transition ${
                activeTab === tab.id
                  ? 'bg-blue-600 text-white shadow-md'
                  : 'bg-white text-gray-700 hover:bg-gray-50 border border-gray-200'
              }`}
            >
              {tab.label}
            </button>
          ))}
        </div>

        {/* Content */}
        {activeTab === 'tracker' && renderJobTracker()}
        {activeTab === 'quoting' && renderQuoteBuilder()}
        {activeTab === 'calendar' && (
          <div className="bg-white rounded-lg p-8 shadow-sm text-center">
            <Calendar size={48} className="mx-auto mb-4 text-gray-400" />
            <h2 className="text-2xl font-bold text-gray-900 mb-2">Calendar View</h2>
            <p className="text-gray-600">Calendar integration coming soon...</p>
          </div>
        )}
        {activeTab === 'van' && (
          <div className="bg-white rounded-lg p-8 shadow-sm text-center">
            <Truck size={48} className="mx-auto mb-4 text-gray-400" />
            <h2 className="text-2xl font-bold text-gray-900 mb-2">Van View</h2>
            <p className="text-gray-600">Van assignments and tracking coming soon...</p>
          </div>
        )}
      </main>
    </div>
  );
}
