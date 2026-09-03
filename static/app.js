// Frontend Client Logic for Feedback Automation App
document.addEventListener('DOMContentLoaded', () => {
  // Elements
  const form = document.getElementById('automationForm');
  const pdfInput = document.getElementById('pdfInput');
  const pdfDropZone = document.getElementById('pdfDropZone');
  const pdfFileLabel = document.getElementById('pdfFileLabel');
  const pdfBadge = document.getElementById('pdfBadge');

  const templateInput = document.getElementById('templateInput');
  const templateDropZone = document.getElementById('templateDropZone');
  const templateFileLabel = document.getElementById('templateFileLabel');
  const templateBadge = document.getElementById('templateBadge');

  const engineCards = document.querySelectorAll('.engine-card');
  const aiSettingsPanel = document.getElementById('aiSettingsPanel');
  const apiKeyLabel = document.getElementById('apiKeyLabel');
  const apiKeyInput = document.getElementById('apiKeyInput');
  const apiKeyHint = document.getElementById('apiKeyHint');
  const modelSelect = document.getElementById('modelSelect');
  const toggleApiKeyVisibility = document.getElementById('toggleApiKeyVisibility');

  const processingCard = document.getElementById('processingCard');
  const terminalLog = document.getElementById('terminalLog');
  const resultsCard = document.getElementById('resultsCard');
  const downloadBtn = document.getElementById('downloadBtn');
  const resetBtn = document.getElementById('resetBtn');

  const kpiParticipants = document.getElementById('kpiParticipants');
  const kpiScores = document.getElementById('kpiScores');
  const kpiSuggestions = document.getElementById('kpiSuggestions');
  const kpiTotalCells = document.getElementById('kpiTotalCells');
  const participantTableBody = document.getElementById('participantTableBody');
  const tableSearch = document.getElementById('tableSearch');

  let selectedPdfFile = null;
  let selectedTemplateFile = null;
  let currentParticipantsData = [];
  let serverInfo = { has_gemini_key: false, has_anthropic_key: false };

  // Fetch server configuration info
  fetch('/api/info')
    .then(res => res.json())
    .then(data => {
      serverInfo = data;
      const checkedMode = document.querySelector('input[name="mode"]:checked').value;
      updateEngineUI(checkedMode);

      if (data.available_sheets && data.available_sheets.length > 0) {
        const sheetSelect = document.getElementById('sheetSelect');
        const participantSheets = data.available_sheets.filter(s => s.toLowerCase().includes("participant"));
        const sheetsToUse = participantSheets.length > 0 ? participantSheets : data.available_sheets;
        sheetSelect.innerHTML = sheetsToUse.map(s => `<option value="${s}">${s}</option>`).join('');
      }
    })
    .catch(err => console.log('Could not fetch server info:', err));

  // Drag & Drop handlers for PDF
  setupDragAndDrop(pdfDropZone, pdfInput, (file) => {
    if (file && file.name.endsWith('.pdf')) {
      selectedPdfFile = file;
      pdfFileLabel.innerHTML = `<strong>${file.name}</strong> (${formatBytes(file.size)})`;
      pdfBadge.style.display = 'inline-block';
      pdfBadge.className = 'file-badge loaded';
      pdfBadge.textContent = 'PDF Attached';
    }
  });

  // Drag & Drop handlers for Template
  setupDragAndDrop(templateDropZone, templateInput, (file) => {
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
      selectedTemplateFile = file;
      templateFileLabel.innerHTML = `<strong>${file.name}</strong> (${formatBytes(file.size)})`;
      templateBadge.className = 'file-badge loaded';
      templateBadge.textContent = 'Custom Template Attached';
    }
  });

  // Engine selection
  engineCards.forEach(card => {
    card.addEventListener('click', () => {
      engineCards.forEach(c => c.classList.remove('active'));
      card.classList.add('active');
      const radio = card.querySelector('input[type="radio"]');
      radio.checked = true;

      const mode = radio.value;
      updateEngineUI(mode);
    });
  });

  function updateEngineUI(mode) {
    if (mode === 'demo') {
      aiSettingsPanel.style.display = 'none';
    } else if (mode === 'gemini') {
      aiSettingsPanel.style.display = 'block';
      apiKeyLabel.textContent = 'Google Gemini API Key';
      if (serverInfo.has_gemini_key) {
        apiKeyInput.placeholder = '•••••••• (Pre-configured in config.py / .env)';
        apiKeyHint.innerHTML = '<span style="color:#34d399;font-weight:600;">✅ Key configured in code / .env (No need to enter key)</span>';
      } else {
        apiKeyInput.placeholder = 'AIzaSy... or paste key here';
        apiKeyHint.innerHTML = 'Paste your key here or save it once in <code>config.py</code> / <code>.env</code>. Free key from <a href="https://aistudio.google.com" target="_blank" rel="noopener">Google AI Studio</a>.';
      }
      modelSelect.innerHTML = `
        <option value="gemini-3.6-flash">gemini-3.6-flash (Recommended, High Quota & Vision)</option>
        <option value="gemini-3.5-flash-lite">gemini-3.5-flash-lite (Ultra-Fast)</option>
        <option value="gemini-3.7-flash">gemini-3.7-flash (Latest Flash Model)</option>
      `;
    } else if (mode === 'claude') {
      aiSettingsPanel.style.display = 'block';
      apiKeyLabel.textContent = 'Anthropic Claude API Key';
      if (serverInfo.has_anthropic_key) {
        apiKeyInput.placeholder = '•••••••• (Pre-configured in config.py / .env)';
        apiKeyHint.innerHTML = '<span style="color:#34d399;font-weight:600;">✅ Key configured in code / .env (No need to enter key)</span>';
      } else {
        apiKeyInput.placeholder = 'sk-ant-... or paste key here';
        apiKeyHint.innerHTML = 'Paste your key here or save it once in <code>config.py</code> / <code>.env</code>. Key from <a href="https://console.anthropic.com" target="_blank" rel="noopener">Anthropic Console</a>.';
      }
      modelSelect.innerHTML = `
        <option value="claude-3-7-sonnet-latest">claude-3-7-sonnet-latest (Recommended)</option>
        <option value="claude-3-5-sonnet-latest">claude-3-5-sonnet-latest</option>
      `;
    }
  }


  // Toggle API key visibility
  toggleApiKeyVisibility.addEventListener('click', () => {
    apiKeyInput.type = apiKeyInput.type === 'password' ? 'text' : 'password';
  });

  // Form submission
  form.addEventListener('submit', async (e) => {
    e.preventDefault();

    const mode = document.querySelector('input[name="mode"]:checked').value;
    const sheetName = document.getElementById('sheetSelect').value;
    const overwrite = document.getElementById('overwriteCheck').checked;
    const pagesPerParticipant = document.getElementById('pagesPerParticipant').value;
    const apiKey = apiKeyInput.value.trim();
    const model = modelSelect.value;
    const progressBarFill = document.getElementById('progressBarFill');

    if (mode !== 'demo' && !selectedPdfFile) {
      alert('Please select or drag-and-drop a scanned feedback PDF file first.');
      return;
    }

    // Switch to processing UI
    form.style.display = 'none';
    resultsCard.style.display = 'none';
    processingCard.style.display = 'block';
    if (progressBarFill) progressBarFill.style.width = '10%';
    terminalLog.innerHTML = '<div class="log-line info">> Initializing automation pipeline...</div>';

    addLog(`Mode selected: ${mode.toUpperCase()}`);
    if (mode === 'demo') {
      addLog('> Using pre-extracted sample data (instant filling)...');
    } else {
      addLog(`> Uploading PDF: ${selectedPdfFile ? selectedPdfFile.name : 'upload.pdf'}...`);
      addLog(`> AI Vision Engine: ${model}...`);
    }

    const formData = new FormData();
    if (selectedPdfFile) formData.append('pdf_file', selectedPdfFile);
    if (selectedTemplateFile) formData.append('template_file', selectedTemplateFile);
    formData.append('mode', mode);
    formData.append('sheet_name', sheetName);
    formData.append('pages_per_participant', pagesPerParticipant);
    formData.append('overwrite', overwrite);
    if (apiKey) formData.append('api_key', apiKey);
    if (model) formData.append('model', model);

    try {
      const resp = await fetch('/api/process-stream', {
        method: 'POST',
        body: formData,
      });

      if (!resp.ok) {
        throw new Error(`Server returned status ${resp.status}`);
      }

      const reader = resp.body.getReader();
      const decoder = new TextDecoder('utf-8');
      let buffer = '';
      let completedResult = null;

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n\n');
        buffer = lines.pop(); // keep last partial chunk

        for (const block of lines) {
          const trimmed = block.trim();
          if (trimmed.startsWith('data:')) {
            const jsonStr = trimmed.replace(/^data:\s*/, '');
            try {
              const data = JSON.parse(jsonStr);
              if (data.type === 'log') {
                addLog(data.message);
              } else if (data.type === 'progress') {
                if (progressBarFill && data.percent !== undefined) {
                  progressBarFill.style.width = `${Math.min(data.percent, 100)}%`;
                }
              } else if (data.type === 'complete') {
                completedResult = data;
                if (progressBarFill) progressBarFill.style.width = '100%';
                addLog('> [SUCCESS] Excel workbook populated successfully!');
              } else if (data.type === 'error') {
                throw new Error(data.message);
              }
            } catch (jsonErr) {
              if (jsonErr.message && !jsonErr.message.includes('Unexpected token')) {
                throw jsonErr;
              }
            }
          }
        }
      }

      if (completedResult) {
        setTimeout(() => {
          showResults(completedResult);
        }, 500);
      } else {
        throw new Error('Process ended without receiving output summary.');
      }

    } catch (err) {
      addLog(`> [ERROR] ${err.message}`);
      alert(`Error processing feedback forms: ${err.message}`);
      form.style.display = 'block';
      processingCard.style.display = 'none';
    }
  });

  function showResults(data) {
    processingCard.style.display = 'none';
    resultsCard.style.display = 'block';

    downloadBtn.href = data.download_url;
    downloadBtn.setAttribute('download', data.filename);

    kpiParticipants.textContent = data.summary.participants_count;
    kpiScores.textContent = data.summary.total_scores;
    kpiSuggestions.textContent = data.summary.total_suggestions;
    kpiTotalCells.textContent = data.summary.total_cells;

    currentParticipantsData = data.participants || [];
    renderTable(currentParticipantsData);
  }

  function renderTable(participants) {
    participantTableBody.innerHTML = '';
    participants.forEach((p, idx) => {
      const row = document.createElement('tr');
      row.innerHTML = `
        <td><strong>${idx + 1}</strong></td>
        <td><code>${p.staff_no}</code></td>
        <td>${p.name}</td>
        <td><span class="file-badge active-badge">${p.scores_filled} scores</span></td>
        <td>${p.suggestions_filled} text responses</td>
        <td style="max-width: 260px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: #94a3b8;" title="${p.suggestion_2 || ''}">
          ${p.suggestion_2 || '<em style="color:#64748b">None</em>'}
        </td>
      `;
      participantTableBody.appendChild(row);
    });
  }

  // Filter Table
  tableSearch.addEventListener('input', (e) => {
    const q = e.target.value.toLowerCase();
    const filtered = currentParticipantsData.filter(p => 
      p.staff_no.toLowerCase().includes(q) || p.name.toLowerCase().includes(q) || (p.suggestion_2 && p.suggestion_2.toLowerCase().includes(q))
    );
    renderTable(filtered);
  });

  // Reset Button
  resetBtn.addEventListener('click', () => {
    resultsCard.style.display = 'none';
    form.style.display = 'block';
  });

  // Helper: Drag & drop
  function setupDragAndDrop(zone, input, onFile) {
    ['dragenter', 'dragover'].forEach(eventName => {
      zone.addEventListener(eventName, (e) => {
        e.preventDefault();
        zone.classList.add('dragover');
      });
    });

    ['dragleave', 'drop'].forEach(eventName => {
      zone.addEventListener(eventName, (e) => {
        e.preventDefault();
        zone.classList.remove('dragover');
      });
    });

    zone.addEventListener('drop', (e) => {
      const files = e.dataTransfer.files;
      if (files && files.length > 0) {
        input.files = files;
        onFile(files[0]);
      }
    });

    input.addEventListener('change', (e) => {
      if (input.files && input.files.length > 0) {
        onFile(input.files[0]);
      }
    });
  }

  function addLog(msg) {
    const line = document.createElement('div');
    line.className = 'log-line';
    line.textContent = msg;
    terminalLog.appendChild(line);
    terminalLog.scrollTop = terminalLog.scrollHeight;
  }

  function formatBytes(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  }
});
