// ==UserScript==
// @name         Momence Schedule Preview Launcher
// @namespace    https://momence.com/
// @version      1.0.0
// @description  Add a floating button on Momence dashboard pages that runs the local static schedule build and previews the generated HTML files in a modal.
// @match        https://momence.com/dashboard/13752/*
// @grant        GM_addStyle
// @grant        GM_xmlhttpRequest
// @connect      127.0.0.1
// @connect      localhost
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';

  const BRIDGE_BASE_URL = 'http://127.0.0.1:3210';
  const POLL_INTERVAL_MS = 2500;
  const FILES = [
    { key: 'kemps', label: 'Kemps', path: 'Kemps.html' },
    { key: 'bandra', label: 'Bandra', path: 'Bandra.html' }
  ];

  let pollTimer = null;
  let activeTab = 'kemps';
  let previews = {
    kemps: '',
    bandra: ''
  };

  const elements = {};

  function gmRequest(method, url, body) {
    return new Promise((resolve, reject) => {
      GM_xmlhttpRequest({
        method,
        url,
        data: body ? JSON.stringify(body) : undefined,
        headers: body ? { 'Content-Type': 'application/json' } : undefined,
        onload: response => {
          if (response.status >= 200 && response.status < 300) {
            resolve(response);
            return;
          }

          reject(new Error(`Request failed with status ${response.status}`));
        },
        onerror: () => reject(new Error(`Unable to reach ${url}`)),
        ontimeout: () => reject(new Error(`Request to ${url} timed out`))
      });
    });
  }

  async function requestJson(method, path, body) {
    const response = await gmRequest(method, `${BRIDGE_BASE_URL}${path}`, body);
    return JSON.parse(response.responseText);
  }

  async function requestText(path) {
    const response = await gmRequest('GET', `${BRIDGE_BASE_URL}${path}`);
    return response.responseText;
  }

  function ensureStyles() {
    GM_addStyle(`
      #momence-schedule-preview-button {
        position: fixed;
        right: 24px;
        bottom: 24px;
        z-index: 2147483640;
        border: none;
        border-radius: 999px;
        padding: 12px 18px;
        font: 700 13px/1.2 Inter, system-ui, sans-serif;
        color: #0f172a;
        background: linear-gradient(135deg, #fef08a, #f9a8d4);
        box-shadow: 0 16px 40px rgba(15, 23, 42, 0.28);
        cursor: pointer;
      }

      #momence-schedule-preview-button:hover {
        transform: translateY(-1px);
        box-shadow: 0 20px 46px rgba(15, 23, 42, 0.34);
      }

      #momence-schedule-preview-overlay {
        position: fixed;
        inset: 0;
        display: none;
        align-items: center;
        justify-content: center;
        background: rgba(15, 23, 42, 0.68);
        backdrop-filter: blur(6px);
        z-index: 2147483646;
      }

      #momence-schedule-preview-overlay.open {
        display: flex;
      }

      #momence-schedule-preview-modal {
        width: min(1280px, calc(100vw - 40px));
        height: min(88vh, 920px);
        display: flex;
        flex-direction: column;
        overflow: hidden;
        border-radius: 20px;
        border: 1px solid rgba(255, 255, 255, 0.12);
        background: #020617;
        color: #e2e8f0;
        box-shadow: 0 30px 80px rgba(2, 6, 23, 0.45);
      }

      #momence-schedule-preview-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 16px;
        padding: 18px 22px;
        border-bottom: 1px solid rgba(148, 163, 184, 0.18);
        background: linear-gradient(180deg, rgba(15, 23, 42, 0.98), rgba(2, 6, 23, 0.98));
      }

      #momence-schedule-preview-title {
        font: 700 18px/1.2 Inter, system-ui, sans-serif;
      }

      #momence-schedule-preview-subtitle {
        margin-top: 4px;
        font: 500 12px/1.4 Inter, system-ui, sans-serif;
        color: #94a3b8;
      }

      .momence-schedule-preview-actions {
        display: flex;
        align-items: center;
        gap: 10px;
      }

      .momence-schedule-preview-action {
        border: 1px solid rgba(148, 163, 184, 0.25);
        border-radius: 12px;
        background: #111827;
        color: #e2e8f0;
        padding: 10px 14px;
        cursor: pointer;
        font: 600 12px/1 Inter, system-ui, sans-serif;
      }

      .momence-schedule-preview-action.primary {
        background: linear-gradient(135deg, #34d399, #38bdf8);
        color: #020617;
        border-color: transparent;
      }

      #momence-schedule-preview-body {
        display: grid;
        grid-template-columns: 320px minmax(0, 1fr);
        min-height: 0;
        flex: 1;
      }

      #momence-schedule-preview-sidebar {
        padding: 18px;
        border-right: 1px solid rgba(148, 163, 184, 0.14);
        background: rgba(15, 23, 42, 0.82);
        overflow: auto;
      }

      .momence-sp-section {
        margin-bottom: 18px;
      }

      .momence-sp-label {
        margin-bottom: 8px;
        font: 700 11px/1.3 Inter, system-ui, sans-serif;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        color: #94a3b8;
      }

      #momence-schedule-status {
        border-radius: 14px;
        padding: 12px 14px;
        background: rgba(15, 23, 42, 0.9);
        border: 1px solid rgba(148, 163, 184, 0.16);
        font: 600 12px/1.5 Inter, system-ui, sans-serif;
      }

      .momence-sp-pill {
        display: inline-flex;
        align-items: center;
        border-radius: 999px;
        padding: 4px 10px;
        margin-bottom: 10px;
        font: 700 11px/1 Inter, system-ui, sans-serif;
      }

      .momence-sp-pill.running { background: rgba(56, 189, 248, 0.18); color: #7dd3fc; }
      .momence-sp-pill.success { background: rgba(52, 211, 153, 0.18); color: #6ee7b7; }
      .momence-sp-pill.error { background: rgba(248, 113, 113, 0.18); color: #fca5a5; }
      .momence-sp-pill.idle { background: rgba(148, 163, 184, 0.18); color: #cbd5e1; }

      .momence-sp-file-list {
        display: grid;
        gap: 10px;
      }

      .momence-sp-file-card {
        border-radius: 14px;
        padding: 12px 14px;
        background: rgba(15, 23, 42, 0.78);
        border: 1px solid rgba(148, 163, 184, 0.14);
      }

      .momence-sp-file-name {
        font: 700 13px/1.3 Inter, system-ui, sans-serif;
      }

      .momence-sp-file-meta {
        margin-top: 4px;
        font: 500 11px/1.45 Inter, system-ui, sans-serif;
        color: #94a3b8;
      }

      #momence-schedule-log {
        max-height: 210px;
        overflow: auto;
        white-space: pre-wrap;
        word-break: break-word;
        border-radius: 14px;
        padding: 12px;
        background: #000;
        color: #cbd5e1;
        border: 1px solid rgba(148, 163, 184, 0.14);
        font: 500 11px/1.45 ui-monospace, SFMono-Regular, Menlo, monospace;
      }

      #momence-schedule-preview-main {
        display: flex;
        flex-direction: column;
        min-width: 0;
        min-height: 0;
      }

      #momence-schedule-preview-tabs {
        display: flex;
        gap: 10px;
        padding: 16px 18px 0;
      }

      .momence-schedule-preview-tab {
        border: 1px solid rgba(148, 163, 184, 0.2);
        border-bottom: none;
        border-radius: 14px 14px 0 0;
        background: rgba(15, 23, 42, 0.86);
        color: #cbd5e1;
        padding: 10px 14px;
        cursor: pointer;
        font: 700 12px/1 Inter, system-ui, sans-serif;
      }

      .momence-schedule-preview-tab.active {
        background: #0f172a;
        color: #f8fafc;
      }

      #momence-schedule-preview-frame-wrap {
        flex: 1;
        min-height: 0;
        padding: 0 18px 18px;
      }

      #momence-schedule-preview-frame {
        width: 100%;
        height: 100%;
        border: 1px solid rgba(148, 163, 184, 0.16);
        border-radius: 0 18px 18px 18px;
        background: #fff;
      }

      #momence-schedule-empty-state {
        display: flex;
        align-items: center;
        justify-content: center;
        height: 100%;
        padding: 24px;
        color: #94a3b8;
        font: 600 14px/1.6 Inter, system-ui, sans-serif;
        text-align: center;
      }

      @media (max-width: 1100px) {
        #momence-schedule-preview-body {
          grid-template-columns: 1fr;
        }

        #momence-schedule-preview-sidebar {
          border-right: none;
          border-bottom: 1px solid rgba(148, 163, 184, 0.14);
          max-height: 280px;
        }
      }
    `);
  }

  function createUi() {
    if (document.getElementById('momence-schedule-preview-button')) {
      return;
    }

    ensureStyles();

    const button = document.createElement('button');
    button.id = 'momence-schedule-preview-button';
    button.type = 'button';
    button.textContent = '⚡ Schedule Preview';
    button.addEventListener('click', handleLaunchClick);

    const overlay = document.createElement('div');
    overlay.id = 'momence-schedule-preview-overlay';
    overlay.innerHTML = `
      <div id="momence-schedule-preview-modal" role="dialog" aria-modal="true" aria-label="Schedule preview modal">
        <div id="momence-schedule-preview-header">
          <div>
            <div id="momence-schedule-preview-title">Schedule Preview</div>
            <div id="momence-schedule-preview-subtitle">Runs <code>npm run update -- --static</code> locally, then loads the generated HTML in this modal.</div>
          </div>
          <div class="momence-schedule-preview-actions">
            <button type="button" class="momence-schedule-preview-action" id="momence-schedule-refresh">Refresh previews</button>
            <button type="button" class="momence-schedule-preview-action" id="momence-schedule-rerun">Run again</button>
            <button type="button" class="momence-schedule-preview-action" id="momence-schedule-close">Close</button>
          </div>
        </div>
        <div id="momence-schedule-preview-body">
          <aside id="momence-schedule-preview-sidebar">
            <div class="momence-sp-section">
              <div class="momence-sp-label">Run status</div>
              <div id="momence-schedule-status"></div>
            </div>
            <div class="momence-sp-section">
              <div class="momence-sp-label">Generated files</div>
              <div class="momence-sp-file-list" id="momence-schedule-files"></div>
            </div>
            <div class="momence-sp-section">
              <div class="momence-sp-label">Bridge log tail</div>
              <div id="momence-schedule-log">Waiting for bridge…</div>
            </div>
          </aside>
          <section id="momence-schedule-preview-main">
            <div id="momence-schedule-preview-tabs"></div>
            <div id="momence-schedule-preview-frame-wrap">
              <iframe id="momence-schedule-preview-frame" sandbox="allow-same-origin allow-scripts"></iframe>
              <div id="momence-schedule-empty-state">Kick off a run to preview the latest generated HTML.</div>
            </div>
          </section>
        </div>
      </div>
    `;

    overlay.addEventListener('click', event => {
      if (event.target === overlay) {
        closeModal();
      }
    });

    document.addEventListener('keydown', event => {
      if (event.key === 'Escape' && overlay.classList.contains('open')) {
        closeModal();
      }
    });

    document.body.append(button, overlay);

    elements.button = button;
    elements.overlay = overlay;
    elements.status = overlay.querySelector('#momence-schedule-status');
    elements.files = overlay.querySelector('#momence-schedule-files');
    elements.log = overlay.querySelector('#momence-schedule-log');
    elements.tabs = overlay.querySelector('#momence-schedule-preview-tabs');
    elements.frame = overlay.querySelector('#momence-schedule-preview-frame');
    elements.emptyState = overlay.querySelector('#momence-schedule-empty-state');

    overlay.querySelector('#momence-schedule-close').addEventListener('click', closeModal);
    overlay.querySelector('#momence-schedule-refresh').addEventListener('click', loadPreviews);
    overlay.querySelector('#momence-schedule-rerun').addEventListener('click', triggerRun);

    renderTabs();
  }

  function renderTabs() {
    elements.tabs.innerHTML = '';
    FILES.forEach(file => {
      const tab = document.createElement('button');
      tab.type = 'button';
      tab.className = `momence-schedule-preview-tab${activeTab === file.key ? ' active' : ''}`;
      tab.textContent = file.label;
      tab.addEventListener('click', () => {
        activeTab = file.key;
        renderTabs();
        renderPreviewFrame();
      });
      elements.tabs.appendChild(tab);
    });
  }

  function renderStatus(statusPayload, extraMessage) {
    const status = statusPayload?.status || 'idle';
    const labels = {
      idle: 'Idle',
      running: 'Running',
      success: 'Completed',
      error: 'Failed'
    };
    const fileSummary = statusPayload?.files || {};
    const startedAt = statusPayload?.startedAt ? new Date(statusPayload.startedAt).toLocaleString() : '—';
    const finishedAt = statusPayload?.finishedAt ? new Date(statusPayload.finishedAt).toLocaleString() : '—';
    const command = statusPayload?.command || 'npm run update -- --static';
    const detail = extraMessage || statusPayload?.error || 'Waiting for action.';

    elements.status.innerHTML = `
      <div class="momence-sp-pill ${status}">${labels[status] || status}</div>
      <div><strong>Command:</strong> ${escapeHtml(command)}</div>
      <div><strong>Started:</strong> ${escapeHtml(startedAt)}</div>
      <div><strong>Finished:</strong> ${escapeHtml(finishedAt)}</div>
      <div><strong>Exit code:</strong> ${escapeHtml(String(statusPayload?.exitCode ?? '—'))}</div>
      <div style="margin-top:10px;color:#cbd5e1;">${escapeHtml(detail)}</div>
      <div style="margin-top:10px;color:#94a3b8;">Files detected: ${fileSummary.kemps?.exists ? 'Kemps ✓' : 'Kemps —'} · ${fileSummary.bandra?.exists ? 'Bandra ✓' : 'Bandra —'}</div>
    `;
  }

  function renderFiles(statusPayload) {
    const entries = [statusPayload?.files?.kemps, statusPayload?.files?.bandra].filter(Boolean);

    elements.files.innerHTML = entries.map(file => `
      <div class="momence-sp-file-card">
        <div class="momence-sp-file-name">${escapeHtml(file.name)}</div>
        <div class="momence-sp-file-meta">${file.exists ? 'Available' : 'Not generated yet'}</div>
        <div class="momence-sp-file-meta">Updated: ${escapeHtml(file.updatedAt ? new Date(file.updatedAt).toLocaleString() : '—')}</div>
        <div class="momence-sp-file-meta">Size: ${escapeHtml(file.exists ? `${Math.round(file.size / 1024)} KB` : '0 KB')}</div>
      </div>
    `).join('');
  }

  function renderLogs(statusPayload, fallbackMessage) {
    const logs = statusPayload?.logs?.length ? statusPayload.logs.join('\n') : fallbackMessage;
    elements.log.textContent = logs;
    elements.log.scrollTop = elements.log.scrollHeight;
  }

  function renderPreviewFrame() {
    const html = previews[activeTab];
    if (!html) {
      elements.frame.style.display = 'none';
      elements.emptyState.style.display = 'flex';
      return;
    }

    elements.emptyState.style.display = 'none';
    elements.frame.style.display = 'block';
    elements.frame.srcdoc = html;
  }

  async function loadPreviews() {
    try {
      const htmlByKey = {};
      for (const file of FILES) {
        htmlByKey[file.key] = await requestText(`/html/${encodeURIComponent(file.path)}`);
      }

      previews = htmlByKey;
      renderPreviewFrame();
    } catch (error) {
      renderPreviewFrame();
      renderLogs(null, `Preview files are not ready yet.\n\n${error.message}`);
    }
  }

  function openModal() {
    elements.overlay.classList.add('open');
  }

  function closeModal() {
    elements.overlay.classList.remove('open');
  }

  function stopPolling() {
    if (pollTimer) {
      clearInterval(pollTimer);
      pollTimer = null;
    }
  }

  function startPolling() {
    stopPolling();
    pollTimer = window.setInterval(refreshStatus, POLL_INTERVAL_MS);
  }

  async function refreshStatus() {
    try {
      const statusPayload = await requestJson('GET', '/status');
      renderStatus(statusPayload);
      renderFiles(statusPayload);
      renderLogs(statusPayload, 'No logs yet.');

      if (statusPayload.status === 'success') {
        stopPolling();
        await loadPreviews();
      } else if (statusPayload.status === 'error') {
        stopPolling();
      }
    } catch (error) {
      stopPolling();
      renderStatus({ status: 'error', exitCode: null, command: 'npm run update -- --static' }, `Bridge unavailable: ${error.message}`);
      renderFiles({ files: { kemps: { name: 'Kemps.html', exists: false, updatedAt: null, size: 0 }, bandra: { name: 'Bandra.html', exists: false, updatedAt: null, size: 0 } } });
      renderLogs(null, `Cannot reach the local bridge at ${BRIDGE_BASE_URL}.\n\nStart it in the project folder with:\n  npm run preview:bridge`);
    }
  }

  async function triggerRun() {
    renderStatus({ status: 'running', command: 'npm run update -- --static', exitCode: null }, 'Requesting a fresh background run…');
    renderLogs(null, 'Waiting for bridge response…');

    try {
      const response = await requestJson('POST', '/run', {});
      renderStatus(response, response.started ? 'Background update started.' : 'A run is already in progress, following along.');
      renderFiles(response);
      renderLogs(response, 'No logs yet.');
      startPolling();
      await refreshStatus();
    } catch (error) {
      renderStatus({ status: 'error', command: 'npm run update -- --static', exitCode: null }, `Could not start run: ${error.message}`);
      renderLogs(null, `Bridge request failed.\n\n${error.message}`);
    }
  }

  async function handleLaunchClick() {
    createUi();
    openModal();
    await refreshStatus();
    await triggerRun();
  }

  function escapeHtml(value) {
    return String(value)
      .replaceAll('&', '&amp;')
      .replaceAll('<', '&lt;')
      .replaceAll('>', '&gt;')
      .replaceAll('"', '&quot;')
      .replaceAll("'", '&#39;');
  }

  createUi();
})();