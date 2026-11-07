/* Trade Flow Report Generator — UI logic */
(function () {
  'use strict';

  const $ = id => document.getElementById(id);
  const logEl = $('log');
  const errorMsg = $('error-msg');
  const hint = $('upload-hint');
  const quarterlyInfo = $('quarterly-info');
  const exportPeriodWrap = $('export-period-wrap');
  const productInfo = $('product-profile-wrap');
  const competitorInfo = $('competitor-wrap');
  const productCfgWrapper = $('product-cfg-wrapper');
  const productCfgSelect = $('product-cfg');
  const filesInput = $('files');
  const runBtn = $('run');
  const runCta = $('run-cta');
  const fileSummary = $('file-summary');
  const progressWrap = $('progress-wrap');
  const progressStageText = $('progress-stage-text');
  const progressElapsed = $('progress-elapsed');
  const progressFill = $('progress-fill');
  const toastContainer = $('toast-container');
  const topWrap = $('top-wrap');
  const dropzone = $('dropzone');
  const sampleSelect = $('sample');
  const loadSampleBtn = $('load-sample');
  const sampleSummary = $('sample-summary');
  const downloadSampleBtn = $('download-sample');
  const sourcesCard = $('sources-card');
  const sourceSelect = $('source-select');
  const sourceFetchBtn = $('source-fetch');
  const sourceDescription = $('source-description');
  const sourceParams = $('source-params');
  const sourceStatus = $('source-status');
  const sourceResult = $('source-result');
  const sourceNotes = $('source-notes');
  const sourceFiles = $('source-files');
  const sourceActions = $('source-actions');
  const insightCard = $('insight-card');
  const previewFileSelect = $('preview-file');
  const previewChart = $('preview-chart');
  const previewTableWrap = $('preview-table-wrap');
  const qualitySummary = $('quality-summary');
  const qualityList = $('quality-list');
  const qualityFiles = $('quality-files');
  const shareLinkInput = $('share-link');
  const shareCreateBtn = $('share-create');
  const shareCopyBtn = $('share-copy');
  const shareBtn = $('share-btn');
  const commentsList = $('comments-list');
  const commentsCount = $('comments-count');
  const commentCompose = $('comment-compose');
  const commentAuthor = $('comment-author');
  const commentBody = $('comment-body');
  const commentAddBtn = $('comment-add');
  const jobsList = $('jobs-list');
  const jobsRefreshBtn = $('jobs-refresh');

  const TOKENS_KEY = 'tf_job_store_v1';
  const PREVIEW_ROWS = 12;
  const PREVIEW_COLS = 10;

  let sourceCatalog = [];
  let activeSource = null;
  let sourceChosen = false;      // the user picked a source, so keep it on tab change
  let activeFetch = null;
  let previewFiles = [];
  let currentJob = null;
  let sharedMode = false;
  let jobStore = readJobStore();

  const POLL_INTERVAL_MS = 2000;
  const POLL_MAX_FAILURES = 5;       // consecutive fetch errors tolerated
  const POLL_MAX_POLLS = 900;        // ~30 minutes at 2 s per poll

  let elapsedTimer = null;
  let currentFiles = [];             // mirrors filesInput.files, removable

  const LOG_EMPTY_HTML =
    '<div class="log-empty"><i class="fas fa-inbox" aria-hidden="true"></i>' +
    'No run yet &mdash; upload your files in step 2, then choose Generate Report.</div>';

  const esc = s => String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const fmtSize = b => b >= 1048576 ? (b / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(b / 1024)) + ' KB';

  // -------------------------------------------------------------------------
  // Log console
  // -------------------------------------------------------------------------
  function classifyLine(text) {
    const t = String(text);
    if (/\[ERROR\]|\bERROR\b|^Error|failed|Traceback/i.test(t)) return 'err';
    if (/\[warn\]|\bWarning\b|\bwarn\b/i.test(t)) return 'warn';
    return 'ok';
  }

  function appendLog(text, cls) {
    const emptyEl = logEl.querySelector('.log-empty');
    if (emptyEl) emptyEl.remove();
    const line = document.createElement('div');
    line.textContent = text;
    line.className = cls || classifyLine(text);
    logEl.appendChild(line);
    // Stick to the bottom only when the user is already at (near) the bottom;
    // never yank them away from history they are reading.
    const atBottom = logEl.scrollHeight - logEl.scrollTop - logEl.clientHeight < 48;
    if (atBottom) logEl.scrollTop = logEl.scrollHeight;
  }

  // -------------------------------------------------------------------------
  // Errors / toasts
  // -------------------------------------------------------------------------
  function normalizeDetail(detail) {
    if (detail == null) return '';
    if (typeof detail === 'string') return detail;
    if (Array.isArray(detail)) {
      return detail.map(d => (d && (d.msg || d.message)) || JSON.stringify(d)).join('; ');
    }
    return detail.msg || detail.message || JSON.stringify(detail);
  }

  function showError(msg) {
    errorMsg.textContent = msg;
    errorMsg.style.display = 'block';
  }

  function clearError() {
    errorMsg.style.display = 'none';
    errorMsg.textContent = '';
  }

  function clearResult() {
    const r = $('result');
    r.style.display = 'none';
    r.classList.remove('showing');
    $('result-meta').innerHTML = '';
    $('dl-xlsx').style.display = 'none';
  }

  function showToast(type, title, body) {
    const t = document.createElement('div');
    t.className = 'toast toast-' + type;
    t.setAttribute('role', 'status');
    t.tabIndex = 0;
    const icon = type === 'success' ? 'fa-circle-check' : 'fa-circle-exclamation';
    t.innerHTML = '<i class="fas ' + icon + ' toast-icon" aria-hidden="true"></i>' +
      '<div class="toast-body"><b>' + esc(title) + '</b><span>' + esc(body || '') + '</span></div>' +
      '<button type="button" class="toast-dismiss" aria-label="Dismiss notification">&times;</button>';
    const dismiss = () => {
      if (!t.parentNode) return;
      t.classList.add('leaving');
      setTimeout(() => t.remove(), 320);
    };
    t.querySelector('.toast-dismiss').addEventListener('click', e => {
      e.stopPropagation();
      dismiss();
    });
    if (type === 'success') {
      const go = () => {
        document.querySelector('.log-card').scrollIntoView({ behavior: 'smooth', block: 'start' });
        dismiss();
      };
      t.addEventListener('click', go);
      t.addEventListener('keydown', e => {
        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); go(); }
      });
    } else {
      t.addEventListener('click', dismiss);
      t.addEventListener('keydown', e => {
        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); dismiss(); }
      });
    }
    toastContainer.appendChild(t);
    setTimeout(dismiss, 10000);
  }

  // -------------------------------------------------------------------------
  // Progress (driven by real state, not a fake stage carousel)
  // -------------------------------------------------------------------------
  function setStage(text) {
    progressStageText.textContent = text;
  }

  function startProgress(initialStage) {
    progressWrap.classList.add('show');
    logEl.setAttribute('aria-busy', 'true');
    setStage(initialStage || 'Starting…');
    progressFill.classList.add('uploading');
    progressFill.style.width = '2%';
    const t0 = Date.now();
    progressElapsed.textContent = '0s';
    elapsedTimer = setInterval(() => {
      progressElapsed.textContent = Math.floor((Date.now() - t0) / 1000) + 's';
    }, 1000);
  }

  function setUploadProgress(frac) {
    // 0..1 within the upload phase (first ~15% of the bar)
    progressFill.style.width = Math.max(2, Math.round(15 * frac)) + '%';
  }

  function setBuildProgress(frac, state) {
    // 0..1 within the build phase (remaining 85%)
    const base = Math.round(15 + 85 * Math.min(1, Math.max(0, frac)));
    progressFill.style.width = base + '%';
    progressFill.classList.remove('uploading', 'building', 'complete');
    if (state === 'building') progressFill.classList.add('building');
    else if (state === 'complete') progressFill.classList.add('complete');
  }

  function stopProgress() {
    if (elapsedTimer) { clearInterval(elapsedTimer); elapsedTimer = null; }
    logEl.setAttribute('aria-busy', 'false');
    progressFill.classList.remove('uploading');
    progressFill.style.width = '';
    progressWrap.classList.remove('show');
  }

  // Map a server log line to a coarse progress fraction so the bar advances
  // with real work instead of a timer.
  function fractionForLogLine(line) {
    const t = String(line);
    if (/waiting for the renderer/i.test(t)) return 0.05;
    if (/Renderer acquired|build starting/i.test(t)) return 0.1;
    if (/Mode:|reading/i.test(t)) return 0.2;
    if (/Country detected|Profile:|Data years/i.test(t)) return 0.35;
    if (/tables written|workbooks written/i.test(t)) return 0.6;
    if (/Building|charts/i.test(t)) return 0.75;
    if (/saved as|deliverable saved/i.test(t)) return 0.9;
    if (/Report ready/i.test(t)) return 1.0;
    return null;
  }

  function stageForLogLine(line) {
    const t = String(line);
    if (/waiting for the renderer/i.test(t)) return 'Waiting for the renderer (another job is running)…';
    if (/Renderer acquired|build starting/i.test(t)) return 'Build starting…';
    if (/Mode: raw ITC/i.test(t)) return 'Reading raw ITC spreadsheets…';
    if (/Mode: raw ITC services/i.test(t)) return 'Reading services spreadsheets…';
    if (/Mode: Export Performance/i.test(t)) return 'Reading monthly extracts…';
    if (/Mode: (per-product|supplying-market)/i.test(t)) return 'Reading supplying-markets files…';
    if (/Layout: per (importing market|Kenya export market)/i.test(t)) return 'Reading supplying-markets files by Kenya export market…';
    if (/Layout: per product/i.test(t)) return 'Reading supplying-markets files by product…';
    if (/Country detected/i.test(t)) return 'Country detected — loading config…';
    if (/Building report|Building services report/i.test(t)) return 'Building the Word report…';
    if (/Building product profile|Building crafts/i.test(t)) return 'Building the product profile…';
    if (/tables written|workbooks written/i.test(t)) return 'Tables written — building the report…';
    if (/Report ready/i.test(t)) return 'Report ready!';
    return null;
  }

  // -------------------------------------------------------------------------
  // Panel visibility (single helper; also keeps aria-hidden in sync)
  // -------------------------------------------------------------------------
  function showPanel(el, visible) {
    el.style.display = visible ? 'block' : 'none';
    if (visible) el.removeAttribute('aria-hidden');
    else el.setAttribute('aria-hidden', 'true');
  }

  const PANELS = {
    productCfgWrapper: () => productCfgWrapper,
    quarterlyInfo: () => quarterlyInfo,
    exportPeriodWrap: () => exportPeriodWrap,
    productInfo: () => productInfo,
    competitorInfo: () => competitorInfo,
    topWrap: () => topWrap,
    sourcesCard: () => sourcesCard,
  };

  function setPanels(state) {
    for (const [k, get] of Object.entries(PANELS)) {
      showPanel(get(), !!state[k]);
    }
  }

  const MODE_HINTS = {
    goods:
      '<i class="fas fa-circle-info" style="color:var(--red);" aria-hidden="true"></i> &nbsp;' +
      'Upload the <b>six ITC downloads</b> or a <b>.zip</b> &mdash; or the ready-made <b>Table 1&ndash;6 + Figure 1</b> workbook. ' +
      '<b>.xls</b> / <b>.xlsb</b> / <b>.csv</b> convert automatically.',
    services:
      '<i class="fas fa-circle-info" style="color:var(--red);" aria-hidden="true"></i> &nbsp;' +
      'Upload the <b>seven ITC services files</b> or a <b>.zip</b> (.xls converts automatically).',
    quarterly:
      '<i class="fas fa-circle-info" style="color:var(--red);" aria-hidden="true"></i> &nbsp;' +
      'Upload the <b>KNBS workbook</b> and/or <b>KRA</b> exports/imports extracts (up to 4 files or a <b>.zip</b>), ' +
      'then pick a <b>review period</b>.',
    product:
      '<i class="fas fa-circle-info" style="color:var(--red);" aria-hidden="true"></i> &nbsp;' +
      'Upload the <b>product-family matrix files</b> or a <b>.zip</b> &mdash; the family is auto-detected from the filenames.',
    competitor:
      '<i class="fas fa-circle-info" style="color:var(--red);" aria-hidden="true"></i> &nbsp;' +
      'Run after the goods report. Upload the <b>&ldquo;List of supplying markets&rdquo;</b> files (or a <b>.zip</b>) &mdash; ' +
      'one file per product or one file per Kenya export market is detected automatically ' +
      '&rarr; .png / .docx / .xlsx.',
  };

  function applyReportType(rt) {
    clearError();
    clearResult();
    hint.innerHTML = MODE_HINTS[rt] || MODE_HINTS.goods;
    // Switching modes invalidates the current run output — clear the log so
    // stale lines from another pipeline don't confuse the next run.
    if (logEl.textContent.trim()) {
      logEl.innerHTML = LOG_EMPTY_HTML;
      refreshLogEmpty();
    }
    if (rt === 'goods') {
      setPanels({ topWrap: true, sourcesCard: true });
    } else if (rt === 'services') {
      setPanels({ topWrap: true, sourcesCard: true });
    } else if (rt === 'quarterly') {
      setPanels({ quarterlyInfo: true, exportPeriodWrap: true, topWrap: true });
    } else if (rt === 'product') {
      setPanels({ productInfo: true, productCfgWrapper: true, topWrap: true,
        sourcesCard: true });
    } else if (rt === 'competitor') {
      setPanels({ competitorInfo: true });
    }
    // The card shows where a source can actually be used: UNCTADstat builds
    // the Services Brief, KNBS refreshes the reference series that the goods,
    // services and product reports all cross-check against.
    if (sourceCatalog.length) renderSourceOptions();
    // Keep the semi-annual half picker consistent with the selected period.
    const semi = document.querySelector('input[name="export_period"]:checked');
    const wrap = $('semi-half-wrap');
    const semiOn = !!semi && semi.value === 'semi' && rt === 'quarterly';
    wrap.style.display = semiOn ? 'flex' : 'none';
    if (semiOn) wrap.removeAttribute('aria-hidden');
    else wrap.setAttribute('aria-hidden', 'true');
  }

  // -------------------------------------------------------------------------
  // File selection: chips with per-file remove + drag & drop
  // -------------------------------------------------------------------------
  function syncInputFromCurrentFiles() {
    const dt = new DataTransfer();
    currentFiles.forEach(f => dt.items.add(f));
    filesInput.files = dt.files;
  }

  function refreshLogEmpty() {
    const emptyEl = logEl.querySelector('.log-empty');
    if (!emptyEl && !logEl.querySelector('div:not(.log-empty)')) logEl.innerHTML = LOG_EMPTY_HTML;
  }

  function renderChips() {
    const files = currentFiles;
    if (!files.length) {
      fileSummary.classList.remove('show');
      fileSummary.innerHTML = '';
      runCta.classList.remove('show');
      runBtn.classList.remove('btn-attention');
      refreshLogEmpty();
      return;
    }
    const totalSize = files.reduce((s, f) => s + f.size, 0);
    let chips = '<span class="file-chip"><i class="fas fa-check-circle" aria-hidden="true"></i>&nbsp;' +
      files.length + ' file' + (files.length === 1 ? '' : 's') + ' &middot; ' + fmtSize(totalSize) + '</span>';
    chips += files.map((f, i) =>
      '<span class="file-chip" title="' + esc(f.name) + '">' +
      '<i class="fas fa-file-excel" aria-hidden="true"></i>' +
      '<span>' + esc(f.name) + '</span>' +
      '<button type="button" class="chip-x" data-index="' + i + '" ' +
      'aria-label="Remove ' + esc(f.name) + '">&times;</button></span>'
    ).join('');
    chips += '<button type="button" class="btn-quiet" id="clear-files">' +
      '<i class="fas fa-xmark" aria-hidden="true"></i> Clear all</button>';
    fileSummary.innerHTML = '<div class="file-chips">' + chips + '</div>';
    fileSummary.classList.add('show');
    runCta.classList.add('show');
    runBtn.classList.add('btn-attention');
    refreshLogEmpty();
  }

  function setFiles(fileList) {
    currentFiles = [...fileList];
    syncInputFromCurrentFiles();
    renderChips();
  }

  fileSummary.addEventListener('click', e => {
    if (e.target.closest('#clear-files')) {
      setFiles([]);
      sampleFiles = [];
      sampleSummary.innerHTML = '';
      sampleSummary.classList.remove('show');
      appendLog('All files cleared.');
      return;
    }
    const btn = e.target.closest('.chip-x');
    if (!btn) return;
    const idx = parseInt(btn.dataset.index, 10);
    currentFiles.splice(idx, 1);
    syncInputFromCurrentFiles();
    renderChips();
  });

  filesInput.addEventListener('change', () => {
    clearError();
    currentFiles = [...filesInput.files];
    renderChips();
  });

  // Drag & drop onto the file zone (window-level guards stop the browser from
  // navigating away when a file misses the dropzone).
  ['dragover', 'drop'].forEach(ev => window.addEventListener(ev, e => e.preventDefault()));
  if (dropzone) {
    ['dragenter', 'dragover'].forEach(ev =>
      dropzone.addEventListener(ev, e => {
        e.preventDefault();
        dropzone.classList.add('is-over');
      }));
    ['dragleave', 'dragend'].forEach(ev =>
      dropzone.addEventListener(ev, e => {
        if (e.relatedTarget && dropzone.contains(e.relatedTarget)) return;
        dropzone.classList.remove('is-over');
      }));
    dropzone.addEventListener('drop', e => {
      e.preventDefault();
      dropzone.classList.remove('is-over');
      const dropped = [...(e.dataTransfer?.files || [])];
      if (!dropped.length) return;
      sampleFiles = [];
      setFiles(dropped);
      clearError();
      appendLog('Added ' + dropped.length + ' dropped file' + (dropped.length === 1 ? '' : 's') + '.');
    });
  }

  // -------------------------------------------------------------------------
  // Profile / sample loading (with real error states)
  // -------------------------------------------------------------------------
  async function loadProductProfiles() {
    try {
      const r = await fetch('/api/product-profiles');
      if (!r.ok) throw new Error('HTTP ' + r.status);
      const profiles = await r.json();
      productCfgSelect.innerHTML = '';
      const auto = document.createElement('option');
      auto.value = '__auto__';
      auto.textContent = 'Auto-detect from files';
      productCfgSelect.appendChild(auto);
      for (const c of profiles) {
        const o = document.createElement('option');
        o.value = c.id;
        o.textContent = c.name;
        productCfgSelect.appendChild(o);
      }
      productCfgSelect.value = '__auto__';
      if (!profiles.length) {
        const o = document.createElement('option');
        o.value = '';
        o.textContent = 'No product profiles found in config/';
        productCfgSelect.appendChild(o);
      }
    } catch (e) {
      productCfgSelect.innerHTML = '';
      const o = document.createElement('option');
      o.value = '';
      o.textContent = 'Could not load profiles — reload to retry';
      productCfgSelect.appendChild(o);
    }
  }

  let sampleGroups = [];
  let sampleFiles = [];

  async function loadSamples() {
    try {
      const r = await fetch('/api/samples');
      if (!r.ok) throw new Error('HTTP ' + r.status);
      sampleGroups = await r.json();
      sampleSelect.innerHTML = '<option value="">Select a sample dataset&hellip;</option>';
      for (const g of sampleGroups) {
        const o = document.createElement('option');
        o.value = g.id;
        o.textContent = g.label + ' — ' + g.description + ' (' + g.files.length + ' files)';
        sampleSelect.appendChild(o);
      }
      if (!sampleGroups.length) {
        const o = document.createElement('option');
        o.value = '';
        o.textContent = 'No bundled sample datasets on this server';
        sampleSelect.appendChild(o);
      }
    } catch (e) {
      sampleSelect.innerHTML = '';
      const o = document.createElement('option');
      o.value = '';
      o.textContent = 'Could not load samples — reload to retry';
      sampleSelect.appendChild(o);
    }
  }

  loadSampleBtn.addEventListener('click', async () => {
    const gid = sampleSelect.value;
    if (!gid) {
      showToast('error', 'No sample selected', 'Pick a bundled sample dataset first.');
      return;
    }
    if (!sampleGroups.length) {
      showError('Sample list not loaded yet — retry in a moment.');
      return;
    }
    const g = sampleGroups.find(x => x.id === gid);
    const btn = loadSampleBtn;
    btn.disabled = true;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin" aria-hidden="true"></i> Loading…';
    clearError();
    try {
      // Pre-select the analysis type this sample belongs to, so e.g. goods
      // samples only ever feed the goods pipeline.
      const radios = document.querySelectorAll('input[name="report_type"]');
      let matched = false;
      radios.forEach(r => {
        if (r.value === g.report_type) {
          r.checked = true;
          applyReportType(r.value);
          matched = true;
        }
      });
      if (!matched && radios.length) {
        radios[0].checked = true;
        applyReportType(radios[0].value);
      }
      // Apply the recommended profile for product samples.
      if (g.report_type === 'product' && g.config &&
          [...productCfgSelect.options].some(o => o.value === g.config)) {
        productCfgSelect.value = g.config;
      }

      // Stream every sample file from the server into the upload queue.
      const loaded = [];
      for (const fn of g.files) {
        const r = await fetch('/api/sample-file/' + encodeURIComponent(g.id) + '/' + encodeURIComponent(fn));
        if (!r.ok) throw new Error(fn + ' — server answered ' + r.status);
        const blob = await r.blob();
        loaded.push(new File([blob], fn, { type: blob.type || 'application/octet-stream' }));
      }
      sampleFiles = loaded;
      setFiles(loaded);
      if (g.label) {
        sampleSummary.innerHTML = '<span class="chip-plain" title="' + esc(g.files.join(', ')) + '">' +
          '<i class="fas fa-database" aria-hidden="true"></i> ' + esc(g.label) +
          ' &middot; ' + g.files.length + ' files</span>';
        sampleSummary.classList.add('show');
      }
      showToast('success', 'Sample loaded', 'Check the options in step 3, then choose Generate Report.');
    } catch (e) {
      const msg = String(e);
      showError(msg);
      showToast('error', 'Could not load sample', msg);
    } finally {
      btn.disabled = false;
      btn.innerHTML = '<i class="fas fa-play" aria-hidden="true"></i> Load sample';
    }
  });

  downloadSampleBtn.addEventListener('click', () => {
    const gid = sampleSelect.value;
    if (!gid) {
      showToast('error', 'No sample selected', 'Pick a bundled sample dataset first.');
      return;
    }
    window.location.href = '/api/sample-download/' + encodeURIComponent(gid);
  });


  // -------------------------------------------------------------------------
  // Report-type switching + review-period picker
  // -------------------------------------------------------------------------
  document.querySelectorAll('input[name="report_type"]').forEach(r => {
    r.addEventListener('change', () => applyReportType(r.value));
  });

  document.querySelectorAll('input[name="export_period"]').forEach(r => {
    r.addEventListener('change', () => {
      const wrap = $('semi-half-wrap');
      const semi = r.value === 'semi';
      wrap.style.display = semi ? 'flex' : 'none';
      if (semi) wrap.removeAttribute('aria-hidden');
      else wrap.setAttribute('aria-hidden', 'true');
    });
  });

  // -------------------------------------------------------------------------
  // Run: upload (XHR with real progress) + poll (with timeout & backoff)
  // -------------------------------------------------------------------------
  function uploadWithProgress(fd) {
    return new Promise((resolve, reject) => {
      const xhr = new XMLHttpRequest();
      xhr.open('POST', '/api/run');
      xhr.upload.onprogress = e => {
        if (e.lengthComputable) {
          const frac = e.loaded / e.total;
          setUploadProgress(frac);
          setStage('Uploading files… ' + Math.round(frac * 100) + '%');
        }
      };
      xhr.onload = () => {
        let data = null;
        try { data = JSON.parse(xhr.responseText); } catch (e) { /* non-JSON */ }
        resolve({ ok: xhr.status >= 200 && xhr.status < 300, status: xhr.status, data });
      };
      xhr.onerror = () => reject(new Error('network error — the server did not respond'));
      xhr.onabort = () => reject(new Error('upload aborted'));
      xhr.send(fd);
    });
  }

  function pollStatus(statusUrl, onLogLine) {
    return new Promise((resolve, reject) => {
      let shown = 0;
      let failures = 0;
      let polls = 0;
      const timer = setInterval(async () => {
        polls += 1;
        if (polls > POLL_MAX_POLLS) {
          clearInterval(timer);
          reject(new Error('the job is still running after 30 minutes — check the server logs'));
          return;
        }
        let s = null;
        try {
          const r = await fetch(statusUrl);
          if (r.status === 404) {
            clearInterval(timer);
            reject(new Error('the job expired or the server restarted — upload the files again'));
            return;
          }
          s = await r.json();
          failures = 0;
        } catch (e) {
          failures += 1;
          if (failures >= POLL_MAX_FAILURES) {
            clearInterval(timer);
            reject(new Error('lost contact with the server while waiting for the report'));
          }
          return; // transient hiccup: keep polling
        }
        const lines = s.log || [];
        for (; shown < lines.length; shown++) {
          const cls = classifyLine(lines[shown]);
          appendLog(lines[shown], cls);
          const stage = stageForLogLine(lines[shown]);
          if (stage) setStage(stage);
          const frac = fractionForLogLine(lines[shown]);
          const state = stageForLogLine(lines[shown]);
          if (frac != null) setBuildProgress(frac, state);
          if (onLogLine) onLogLine(lines[shown]);
        }
        if (s.state === 'done' || s.state === 'error') {
          clearInterval(timer);
          resolve(s);
        }
      }, POLL_INTERVAL_MS);
    });
  }

  // Client-side validation mirrors the server rules so mistakes surface
  // instantly instead of after a full upload.
  const ALLOWED_EXT = /\.(xlsx?|xlsm|xlsb|csv|zip|xls)$/i;

  function validateSelection(rt) {
    const files = filesInput.files;
    if (!files.length && !sampleFiles.length) {
      // A KNBS install refreshes the cross-check series; it is not an input,
      // so "Generate report" with nothing uploaded needs to say why.
      if (activeFetch && activeFetch.install_url) {
        return 'The installed KNBS snapshot is a cross-check, not a report ' +
          'input. Upload your ITC files to generate this report' +
          (rt === 'services'
            ? ', or fetch UNCTADstat services data and build from that fetch'
            : '') + '.';
      }
      return 'Please select files before generating the report.';
    }
    for (const f of files) {
      if (!ALLOWED_EXT.test(f.name)) {
        return 'Unsupported file type: ' + f.name + ' — use .xlsx, .xls, .xlsb, .csv or .zip.';
      }
    }
    const topN = parseInt($('top').value, 10);
    if (!(topN >= 1 && topN <= 50)) {
      return 'The "top" value must be between 1 and 50.';
    }
    if (rt === 'goods' && !sampleFiles.length && files.length < 4 && !/\.zip$/i.test(files[0]?.name || '')) {
      appendLog('Note: only ' + files.length + ' file(s) selected — a full goods run usually needs ' +
        'six raw ITC downloads (or a zip / the seven ready-made tables).', 'warn');
    }
    return '';
  }

  runBtn.addEventListener('click', async () => {
    clearError();
    clearResult();
    let files = filesInput.files;
    if (!files.length && sampleFiles.length) {
      const dt = new DataTransfer();
      sampleFiles.forEach(f => dt.items.add(f));
      files = filesInput.files = dt.files;
      currentFiles = [...sampleFiles];
      renderChips();
    }
    const rt = document.querySelector('input[name="report_type"]:checked').value;
    const invalid = validateSelection(rt);
    if (invalid) {
      showError(invalid);
      appendLog(invalid, 'err');
      return;
    }

    const fd = new FormData();
    fd.append('config', rt === 'product' ? productCfgSelect.value
      : (rt === 'services' ? 'services_world' : '__auto__'));
    fd.append('top', parseInt($('top').value, 10) || 10);
    fd.append('report_type', rt);
    let periodVal = '';
    if (rt === 'quarterly') {
      // No default on purpose: a silent full-year run when only one quarter
      // was wanted is exactly what the picker exists to prevent.
      const pr = document.querySelector('input[name="export_period"]:checked');
      if (!pr) {
        showError('Please choose a review period: Q1-Q4, Semi-Annual or Annual.');
        appendLog('Please choose a review period first.', 'err');
        return;
      }
      periodVal = pr.value;
      if (periodVal === 'semi') {
        periodVal = 'semi' +
          document.querySelector('input[name="export_half"]:checked').value;
      }
    }
    fd.append('period', periodVal);
    let zips = 0, others = 0;
    for (const f of files) {
      if (f.name.toLowerCase().endsWith('.zip')) { zips++; fd.append('zipfile', f); }
      else { others++; fd.append('files', f); }
    }

    const btn = runBtn;
    const typeRadios = document.querySelectorAll('input[name="report_type"]');
    btn.disabled = true;
    btn.setAttribute('aria-busy', 'true');
    btn.classList.remove('btn-attention');
    runCta.classList.remove('show');
    typeRadios.forEach(r => { r.disabled = true; });
    btn.innerHTML = '<i class="fas fa-spinner fa-spin" aria-hidden="true"></i> Generating…';
    logEl.innerHTML = '';
    appendLog(`Uploading ${others} file(s), ${zips} zip(s)…`);
    startProgress('Uploading files…');

    try {
      const up = await uploadWithProgress(fd);
      if (!up.ok) {
        const detail = normalizeDetail(up.data && up.data.detail) ||
          'Request failed with status ' + up.status;
        appendLog('ERROR: ' + detail, 'err');
        showError(detail);
        showToast('error', 'Generation failed', detail);
        return;
      }
      const data = up.data;
      setStage('Job accepted - building in the background.');
      setBuildProgress(0.05);
      rememberJob(data.job_id, data.job_token, {
        mode: data.mode,
        report_type: rt,
      });

      // The server accepted the upload and builds the report in the
      // background; follow progress by polling the job status.
      appendLog(`Job accepted (${data.job_id}) - building in the background.`);
      let final;
      try {
        final = await pollStatus(withAccess(data.status_url, {
          token: data.job_token,
        }));
      } catch (pollErr) {
        appendLog('ERROR: ' + pollErr.message, 'err');
        showError(pollErr.message);
        showToast('error', 'Lost the job', pollErr.message);
        return;
      }

      if (final.state === 'error') {
        appendLog('ERROR: ' + (final.detail || 'unknown error'), 'err');
        showError(final.detail || 'The job failed.');
        showToast('error', 'Generation failed', final.detail || 'The job failed.');
        return;
      }
      appendLog('Done.', 'ok');
      setBuildProgress(1);
      const access = { token: jobStore[final.job_id]?.token || '' };
      applyFinalResult(final, access, { notify: true, scroll: true });
    } catch (e) {
      appendLog('ERROR: ' + e + ' — the server did not respond. Reload this page'
        + ' and try again. If it keeps happening, restart the app and reload.', 'err');
      showError('Server did not respond. Please check the connection and try again.');
      showToast('error', 'Server did not respond', 'Please check the connection and try again.');
    } finally {
      stopProgress();
      btn.disabled = false;
      btn.removeAttribute('aria-busy');
      typeRadios.forEach(r => { r.disabled = false; });
      btn.innerHTML = '<i class="fas fa-play" aria-hidden="true"></i> Generate Report';
    }
  });

  // -------------------------------------------------------------------------
  // Platform UI: stored job ownership, previews, quality, sharing, comments
  // -------------------------------------------------------------------------
  function readJobStore() {
    try {
      const raw = localStorage.getItem(TOKENS_KEY);
      const parsed = raw ? JSON.parse(raw) : {};
      return parsed && typeof parsed === 'object' ? parsed : {};
    } catch (e) {
      return {};
    }
  }

  function persistJobStore() {
    try { localStorage.setItem(TOKENS_KEY, JSON.stringify(jobStore)); }
    catch (e) { /* storage may be unavailable in private mode */ }
  }

  function rememberJob(jobId, token, extra) {
    if (!jobId) return;
    const previous = jobStore[jobId] || {};
    jobStore[jobId] = Object.assign({}, previous, extra || {}, {
      token: token || previous.token || '',
      created: previous.created || (Date.now() / 1000),
      state: (extra && extra.state) || previous.state || 'queued',
    });
    persistJobStore();
    renderJobRows();
  }

  function updateJobState(jobId, state, extra) {
    if (!jobStore[jobId]) return;
    jobStore[jobId] = Object.assign({}, jobStore[jobId], { state }, extra || {});
    persistJobStore();
  }

  function accessQuery(access) {
    const params = new URLSearchParams();
    if (access && access.share) params.set('share', access.share);
    else if (access && access.token) params.set('token', access.token);
    const text = params.toString();
    return text ? '?' + text : '';
  }

  function withAccess(url, access) {
    if (!url) return url;
    const query = accessQuery(access).slice(1);
    if (!query) return url;
    return url + (url.includes('?') ? '&' : '?') + query;
  }

  function ensureAccess(url, access) {
    if (!url || !access || (!access.token && !access.share)) return url;
    if (/[?&](share|token)=/.test(url)) return url;
    return withAccess(url, access);
  }

  async function apiJson(url, options) {
    const opts = Object.assign({}, options || {});
    const token = opts.token;
    delete opts.token;
    opts.headers = Object.assign({ Accept: 'application/json' },
      opts.headers || {});
    if (token) opts.headers['X-Job-Token'] = token;
    if (opts.body && typeof opts.body !== 'string') {
      opts.body = JSON.stringify(opts.body);
      opts.headers['Content-Type'] = 'application/json';
    }
    const r = await fetch(url, opts);
    let data = null;
    try { data = await r.json(); } catch (e) { /* non-JSON error */ }
    if (!r.ok) {
      throw new Error(normalizeDetail(data && data.detail) ||
        ('Request failed with status ' + r.status));
    }
    return data;
  }

  function fmtDate(epoch) {
    if (!epoch) return '';
    const date = new Date(epoch * 1000);
    return isNaN(date.getTime()) ? '' : date.toLocaleString();
  }

  function compactNumber(value) {
    const n = Number(value);
    if (!isFinite(n)) return String(value == null ? '' : value);
    try {
      return new Intl.NumberFormat('en',
        { notation: 'compact', maximumFractionDigits: 1 }).format(n);
    } catch (e) {
      return n.toLocaleString();
    }
  }

  function modeLabelFor(mode) {
    const isProduct = mode === 'product' || mode === 'crafts';
    const isCompetitor = mode === 'competitor';
    const isExport = (mode || '').startsWith('quarterly');
    if (isCompetitor) return 'Competitor analysis';
    if (isProduct) return 'Product profile';
    if (isExport) return 'Export performance (raw monthly extracts)';
    if (mode === 'raw') return 'Built from raw ITC downloads';
    if ((mode || '').startsWith('services')) {
      return 'Built from raw services downloads';
    }
    return 'Ready-made tables';
  }

  function applyAccessMode(access) {
    sharedMode = !!(access && access.share);
    $('share-banner').style.display = sharedMode ? 'flex' : 'none';
    if (sharedMode) {
      ['analysis-card', 'upload-card', 'sources-card', 'options-card',
        'jobs-card'].forEach(id => {
        const el = $(id);
        if (el) el.style.display = 'none';
      });
    }
    if (commentCompose) {
      commentCompose.style.display = sharedMode ? 'none' : 'flex';
    }
    if (shareBtn) shareBtn.style.display = sharedMode ? 'none' : 'inline-flex';
    if (shareCreateBtn) shareCreateBtn.disabled = sharedMode;
  }

  function applyFinalResult(final, access, opts) {
    opts = opts || {};
    currentJob = { id: final.job_id, access: access || {} };
    setStage('Report ready!');
    setBuildProgress(1);
    const isProduct = final.mode === 'product' || final.mode === 'crafts';
    const isCompetitor = final.mode === 'competitor';
    const modeLabel = modeLabelFor(final.mode);
    $('result-meta').innerHTML =
      '<i class="fas fa-check-circle" aria-hidden="true"></i> ' +
      esc(modeLabel) + ' &middot; job ' + esc(final.job_id || '') +
      (final.report_version ? ' &middot; v' + esc(final.report_version) : '');
    $('dl-report').href = final.report_url || '#';
    $('dl-report').innerHTML = isCompetitor
      ? '<i class="fas fa-file-word" aria-hidden="true"></i> Download Table (.docx)'
      : '<i class="fas fa-file-word" aria-hidden="true"></i> Download Word Report (.docx)';
    $('dl-tables').href = isCompetitor ? (final.png_url || '#')
      : (isProduct ? (final.product_tables_url || '#') : (final.tables_url || '#'));
    $('dl-tables').innerHTML = isCompetitor
      ? '<i class="fas fa-image" aria-hidden="true"></i> Download Table Picture (.png)'
      : (isProduct
        ? '<i class="fas fa-file-excel" aria-hidden="true"></i> Download Excel Deliverable'
        : '<i class="fas fa-file-archive" aria-hidden="true"></i> Download Tables (.zip)');
    const dlXlsx = $('dl-xlsx');
    if (isCompetitor && final.xlsx_url) {
      dlXlsx.href = final.xlsx_url;
      dlXlsx.innerHTML = '<i class="fas fa-file-excel" aria-hidden="true"></i> Download Table (.xlsx)';
      dlXlsx.style.display = 'inline-flex';
    } else {
      dlXlsx.style.display = 'none';
    }
    shareBtn.style.display = sharedMode ? 'none' : 'inline-flex';
    $('result').style.display = 'block';
    $('result').classList.add('showing');
    if (final.job_id && !sharedMode) {
      updateJobState(final.job_id, 'done', {
        mode: final.mode, report_type: final.report_type,
      });
    }
    loadInsight(final, access || {});
    if (opts.notify) {
      showToast('success', 'Your report is ready!',
        'Downloads, preview, quality checks and sharing are available below.');
    }
    if (opts.scroll) {
      setTimeout(() => {
        document.querySelector('.log-card')
          .scrollIntoView({ behavior: 'smooth', block: 'start' });
      }, 400);
    }
    if (!sharedMode) renderJobRows();
  }

  function loadInsight(final, access) {
    insightCard.style.display = 'block';
    applyAccessMode(access);
    const jobId = final.job_id;
    loadPreview(ensureAccess(final.preview_url ||
      ('/api/preview/' + jobId), access), access);
    loadQuality(ensureAccess(final.quality_url ||
      ('/api/quality/' + jobId), access), access);
    loadComments(jobId, access);
    shareLinkInput.value = '';
  }

  async function loadPreview(url, access) {
    previewChart.innerHTML = '';
    previewTableWrap.innerHTML =
      '<div class="jobs-empty">Loading preview&hellip;</div>';
    try {
      const data = await apiJson(url, access);
      previewFiles = (data && data.files) || [];
      if (!previewFiles.length) {
        previewFileSelect.innerHTML = '';
        previewTableWrap.innerHTML =
          '<div class="jobs-empty">No workbook preview is available for this report type.</div>';
        return;
      }
      previewFileSelect.innerHTML = previewFiles.map((file, index) =>
        '<option value="' + index + '">' + esc(file.name || ('Workbook ' + (index + 1))) +
        '</option>').join('');
      renderPreviewFile(0);
    } catch (e) {
      previewFileSelect.innerHTML = '';
      previewTableWrap.innerHTML =
        '<div class="jobs-empty">' + esc(e.message) + '</div>';
    }
  }

  function renderPreviewFile(index) {
    const file = previewFiles[index];
    if (!file) return;
    if (file.error) {
      previewChart.innerHTML = '';
      previewTableWrap.innerHTML =
        '<div class="jobs-empty">' + esc(file.error) + '</div>';
      return;
    }
    const series = file.series || [];
    if (series.length) {
      previewChart.innerHTML = series.map(s => {
        const years = s.years || [];
        const values = (s.values || []).map(v => Number(v) || 0);
        const max = Math.max.apply(null,
          values.map(v => Math.abs(v)).concat([1]));
        const bars = years.map((year, i) => {
          const value = values[i] || 0;
          const height = Math.max(3,
            Math.round(Math.abs(value) / max * 100));
          return '<div class="preview-bar" title="' + esc(year + ': ' + compactNumber(value)) + '">' +
            '<span class="bar-value">' + esc(compactNumber(value)) + '</span>' +
            '<span class="bar" style="height:' + height + 'px"></span>' +
            '<span class="bar-label">' + esc(year) + '</span></div>';
        }).join('');
        return '<div><div class="preview-series-label">' + esc(s.label || file.name) +
          '</div><div class="preview-bars">' + bars + '</div></div>';
      }).join('');
    } else {
      previewChart.innerHTML = '';
    }

    const columns = file.columns || [];
    const rows = file.rows || [];
    const maxCols = Math.min(columns.length, PREVIEW_COLS);
    if (!columns.length && !rows.length) {
      previewTableWrap.innerHTML =
        '<div class="jobs-empty">This workbook has no readable rows.</div>';
      return;
    }
    const head = columns.slice(0, maxCols).map(c =>
      '<th>' + esc(c == null ? '' : c) + '</th>').join('');
    const body = rows.slice(0, PREVIEW_ROWS).map(row => {
      const cells = [];
      for (let i = 0; i < maxCols; i++) {
        const value = row[i];
        const numeric = typeof value === 'number';
        cells.push('<td' + (numeric ? ' class="num"' : '') + '>' +
          esc(numeric ? value.toLocaleString(undefined,
            { maximumFractionDigits: 2 }) : (value == null ? '' : value)) +
          '</td>');
      }
      return '<tr>' + cells.join('') + '</tr>';
    }).join('');
    previewTableWrap.innerHTML = '<table><thead><tr>' + head +
      '</tr></thead><tbody>' + body + '</tbody></table>';
  }

  async function loadQuality(url, access) {
    qualitySummary.className = 'badge';
    qualitySummary.textContent = 'Checking';
    qualityList.innerHTML = '<div class="jobs-empty">Loading quality checks&hellip;</div>';
    qualityFiles.innerHTML = '';
    try {
      const data = await apiJson(url, access);
      const summary = data.summary || {};
      const errors = summary.errors || 0;
      const warnings = summary.warnings || 0;
      const files = summary.files || 0;
      const cls = errors ? 'error' : (warnings ? 'warn' : 'ok');
      qualitySummary.className = 'badge ' + cls;
      qualitySummary.textContent = errors + ' errors \u00b7 ' +
        warnings + ' warnings \u00b7 ' + files + ' inputs';
      const checks = data.checks || [];
      qualityList.innerHTML = checks.length ? checks.map(check => {
        const level = ['ok', 'warn', 'error'].includes(check.level)
          ? check.level : 'ok';
        const icon = level === 'error' ? 'fa-circle-xmark'
          : (level === 'warn' ? 'fa-triangle-exclamation' : 'fa-circle-check');
        return '<div class="quality-item ' + level + '">' +
          '<i class="fas ' + icon + '" aria-hidden="true"></i>' +
          '<span>' + esc(check.message || '') + '</span></div>';
      }).join('') :
        '<div class="quality-item ok"><i class="fas fa-circle-check" aria-hidden="true"></i><span>No quality issues reported.</span></div>';
      const inputs = data.files || [];
      qualityFiles.innerHTML = inputs.length ? inputs.map(item => {
        const hash = item.sha256 ? String(item.sha256).slice(0, 16) + '&hellip;' : '';
        return '<div title="' + esc(item.sha256 || '') + '"><b>' +
          esc(item.name || 'input') + '</b> &middot; ' +
          esc(fmtSize(item.bytes || 0)) +
          (item.source ? ' &middot; ' + esc(item.source) : '') +
          (hash ? '<br><span class="hash">sha256 ' + hash + '</span>' : '') +
          '</div>';
      }).join('') :
        '<div class="jobs-empty">No provenance fingerprints recorded.</div>';
    } catch (e) {
      qualitySummary.className = 'badge error';
      qualitySummary.textContent = 'Unavailable';
      qualityList.innerHTML =
        '<div class="quality-item error"><i class="fas fa-circle-xmark" aria-hidden="true"></i><span>' +
        esc(e.message) + '</span></div>';
    }
  }

  async function loadComments(jobId, access) {
    if (!jobId) return;
    try {
      const data = await apiJson(
        ensureAccess('/api/jobs/' + jobId + '/comments', access), access);
      renderComments((data && data.comments) || []);
    } catch (e) {
      commentsList.innerHTML =
        '<div class="comments-empty">' + esc(e.message) + '</div>';
      commentsCount.textContent = '';
    }
  }

  function renderComments(comments) {
    commentsCount.textContent = comments.length + ' total';
    commentsList.innerHTML = comments.length ? comments.map(comment =>
      '<div class="comment"><div class="comment-head"><span class="comment-author">' +
      esc(comment.author || 'Anonymous') + '</span><span class="comment-time">' +
      esc(fmtDate(comment.created)) + '</span></div><div class="comment-body">' +
      esc(comment.body || '') + '</div></div>').join('') :
      '<div class="comments-empty">No comments yet.</div>';
  }

  async function addComment() {
    if (!currentJob || sharedMode) return;
    const body = commentBody.value.trim();
    if (!body) {
      showToast('error', 'Comment is empty', 'Write a note before adding it.');
      return;
    }
    commentAddBtn.disabled = true;
    try {
      await apiJson('/api/jobs/' + currentJob.id + '/comments', {
        method: 'POST',
        token: currentJob.access.token,
        body: { body, author: commentAuthor.value.trim() },
      });
      commentBody.value = '';
      await loadComments(currentJob.id, currentJob.access);
      showToast('success', 'Comment added', 'It is now visible on this report.');
    } catch (e) {
      showToast('error', 'Could not add comment', e.message);
    } finally {
      commentAddBtn.disabled = false;
    }
  }

  async function createShareLink() {
    if (!currentJob || sharedMode) return;
    shareCreateBtn.disabled = true;
    try {
      const data = await apiJson('/api/share/' + currentJob.id, {
        method: 'POST', token: currentJob.access.token,
      });
      shareLinkInput.value = data.share_url || '';
      showToast('success', 'Share link created',
        'The read-only link expires ' + fmtDate(data.expires_at) + '.');
    } catch (e) {
      showToast('error', 'Could not create share link', e.message);
    } finally {
      shareCreateBtn.disabled = false;
    }
  }

  async function copyShareLink() {
    if (!shareLinkInput.value) await createShareLink();
    if (!shareLinkInput.value) return;
    try {
      await navigator.clipboard.writeText(shareLinkInput.value);
      showToast('success', 'Link copied', 'Paste it wherever you need to share.');
    } catch (e) {
      shareLinkInput.select();
      document.execCommand('copy');
      showToast('success', 'Link copied', 'Paste it wherever you need to share.');
    }
  }

  // -------------------------------------------------------------------------
  // Open-data sources
  // -------------------------------------------------------------------------
  async function loadSources() {
    if (!sourceSelect) return;
    try {
      const data = await apiJson('/api/sources');
      sourceCatalog = (data && data.sources) || [];
      if (!sourceCatalog.length) {
        sourceSelect.innerHTML = '<option value="">No sources configured</option>';
        sourceDescription.textContent = 'No sanctioned data connectors are available on this server.';
        return;
      }
      renderSourceOptions();
    } catch (e) {
      sourceSelect.innerHTML = '<option value="">Could not load sources</option>';
      sourceDescription.textContent = e.message;
    }
  }

  function currentReportType() {
    const checked = document.querySelector('input[name="report_type"]:checked');
    return checked ? checked.value : 'goods';
  }

  // A reference source serves every report that reads it; a pipeline source
  // only the tab it builds for.
  function sourcesForTab() {
    const rt = currentReportType();
    return sourceCatalog.filter(source => {
      const types = source.report_types || [];
      return types.includes(rt) || types.includes('reference');
    });
  }

  function renderSourceOptions() {
    const list = sourcesForTab();
    if (!list.length) {
      sourceSelect.innerHTML = '<option value="">No source for this report</option>';
      sourceSelect.value = '';
      renderSource();
      return;
    }
    const keep = sourceChosen ? list.find(s => s.id === sourceSelect.value) : null;
    sourceSelect.innerHTML = list.map(source =>
      '<option value="' + esc(source.id) + '">' + esc(source.label) +
      (source.experimental ? ' (experimental)' : '') +
      (source.requires_credentials && !source.credentials_ready
        ? ' - credentials needed' : '') + '</option>').join('');
    const preferred = keep ||
      list.find(s => s.id === 'unctad' && s.credentials_ready) ||
      list.find(s => s.credentials_ready && !s.experimental) ||
      list[0];
    sourceSelect.value = preferred.id;
    renderSource();
  }

  function renderSource() {
    activeSource = sourceCatalog.find(s => s.id === sourceSelect.value) || null;
    sourceResult.style.display = 'none';
    activeFetch = null;
    if (!activeSource) {
      sourceDescription.textContent = '';
      sourceParams.innerHTML = '';
      sourceFetchBtn.disabled = true;
      sourceStatus.textContent = '';
      return;
    }
    const tags = [];
    if (activeSource.requires_credentials) {
      tags.push(activeSource.credentials_ready
        ? 'credentials configured'
        : 'credentials not configured');
    }
    if (activeSource.experimental) tags.push('experimental');
    tags.push('output: ' + (activeSource.report_types || []).join(', '));
    // A reference source builds nothing, so say which reports read it before
    // the user spends a fetch finding out.
    const usedBy = activeSource.used_by || [];
    sourceDescription.innerHTML = esc(activeSource.description || '') +
      (tags.length ? ' <span class="badge">' +
        tags.map(tag => esc(tag)).join(' &middot; ') +
        '</span>' : '') +
      (usedBy.length ? '<br><strong>Used by:</strong><ul class="source-notes">' +
        usedBy.map(item => '<li>' + esc(item) + '</li>').join('') +
        '</ul>' : '');
    sourceParams.innerHTML = (activeSource.params || []).map(param => {
      const id = 'source-p-' + param.id;
      let input;
      if (param.type === 'select' && Array.isArray(param.choices) && param.choices.length) {
        input = '<select id="' + esc(id) + '">' + param.choices.map(choice => {
          const value = typeof choice === 'object' ? choice.value : choice;
          const label = typeof choice === 'object' ? choice.label : choice;
          return '<option value="' + esc(value) + '"' +
            (String(value) === String(param.default) ? ' selected' : '') +
            '>' + esc(label) + '</option>';
        }).join('') + '</select>';
      } else {
        input = '<input id="' + esc(id) + '" type="' +
          (param.type === 'number' ? 'number' : 'text') + '" value="' +
          esc(param.default == null ? '' : param.default) + '"' +
          (param.required ? ' required' : '') + '>';
      }
      return '<div class="field"><label for="' + esc(id) + '">' +
        esc(param.label) + (param.required ? ' *' : '') + '</label>' +
        input + (param.help ? '<span class="help">' + esc(param.help) +
          '</span>' : '') + '</div>';
    }).join('');
    const blocked = activeSource.requires_credentials &&
      !activeSource.credentials_ready;
    sourceFetchBtn.disabled = blocked;
    sourceStatus.className = 'source-status' + (blocked ? ' error' : '');
    sourceStatus.textContent = blocked
      ? 'This source needs server credentials before it can be fetched.'
      : '';
  }

  function sourceValues() {
    const values = {};
    (activeSource && activeSource.params || []).forEach(param => {
      const input = $('source-p-' + param.id);
      values[param.id] = input ? input.value : param.default;
    });
    return values;
  }

  async function fetchSource() {
    if (!activeSource || sourceFetchBtn.disabled) return;
    sourceFetchBtn.disabled = true;
    sourceStatus.className = 'source-status busy';
    sourceStatus.innerHTML = '<i class="fas fa-spinner fa-spin" aria-hidden="true"></i> Fetching ' +
      esc(activeSource.label) + '&hellip; this can take a minute.';
    try {
      const data = await apiJson('/api/sources/' + activeSource.id + '/fetch', {
        method: 'POST', body: sourceValues(),
      });
      activeFetch = data;
      renderSourceResult(data);
      const runnable = (activeSource.report_types || [])
        .some(t => t !== 'reference');
      sourceStatus.className = 'source-status';
      sourceStatus.textContent = 'Fetched ' + data.files.length +
        ' file(s). ' + (runnable
          ? 'Download them or build a report below.'
          : data.install_url
            ? 'Install the snapshot to refresh the cross-check series, or download it for inspection. KNBS alone cannot generate a report.'
            : 'Download them for inspection; this reference source cannot build a report directly.');
    } catch (e) {
      sourceStatus.className = 'source-status error';
      sourceStatus.textContent = e.message;
    } finally {
      const blocked = activeSource.requires_credentials &&
        !activeSource.credentials_ready;
      sourceFetchBtn.disabled = blocked;
    }
  }

  function renderSourceResult(data) {
    sourceResult.style.display = 'block';
    const notes = data.notes || [];
    sourceNotes.innerHTML = notes.length
      ? '<ul class="source-notes">' + notes.map(note =>
        '<li>' + esc(note) + '</li>').join('') + '</ul>'
      : '';
    sourceFiles.innerHTML = (data.files || []).map(file =>
      '<a href="' + esc(file.url) + '" download><i class="fas fa-file-arrow-down" aria-hidden="true"></i>' +
      esc(file.name) + ' <span>(' + esc(fmtSize(file.bytes || 0)) + ')</span></a>').join('');
    const runnable = (activeSource.report_types || []).some(t => t !== 'reference');
    sourceActions.innerHTML =
      (data.zip_url ? '<a class="btn btn-quiet btn-compact" href="' +
        esc(data.zip_url) + '"><i class="fas fa-file-zipper" aria-hidden="true"></i> Download all as ZIP</a>' : '') +
      (data.install_url
        ? '<button id="source-install" type="button" class="btn btn-secondary btn-compact">' +
          '<i class="fas fa-rotate" aria-hidden="true"></i> Install snapshot</button>'
        : '') +
      (runnable
        ? '<button id="source-run" type="button" class="btn btn-secondary btn-compact">' +
          '<i class="fas fa-play" aria-hidden="true"></i> Build report from this data</button>'
        : '<span class="hint"><i class="fas fa-circle-info" aria-hidden="true"></i> Reference dataset &mdash; ' +
          (data.install_url
            ? 'a cross-check layer, not a report input. Installing it refreshes the KNBS series listed above; it cannot generate a report on its own. '
            : 'download for inspection or validation. ') +
          'These files cannot be uploaded as report inputs.</span>');
    const run = $('source-run');
    if (run) run.addEventListener('click', runFetchedJob);
    const install = $('source-install');
    if (install) install.addEventListener('click', installSnapshot);
  }

  async function installSnapshot() {
    if (!activeFetch || !activeFetch.install_url || sharedMode) return;
    const button = $('source-install');
    if (button) button.disabled = true;
    sourceStatus.className = 'source-status busy';
    sourceStatus.innerHTML = '<i class="fas fa-spinner fa-spin" aria-hidden="true"></i> Installing the KNBS reference snapshot&hellip;';
    try {
      const data = await apiJson(activeFetch.install_url, { method: 'POST' });
      sourceStatus.className = 'source-status';
      sourceStatus.textContent = 'Installed ' + data.datasets + ' datasets (' +
        Number(data.observations).toLocaleString() + ' observations, taken ' +
        (data.generated_at || 'just now') + '). Reports built from your ITC ' +
        'files or a UNCTADstat fetch now cross-check against them.';
      appendLog('KNBS reference snapshot installed: ' + data.datasets +
        ' datasets, ' + data.observations + ' observations.', 'ok');
    } catch (e) {
      sourceStatus.className = 'source-status error';
      sourceStatus.textContent = e.message;
      appendLog('ERROR: ' + e.message, 'err');
      showToast('error', 'Could not install the snapshot', e.message);
    } finally {
      if (button) button.disabled = false;
    }
  }

  async function runFetchedJob() {
    if (!activeFetch || !activeSource || sharedMode) return;
    const reportType = (activeSource.report_types || [])
      .find(t => t !== 'reference') || 'goods';
    const run = $('source-run');
    if (run) run.disabled = true;
    sourceStatus.className = 'source-status busy';
    sourceStatus.textContent = 'Staging fetched files and starting the report build.';
    logEl.innerHTML = '';
    appendLog('Starting report from ' + activeSource.label + ' data.');
    startProgress('Staging fetched data.');
    try {
      const data = await apiJson('/api/fetches/' + activeFetch.fetch_id + '/run', {
        method: 'POST',
        body: {
          report_type: reportType,
          config: reportType === 'product' ? productCfgSelect.value : '__auto__',
          top: parseInt($('top').value, 10) || 20,
        },
      });
      rememberJob(data.job_id, data.job_token, {
        mode: data.mode, report_type: data.report_type,
      });
      appendLog('Job accepted (' + data.job_id + ').');
      const final = await pollStatus(withAccess(data.status_url,
        { token: data.job_token }));
      if (final.state === 'error') {
        throw new Error(final.detail || 'The job failed.');
      }
      appendLog('Done.', 'ok');
      applyFinalResult(final, { token: data.job_token },
        { notify: true, scroll: true });
      sourceStatus.className = 'source-status';
      sourceStatus.textContent = 'Report built from fetched data.';
    } catch (e) {
      sourceStatus.className = 'source-status error';
      sourceStatus.textContent = e.message;
      appendLog('ERROR: ' + e.message, 'err');
      showToast('error', 'Could not build from fetched data', e.message);
    } finally {
      stopProgress();
      if (run) run.disabled = false;
    }
  }

  // -------------------------------------------------------------------------
  // Recent jobs (stored ownership tokens never appear in the DOM)
  // -------------------------------------------------------------------------
  function renderJobRows(states) {
    if (!jobsList) return;
    states = states || {};
    const entries = Object.entries(jobStore)
      .sort((a, b) => (b[1].created || 0) - (a[1].created || 0))
      .slice(0, 8);
    if (!entries.length) {
      jobsList.innerHTML =
        '<div class="jobs-empty">No jobs in this browser yet.</div>';
      return;
    }
    jobsList.innerHTML = entries.map(([id, record]) => {
      const state = states[id] || record.state || 'done';
      const label = (record.report_type || record.mode || 'report');
      return '<div class="job-row">' +
        '<button type="button" data-job-id="' + esc(id) + '">' +
        '<span class="job-id">' + esc(id.slice(0, 8)) + '&hellip;</span> ' +
        esc(label) + ' &middot; ' + esc(fmtDate(record.created)) +
        '</button><span class="job-state ' + esc(state) + '">' +
        esc(state) + '</span></div>';
    }).join('');
  }

  async function refreshRecentJobs() {
    if (sharedMode || !jobsList) return;
    const entries = Object.entries(jobStore)
      .sort((a, b) => (b[1].created || 0) - (a[1].created || 0))
      .slice(0, 8);
    if (!entries.length) {
      renderJobRows();
      return;
    }
    jobsList.innerHTML = '<div class="jobs-empty">Checking recent jobs&hellip;</div>';
    const states = {};
    await Promise.all(entries.map(async ([id, record]) => {
      try {
        const status = await apiJson(
          withAccess('/api/status/' + id, { token: record.token }),
          { token: record.token });
        states[id] = status.state || 'unknown';
      } catch (e) {
        states[id] = 'error';
      }
    }));
    renderJobRows(states);
  }

  async function openRecentJob(jobId) {
    const record = jobStore[jobId];
    if (!record || sharedMode) return;
    clearError();
    logEl.innerHTML = '';
    appendLog('Opening job ' + jobId.slice(0, 8) + '&hellip;');
    startProgress('Loading saved job.');
    try {
      const statusUrl = withAccess('/api/status/' + jobId,
        { token: record.token });
      let final = await apiJson(statusUrl, { token: record.token });
      if (final.state === 'running' || final.state === 'queued') {
        final = await pollStatus(statusUrl);
      }
      if (final.state === 'error') {
        throw new Error(final.detail || 'The job failed.');
      }
      applyFinalResult(final, { token: record.token },
        { notify: false, scroll: true });
    } catch (e) {
      showError(e.message);
      appendLog('ERROR: ' + e.message, 'err');
      showToast('error', 'Could not open job', e.message);
    } finally {
      stopProgress();
    }
  }

  async function initSharedView() {
    const match = location.pathname.match(/^\/share\/([0-9a-f]{32})\/?$/);
    const params = new URLSearchParams(location.search);
    const share = params.get('share');
    if (!match || !share) return false;
    const jobId = match[1];
    sharedMode = true;
    applyAccessMode({ share });
    document.title = 'Shared trade report | KEPROBA';
    clearError();
    logEl.innerHTML = '';
    appendLog('Opening shared report ' + jobId.slice(0, 8) + '&hellip;');
    startProgress('Loading shared report.');
    try {
      const statusUrl = withAccess('/api/status/' + jobId, { share });
      let final = await apiJson(statusUrl);
      if (final.state === 'running' || final.state === 'queued') {
        final = await pollStatus(statusUrl);
      }
      if (final.state === 'error') {
        throw new Error(final.detail || 'This shared job failed.');
      }
      applyFinalResult(final, { share }, { notify: false, scroll: false });
      appendLog('Shared report loaded.', 'ok');
    } catch (e) {
      showError(e.message);
      appendLog('ERROR: ' + e.message, 'err');
      showToast('error', 'Could not open shared report', e.message);
    } finally {
      stopProgress();
    }
    return true;
  }

  // -------------------------------------------------------------------------
  // Platform event wiring
  // -------------------------------------------------------------------------
  if (sourceSelect) sourceSelect.addEventListener('change', () => {
    sourceChosen = true;
    renderSource();
  });
  if (sourceFetchBtn) sourceFetchBtn.addEventListener('click', fetchSource);
  if (previewFileSelect) {
    previewFileSelect.addEventListener('change',
      () => renderPreviewFile(parseInt(previewFileSelect.value, 10) || 0));
  }
  if (shareBtn) shareBtn.addEventListener('click', () => {
    if (insightCard.style.display === 'none') return;
    insightCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
    createShareLink();
  });
  if (shareCreateBtn) shareCreateBtn.addEventListener('click', createShareLink);
  if (shareCopyBtn) shareCopyBtn.addEventListener('click', copyShareLink);
  if (commentAddBtn) commentAddBtn.addEventListener('click', addComment);
  if (jobsRefreshBtn) jobsRefreshBtn.addEventListener('click', refreshRecentJobs);
  if (jobsList) {
    jobsList.addEventListener('click', e => {
      const button = e.target.closest('[data-job-id]');
      if (button) openRecentJob(button.dataset.jobId);
    });
  }

  // Initial panel state matches the default (goods) radio.
  const initialRadio = document.querySelector('input[name="report_type"]:checked');
  if (initialRadio) initialRadio.dispatchEvent(new Event('change', { bubbles: true }));
  else applyReportType('goods');
  loadProductProfiles();
  loadSamples();
  renderJobRows();
  initSharedView().then(isShared => {
    if (!isShared) {
      loadSources();
      refreshRecentJobs();
    }
  });
})();
