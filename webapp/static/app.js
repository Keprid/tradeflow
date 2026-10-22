/* Trade Flow Report Generator — UI logic */
(function () {
  'use strict';

  const $ = id => document.getElementById(id);
  const logEl = $('log');
  const errorMsg = $('error-msg');
  const hint = $('upload-hint');
  const cfgSelect = $('cfg');
  const cfgWrapper = $('cfg-wrapper');
  const servicesInfo = $('services-info');
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
  const partnerWrap = $('goods-partner-wrap');
  const partnerInput = $('partner');
  const prepBtn = $('prep-checklist');
  const checklistBox = $('checklist');
  const promotionWrap = $('promotion-wrap');
  const topWrap = $('top-wrap');
  const dropzone = $('dropzone');
  const sampleSelect = $('sample');
  const loadSampleBtn = $('load-sample');
  const sampleSummary = $('sample-summary');

  const POLL_INTERVAL_MS = 2000;
  const POLL_MAX_FAILURES = 5;       // consecutive fetch errors tolerated
  const POLL_MAX_POLLS = 900;        // ~30 minutes at 2 s per poll

  let elapsedTimer = null;
  let currentFiles = [];             // mirrors filesInput.files, removable
  let checklistText = '';
  let checklistFileName = 'itc-checklist.txt';

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
    $('dl-promotion').style.display = 'none';
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

  function setBuildProgress(frac) {
    // 0..1 within the build phase (remaining 85%)
    progressFill.style.width = Math.round(15 + 85 * Math.min(1, Math.max(0, frac))) + '%';
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
    if (/Mode: per-product/i.test(t)) return 'Reading supplying-markets files…';
    if (/Country detected/i.test(t)) return 'Country detected — loading config…';
    if (/Building report|Building services report/i.test(t)) return 'Building the Word report…';
    if (/Building product profile|Building crafts/i.test(t)) return 'Building the product profile…';
    if (/tables written|workbooks written/i.test(t)) return 'Tables written — building the report…';
    if (/promotion workbook/i.test(t)) return 'Building the promotion workbook…';
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

  function showPanelFlex(el, visible) {
    el.style.display = visible ? 'flex' : 'none';
    if (visible) el.removeAttribute('aria-hidden');
    else el.setAttribute('aria-hidden', 'true');
  }

  const PANELS = {
    cfgWrapper: () => cfgWrapper,
    productCfgWrapper: () => productCfgWrapper,
    servicesInfo: () => servicesInfo,
    quarterlyInfo: () => quarterlyInfo,
    exportPeriodWrap: () => exportPeriodWrap,
    partnerWrap: () => partnerWrap,
    promotionWrap: () => promotionWrap,
    productInfo: () => productInfo,
    competitorInfo: () => competitorInfo,
    topWrap: () => topWrap,
  };

  function setPanels(state) {
    for (const [k, get] of Object.entries(PANELS)) {
      const want = !!state[k];
      if (k === 'promotionWrap') showPanelFlex(get(), want);
      else showPanel(get(), want);
    }
  }

  const MODE_HINTS = {
    goods:
      '<i class="fas fa-info-circle" style="color:var(--red);" aria-hidden="true"></i> &nbsp;' +
      'Upload the <b>six raw ITC downloads</b> (or a <b>.zip</b> containing them), <i>or</i> the seven ready-made tables: ' +
      '<b>Table 1&ndash;6</b> + <b>Figure 1 Trade Balance.xlsx</b>. ' +
      '<b>.xls</b>, <b>.xlsb</b> and <b>.csv</b> files are converted automatically. ' +
      'Leave the country dropdown on <b>Auto-detect</b> unless the data could not be detected.',
    services:
      '<i class="fas fa-info-circle" style="color:var(--red);" aria-hidden="true"></i> &nbsp;' +
      'Upload the <b>seven raw ITC services files</b> (or a <b>.zip</b> containing them), ' +
      'e.g. <b>exported_services</b>, <b>importers_for</b>, <b>services_exported_by_Kenya</b> &hellip; ' +
      '<b>.xls</b> files are converted automatically.',
    quarterly:
      '<i class="fas fa-info-circle" style="color:var(--red);" aria-hidden="true"></i> &nbsp;' +
      'Upload the <b>raw monthly extracts</b> &mdash; the <b>KNBS</b> workbook and/or the <b>KRA</b> ' +
      '&hellip;<b>Exports by HS Destination&hellip;</b> / &hellip;<b>Imports by HS Origin&hellip;</b> files ' +
      '(up to 4 files, or a .zip) &mdash; then pick the <b>review period</b> below. ' +
      'Every sheet is read individually; pivot and summary sheets are ignored. ' +
      '<b>.xls</b> files are converted automatically.',
    product:
      '<i class="fas fa-info-circle" style="color:var(--red);" aria-hidden="true"></i> &nbsp;' +
      'Upload the <b>ITC matrix downloads</b> for a product family (or a <b>.zip</b> containing them). ' +
      'Expected filenames contain keywords like "exporting_economies", "kenyas_exports_to_world_by_product", etc. ' +
      '<b>.xls</b> files are converted automatically.',
    competitor:
      '<i class="fas fa-info-circle" style="color:var(--red);" aria-hidden="true"></i> &nbsp;' +
      'Run this <b>after the goods trade flow</b>. Upload the per-product ITC Trade Map ' +
      '<b>&ldquo;List of supplying markets for a product imported by &lt;partner&gt;&rdquo;</b> files, ' +
      'one per product (names like <code>nigerias-imports-from-world-by-exporter_0902.xlsx</code>), or a <b>.zip</b>. ' +
      'Produces <b>COMPETITOR ANALYSIS</b> as .png / .docx / .xlsx.',
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
      cfgSelect.value = '__auto__';
      setPanels({ cfgWrapper: true, partnerWrap: true, promotionWrap: true, topWrap: true });
    } else if (rt === 'services') {
      cfgSelect.value = 'services_world';
      setPanels({ servicesInfo: true, topWrap: true });
    } else if (rt === 'quarterly') {
      setPanels({ quarterlyInfo: true, exportPeriodWrap: true, topWrap: true });
    } else if (rt === 'product') {
      cfgSelect.value = '__auto__';
      setPanels({ productInfo: true, productCfgWrapper: true, topWrap: true });
    } else if (rt === 'competitor') {
      cfgSelect.value = '__auto__';
      setPanels({ competitorInfo: true });
    }
    if (rt !== 'goods') $('promotion').checked = false;
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
  // Config / profile / sample loading (with real error states)
  // -------------------------------------------------------------------------
  async function loadConfigs() {
    try {
      const r = await fetch('/api/configs');
      if (!r.ok) throw new Error('HTTP ' + r.status);
      const cfgs = await r.json();
      cfgSelect.innerHTML = '';
      for (const c of cfgs) {
        const o = document.createElement('option');
        o.value = c.id;
        o.textContent = c.year ? `${c.name} (year ${c.year})` : c.name;
        cfgSelect.appendChild(o);
      }
      if (!cfgs.length) {
        const o = document.createElement('option');
        o.value = '';
        o.textContent = 'No country configs found in config/';
        cfgSelect.appendChild(o);
      }
      cfgSelect.value = '__auto__';
    } catch (e) {
      cfgSelect.innerHTML = '';
      const o = document.createElement('option');
      o.value = '';
      o.textContent = 'Could not load country configs — reload to retry';
      cfgSelect.appendChild(o);
      showToast('error', 'Could not load configs', String(e));
    }
  }

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
      // Apply the recommended config when the relevant dropdown has it.
      const cfgSel = g.report_type === 'product' ? productCfgSelect : cfgSelect;
      if (g.config && [...cfgSel.options].some(o => o.value === g.config)) {
        cfgSel.value = g.config;
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
      btn.innerHTML = '<i class="fas fa-download" aria-hidden="true"></i> Load sample';
    }
  });

  // -------------------------------------------------------------------------
  // Goods partner checklist
  // -------------------------------------------------------------------------
  prepBtn.addEventListener('click', async () => {
    const p = (partnerInput.value || '').trim();
    if (!p) {
      showError('Enter a partner first, e.g. Saudi Arabia.');
      showToast('error', 'Partner required', 'Type a market name, e.g. Saudi Arabia.');
      return;
    }
    prepBtn.disabled = true;
    prepBtn.innerHTML = '<i class="fas fa-spinner fa-spin" aria-hidden="true"></i> Building…';
    try {
      const r = await fetch('/api/goods/partner-setup?partner=' + encodeURIComponent(p));
      const data = await r.json();
      if (!r.ok) { throw new Error(normalizeDetail(data.detail) || 'Request failed'); }

      const created = data.config.created ? ' (config created)' : ' (config reused)';
      let html = '<div style="font-size:13px;line-height:1.6;">' +
        '<b style="color:var(--green);">Checklist ready for ' + esc(data.partner) + '</b> &mdash; ' +
        'config <b>config/' + esc(data.config.id) + '.json</b>' + created + ' (report year ' + data.config.year + ').<br>' +
        '<i class="fas fa-circle-info" style="color:var(--red);" aria-hidden="true"></i> &nbsp;<span style="opacity:.9">' +
        esc(data.note) + '</span></div>';
      html += '<ol style="margin:10px 0 0 18px;font-size:13px;line-height:1.55;">';
      const plain = [
        'ITC Trade Map checklist for ' + data.partner + ' — config/' + data.config.id + '.json' + created,
        data.note,
        '',
      ];
      for (const it of data.checklist) {
        const pats = it.patterns.map(esc).join(' &nbsp;or&nbsp; ');
        html += '<li><b>' + esc(it.short) + '</b><div style="opacity:.85">' +
          '<i class="fas fa-magnifying-glass" style="color:var(--red);" aria-hidden="true"></i> ' + esc(it.query) +
          '<div class="hint" style="font-size:12px;"><b>Filename:</b> <code>' + pats + '</code></div>' +
          '</div></li>';
        plain.push('• ' + it.short + ' — ' + it.query + '\n  Filename: ' + it.patterns.join(' or '));
      }
      html += '</ol>';
      html += '<div class="checklist-actions">' +
        '<button type="button" class="btn-quiet" id="copy-checklist">' +
        '<i class="fas fa-copy" aria-hidden="true"></i> Copy checklist</button>' +
        '<button type="button" class="btn-quiet" id="download-checklist">' +
        '<i class="fas fa-download" aria-hidden="true"></i> Download .txt</button></div>';
      checklistBox.innerHTML = html;
      checklistBox.style.display = 'block';
      checklistText = plain.join('\n');
      checklistFileName = data.partner.replace(/[^a-z0-9]+/gi, '-').replace(/^-|-$/g, '').toLowerCase() +
        '-itc-checklist.txt';
      appendLog('Built ITC checklist for ' + data.partner + ' (config/' + data.config.id + '.json) ' +
        (data.config.created ? 'created' : 'reused') + '.');
      // Refresh the country dropdown so a newly created config is selectable.
      loadConfigs().then(() => { cfgSelect.value = data.config.id; });
      showToast('success', 'Checklist ready', 'Copy it, or download it as a .txt file, then fetch the Trade Map files.');
    } catch (e) {
      showError(String(e));
      appendLog('Build checklist failed: ' + e, 'err');
      showToast('error', 'Could not build checklist', String(e));
    } finally {
      prepBtn.disabled = false;
      prepBtn.innerHTML = '<i class="fas fa-list-check" aria-hidden="true"></i> Build checklist';
    }
  });

  function copyText(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      return navigator.clipboard.writeText(text);
    }
    return new Promise((resolve, reject) => {
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.style.position = 'fixed';
      ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.select();
      try { document.execCommand('copy'); resolve(); } catch (e) { reject(e); }
      ta.remove();
    });
  }

  checklistBox.addEventListener('click', e => {
    if (e.target.closest('#copy-checklist')) {
      if (!checklistText) return;
      copyText(checklistText).then(() =>
        showToast('success', 'Checklist copied', 'Paste it into a message, note or email.')
      ).catch(() => showToast('error', 'Copy failed', 'Select the text and copy it manually.'));
    } else if (e.target.closest('#download-checklist')) {
      if (!checklistText) return;
      const blob = new Blob([checklistText], { type: 'text/plain;charset=utf-8' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = checklistFileName;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
    }
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
          if (frac != null) setBuildProgress(frac);
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
    fd.append('config', rt === 'product' ? productCfgSelect.value : cfgSelect.value);
    fd.append('top', parseInt($('top').value, 10) || 10);
    fd.append('report_type', rt);
    fd.append('promotion', $('promotion').checked ? 'true' : 'false');
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
      setStage('Job accepted — building in the background…');
      setBuildProgress(0.05);

      // The server accepted the upload and builds the report in the
      // background; follow progress by polling the job status.
      appendLog(`Job accepted (${data.job_id}) — building in the background…`);
      let final;
      try {
        final = await pollStatus(data.status_url);
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
      setStage('Report ready!');
      setBuildProgress(1);
      const isProduct = final.mode === 'product' || final.mode === 'crafts';
      const isCompetitor = final.mode === 'competitor';
      const isExport = (final.mode || '').startsWith('quarterly');
      const modeLabel = isCompetitor ? 'Competitor analysis'
        : (isProduct ? 'Product profile'
        : (isExport ? 'Export performance (raw monthly extracts)'
        : (final.mode === 'raw' ? 'Built from raw ITC downloads'
        : ((final.mode || '').startsWith('services') ? 'Built from raw services downloads' : 'Ready-made tables'))));
      $('result-meta').innerHTML = '<i class="fas fa-check-circle" aria-hidden="true"></i> ' +
        modeLabel + ` &middot; job ${final.job_id}`;
      $('dl-report').href = final.report_url;
      $('dl-report').innerHTML = isCompetitor
        ? '<i class="fas fa-file-word" aria-hidden="true"></i> Download Table (.docx)'
        : '<i class="fas fa-file-word" aria-hidden="true"></i> Download Word Report (.docx)';
      $('dl-tables').href = isCompetitor ? (final.png_url || '#')
        : (isProduct ? (final.product_tables_url || '#') : final.tables_url || '#');
      $('dl-tables').innerHTML = isCompetitor
        ? '<i class="fas fa-image" aria-hidden="true"></i> Download Table Picture (.png)'
        : (isProduct
          ? '<i class="fas fa-file-excel" aria-hidden="true"></i> Download Excel Deliverable'
          : '<i class="fas fa-file-archive" aria-hidden="true"></i> Download Tables (.zip)');
      const dlPromo = $('dl-promotion');
      if (isCompetitor && final.xlsx_url) {
        dlPromo.href = final.xlsx_url;
        dlPromo.innerHTML = '<i class="fas fa-file-excel" aria-hidden="true"></i> Download Table (.xlsx)';
        dlPromo.style.display = 'inline-flex';
      } else if (final.promotion_url) {
        dlPromo.href = final.promotion_url;
        dlPromo.innerHTML = '<i class="fas fa-file-excel" aria-hidden="true"></i> Download Promotion Workbook (.xlsx)';
        dlPromo.style.display = 'inline-flex';
      } else {
        dlPromo.style.display = 'none';
      }
      $('result').style.display = 'block';
      $('result').classList.add('showing');
      showToast('success', 'Your report is ready!', 'The download links are waiting below — click here to jump to them.');
      setTimeout(() => {
        document.querySelector('.log-card').scrollIntoView({ behavior: 'smooth', block: 'start' });
      }, 400);
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

  // Initial panel state matches the default (goods) radio.
  const initialRadio = document.querySelector('input[name="report_type"]:checked');
  if (initialRadio) initialRadio.dispatchEvent(new Event('change', { bubbles: true }));
  else applyReportType('goods');
  loadConfigs();
  loadProductProfiles();
  loadSamples();
})();
