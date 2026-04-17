/**
 * Monkey Test Dashboard
 * HTTP 서버 + SSE 기반 실시간 모니터링 대시보드
 *
 * Usage:
 *   node src/dashboard.js [--port 3000]
 */
'use strict';

const http = require('http');
const { MonkeyTester } = require('./monkey-test');

// =============================================================================
// Config
// =============================================================================
const PORT = (() => {
    const idx = process.argv.indexOf('--port');
    return idx !== -1 ? parseInt(process.argv[idx + 1], 10) : 3000;
})();

// =============================================================================
// State
// =============================================================================
let activeTester = null;
let testPromise = null;
const sseClients = new Set();

// =============================================================================
// SSE broadcast
// =============================================================================
function broadcast(event, data) {
    const payload = `event: ${event}\ndata: ${JSON.stringify(data)}\n\n`;
    for (const res of sseClients) {
        try { res.write(payload); } catch (_) { sseClients.delete(res); }
    }
}

// =============================================================================
// Inline HTML dashboard
// =============================================================================
const DASHBOARD_HTML = `<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Monkey Test Dashboard — 한컴오피스</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans+KR:wght@300;400;600;700&display=swap');

  :root {
    --bg-base:    #0a0c10;
    --bg-panel:   #0f1117;
    --bg-elevated:#161b24;
    --bg-hover:   #1c2330;
    --border:     #1e2736;
    --border-lit: #2a3a52;
    --cyan:       #00d4ff;
    --cyan-dim:   #0099bb;
    --cyan-glow:  rgba(0,212,255,.18);
    --green:      #22c55e;
    --green-dim:  #166534;
    --amber:      #f59e0b;
    --amber-dim:  #78350f;
    --red:        #ef4444;
    --red-dim:    #7f1d1d;
    --purple:     #a855f7;
    --text-primary:  #e2e8f0;
    --text-secondary:#7c8fa6;
    --text-muted:    #3d5068;
    --mono: 'IBM Plex Mono', 'Courier New', monospace;
    --sans: 'IBM Plex Sans KR', 'Noto Sans KR', sans-serif;
    --radius: 6px;
    --transition: 180ms ease;
  }

  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  html, body {
    background: var(--bg-base);
    color: var(--text-primary);
    font-family: var(--sans);
    font-size: 14px;
    height: 100%;
    overflow: hidden;
  }

  /* ---- Layout ---- */
  .app {
    display: grid;
    grid-template-rows: auto auto 1fr;
    height: 100vh;
    gap: 0;
  }

  /* ---- Header ---- */
  .header {
    background: var(--bg-panel);
    border-bottom: 1px solid var(--border);
    padding: 14px 24px;
    display: flex;
    align-items: center;
    gap: 24px;
    flex-wrap: wrap;
    position: relative;
    z-index: 10;
  }

  .header::after {
    content: '';
    position: absolute;
    bottom: 0; left: 0; right: 0;
    height: 1px;
    background: linear-gradient(90deg, transparent, var(--cyan-dim), transparent);
    opacity: .4;
  }

  .brand {
    display: flex;
    align-items: center;
    gap: 10px;
    flex-shrink: 0;
  }

  .brand-icon {
    width: 28px; height: 28px;
    background: linear-gradient(135deg, var(--cyan), var(--cyan-dim));
    border-radius: 6px;
    display: flex; align-items: center; justify-content: center;
    font-size: 14px;
  }

  .brand-text {
    font-family: var(--mono);
    font-size: 15px;
    font-weight: 600;
    color: var(--text-primary);
    letter-spacing: -.3px;
  }

  .brand-sub {
    font-size: 10px;
    color: var(--text-muted);
    font-family: var(--mono);
    letter-spacing: .5px;
    text-transform: uppercase;
  }

  .divider-v {
    width: 1px;
    height: 32px;
    background: var(--border);
    flex-shrink: 0;
  }

  .controls {
    display: flex;
    align-items: center;
    gap: 12px;
    flex-wrap: wrap;
    flex: 1;
  }

  .field {
    display: flex;
    align-items: center;
    gap: 8px;
  }

  .field label {
    font-size: 11px;
    font-weight: 600;
    color: var(--text-muted);
    letter-spacing: .6px;
    text-transform: uppercase;
    white-space: nowrap;
  }

  .field select, .field input {
    background: var(--bg-elevated);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    color: var(--text-primary);
    font-family: var(--mono);
    font-size: 13px;
    padding: 6px 10px;
    outline: none;
    transition: border-color var(--transition), box-shadow var(--transition);
    -webkit-appearance: none;
  }

  .field select { cursor: pointer; }

  .field select:focus,
  .field input:focus {
    border-color: var(--cyan-dim);
    box-shadow: 0 0 0 3px var(--cyan-glow);
  }

  .field input[type="number"] { width: 72px; }
  .field input[type="text"]   { width: 110px; }

  .field select option { background: var(--bg-elevated); }

  /* ---- Buttons ---- */
  .btn {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 7px 16px;
    border-radius: var(--radius);
    font-size: 13px;
    font-weight: 600;
    font-family: var(--sans);
    cursor: pointer;
    border: 1px solid transparent;
    transition: all var(--transition);
    white-space: nowrap;
    letter-spacing: .2px;
  }

  .btn:disabled {
    opacity: .35;
    cursor: not-allowed;
    pointer-events: none;
  }

  .btn-start {
    background: var(--cyan);
    color: #000;
    border-color: var(--cyan);
  }
  .btn-start:hover {
    background: #33ddff;
    box-shadow: 0 0 16px var(--cyan-glow), 0 0 4px rgba(0,212,255,.5);
  }

  .btn-stop {
    background: transparent;
    color: var(--red);
    border-color: var(--red);
  }
  .btn-stop:hover {
    background: rgba(239,68,68,.1);
    box-shadow: 0 0 12px rgba(239,68,68,.2);
  }

  /* ---- Live indicator ---- */
  .live-badge {
    display: inline-flex;
    align-items: center;
    gap: 5px;
    font-family: var(--mono);
    font-size: 10px;
    font-weight: 600;
    letter-spacing: .8px;
    text-transform: uppercase;
    color: var(--text-muted);
    transition: color var(--transition);
  }
  .live-badge.active { color: var(--green); }

  .live-dot {
    width: 6px; height: 6px;
    border-radius: 50%;
    background: var(--text-muted);
    transition: background var(--transition), box-shadow var(--transition);
  }
  .live-badge.active .live-dot {
    background: var(--green);
    box-shadow: 0 0 6px var(--green);
    animation: pulse-dot 1.4s ease-in-out infinite;
  }

  @keyframes pulse-dot {
    0%, 100% { opacity: 1; box-shadow: 0 0 4px var(--green); }
    50%       { opacity: .5; box-shadow: 0 0 10px var(--green); }
  }

  /* ---- Stats bar ---- */
  .stats-bar {
    background: var(--bg-panel);
    border-bottom: 1px solid var(--border);
    padding: 0 24px;
    display: flex;
    align-items: stretch;
    gap: 0;
  }

  .stat-cell {
    display: flex;
    flex-direction: column;
    justify-content: center;
    padding: 14px 28px 14px 0;
    margin-right: 28px;
    border-right: 1px solid var(--border);
    min-width: 100px;
  }
  .stat-cell:last-child {
    border-right: none;
    margin-right: 0;
    flex: 1;
  }

  .stat-label {
    font-size: 10px;
    font-weight: 700;
    letter-spacing: .8px;
    text-transform: uppercase;
    color: var(--text-muted);
    margin-bottom: 4px;
  }

  .stat-value {
    font-family: var(--mono);
    font-size: 22px;
    font-weight: 600;
    color: var(--text-primary);
    line-height: 1;
    transition: color .15s;
  }

  .stat-value.cyan  { color: var(--cyan); }
  .stat-value.green { color: var(--green); }
  .stat-value.amber { color: var(--amber); }
  .stat-value.red   { color: var(--red); }

  .stat-flash {
    animation: flash-up .3s ease-out;
  }

  @keyframes flash-up {
    0%   { transform: translateY(-4px); opacity: .5; }
    100% { transform: translateY(0);    opacity: 1; }
  }

  /* ---- Progress bar ---- */
  .progress-row {
    padding: 0 24px;
    background: var(--bg-panel);
    border-bottom: 1px solid var(--border);
  }

  .progress-track {
    height: 3px;
    background: var(--border);
    border-radius: 2px;
    overflow: hidden;
  }

  .progress-fill {
    height: 100%;
    background: linear-gradient(90deg, var(--cyan-dim), var(--cyan));
    border-radius: 2px;
    width: 0%;
    transition: width .8s linear;
    box-shadow: 0 0 8px rgba(0,212,255,.4);
  }

  .progress-meta {
    display: flex;
    justify-content: space-between;
    padding: 6px 0 8px;
    font-family: var(--mono);
    font-size: 11px;
    color: var(--text-muted);
  }

  /* ---- Main area ---- */
  .main {
    display: grid;
    grid-template-columns: 1fr;
    overflow: hidden;
    background: var(--bg-base);
  }

  /* ---- Log panel ---- */
  .log-panel {
    display: flex;
    flex-direction: column;
    overflow: hidden;
    padding: 0;
  }

  .log-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 10px 20px;
    border-bottom: 1px solid var(--border);
    background: var(--bg-panel);
    flex-shrink: 0;
  }

  .log-title {
    font-size: 11px;
    font-weight: 700;
    letter-spacing: .8px;
    text-transform: uppercase;
    color: var(--text-muted);
    display: flex;
    align-items: center;
    gap: 8px;
  }

  .log-count {
    font-family: var(--mono);
    font-size: 10px;
    background: var(--bg-elevated);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 1px 8px;
    color: var(--text-secondary);
  }

  .btn-clear {
    font-size: 11px;
    padding: 3px 10px;
    background: transparent;
    border: 1px solid var(--border);
    color: var(--text-muted);
    border-radius: 4px;
    cursor: pointer;
    transition: all var(--transition);
    font-family: var(--sans);
  }
  .btn-clear:hover {
    border-color: var(--border-lit);
    color: var(--text-secondary);
  }

  .log-scroll {
    flex: 1;
    overflow-y: auto;
    scroll-behavior: smooth;
    font-family: var(--mono);
    font-size: 12px;
    line-height: 1;
  }

  .log-scroll::-webkit-scrollbar { width: 6px; }
  .log-scroll::-webkit-scrollbar-track { background: transparent; }
  .log-scroll::-webkit-scrollbar-thumb { background: var(--border-lit); border-radius: 3px; }
  .log-scroll::-webkit-scrollbar-thumb:hover { background: var(--text-muted); }

  .log-entry {
    display: grid;
    grid-template-columns: 56px 52px 90px 1fr;
    align-items: center;
    gap: 0;
    padding: 0 8px;
    min-height: 28px;
    border-bottom: 1px solid transparent;
    animation: entry-in .2s ease-out;
    transition: background var(--transition);
  }

  .log-entry:hover { background: var(--bg-hover); }

  .log-entry.new-entry {
    animation: entry-in .25s ease-out, entry-highlight .6s ease-out;
  }

  @keyframes entry-in {
    from { opacity: 0; transform: translateX(-6px); }
    to   { opacity: 1; transform: translateX(0); }
  }

  @keyframes entry-highlight {
    0%   { background: rgba(0,212,255,.08); }
    100% { background: transparent; }
  }

  .log-seq {
    font-size: 10px;
    color: var(--text-muted);
    text-align: right;
    padding-right: 8px;
    letter-spacing: -.5px;
  }

  .log-time {
    font-size: 10px;
    color: var(--text-muted);
    padding-right: 8px;
  }

  .log-badge {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 60px;
    height: 18px;
    border-radius: 3px;
    font-size: 9px;
    font-weight: 700;
    letter-spacing: .5px;
    text-transform: uppercase;
    flex-shrink: 0;
    margin-right: 10px;
  }

  .badge-session  { background: rgba(168,85,247,.2);  color: #c084fc; border: 1px solid rgba(168,85,247,.3); }
  .badge-click_menu   { background: rgba(0,212,255,.12);  color: var(--cyan); border: 1px solid rgba(0,212,255,.25); }
  .badge-click_button { background: rgba(0,212,255,.12);  color: var(--cyan); border: 1px solid rgba(0,212,255,.25); }
  .badge-type_text    { background: rgba(34,197,94,.12);  color: var(--green); border: 1px solid rgba(34,197,94,.25); }
  .badge-press_keys   { background: rgba(34,197,94,.12);  color: var(--green); border: 1px solid rgba(34,197,94,.25); }
  .badge-explore_ui   { background: rgba(100,116,139,.15); color: #94a3b8; border: 1px solid rgba(100,116,139,.3); }
  .badge-handle_dialog{ background: rgba(245,158,11,.12); color: var(--amber); border: 1px solid rgba(245,158,11,.25); }
  .badge-warning  { background: rgba(245,158,11,.15); color: var(--amber); border: 1px solid rgba(245,158,11,.3); }
  .badge-crash    { background: rgba(239,68,68,.18);  color: var(--red); border: 1px solid rgba(239,68,68,.4); }
  .badge-error    { background: rgba(239,68,68,.12);  color: var(--red); border: 1px solid rgba(239,68,68,.25); }
  .badge-unknown  { background: rgba(100,116,139,.15); color: #94a3b8; border: 1px solid rgba(100,116,139,.3); }

  .log-detail {
    color: var(--text-secondary);
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    font-size: 12px;
    padding-left: 2px;
  }

  .log-detail .action-name {
    color: var(--text-primary);
    font-weight: 600;
    margin-right: 6px;
  }

  .log-detail .detail-kv {
    color: var(--text-muted);
    font-size: 11px;
  }

  .log-detail .detail-kv .key { color: var(--cyan-dim); }
  .log-detail .detail-kv .val { color: var(--text-secondary); }

  .log-entry.type-crash   .log-detail { color: var(--red); }
  .log-entry.type-error   .log-detail { color: #fb7185; }
  .log-entry.type-warning .log-detail { color: var(--amber); }

  /* ---- Empty state ---- */
  .log-empty {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    height: 100%;
    gap: 12px;
    color: var(--text-muted);
    user-select: none;
  }

  .log-empty-icon {
    font-size: 40px;
    opacity: .3;
  }

  .log-empty-text {
    font-size: 13px;
    letter-spacing: .3px;
  }

  /* ---- Toast ---- */
  .toast-container {
    position: fixed;
    bottom: 24px;
    right: 24px;
    display: flex;
    flex-direction: column;
    gap: 8px;
    z-index: 100;
  }

  .toast {
    background: var(--bg-elevated);
    border: 1px solid var(--border-lit);
    border-radius: var(--radius);
    padding: 10px 16px;
    font-size: 13px;
    color: var(--text-primary);
    animation: toast-in .25s ease-out;
    display: flex;
    align-items: center;
    gap: 8px;
    max-width: 320px;
    box-shadow: 0 4px 20px rgba(0,0,0,.5);
  }

  .toast.error { border-color: var(--red); color: #fca5a5; }
  .toast.success { border-color: var(--green); }

  @keyframes toast-in {
    from { opacity: 0; transform: translateX(20px); }
    to   { opacity: 1; transform: translateX(0); }
  }

  @keyframes toast-out {
    from { opacity: 1; transform: translateX(0); }
    to   { opacity: 0; transform: translateX(20px); }
  }

  /* ---- Responsive ---- */
  @media (max-width: 640px) {
    .header { padding: 10px 12px; gap: 10px; }
    .stats-bar { gap: 0; overflow-x: auto; }
    .stat-cell { min-width: 80px; padding: 10px 16px 10px 0; }
    .stat-value { font-size: 18px; }
    .log-entry { grid-template-columns: 40px 48px 70px 1fr; font-size: 11px; }
  }
</style>
</head>
<body>
<div class="app">

  <!-- ========== HEADER ========== -->
  <header class="header">
    <div class="brand">
      <div class="brand-icon">🐒</div>
      <div>
        <div class="brand-text">Monkey Test Dashboard</div>
        <div class="brand-sub">한컴오피스 자동화</div>
      </div>
    </div>

    <div class="divider-v"></div>

    <div class="controls">
      <div class="field">
        <label>제품</label>
        <select id="sel-product">
          <option value="hwp">한글 (HWP)</option>
          <option value="hword">한워드</option>
          <option value="hshow">한쇼</option>
          <option value="hcell">한셀</option>
        </select>
      </div>

      <div class="field">
        <label>시간(분)</label>
        <input type="number" id="inp-duration" value="5" min="1" max="180">
      </div>

      <div class="field">
        <label>시드</label>
        <input type="text" id="inp-seed" placeholder="auto">
      </div>

      <button class="btn btn-start" id="btn-start" onclick="startTest()">
        <span>▶</span> 시작
      </button>
      <button class="btn btn-stop" id="btn-stop" onclick="stopTest()" disabled>
        <span>■</span> 중단
      </button>
    </div>

    <div class="live-badge" id="live-badge">
      <div class="live-dot"></div>
      <span id="live-label">대기중</span>
    </div>
  </header>

  <!-- ========== STATS BAR ========== -->
  <div class="stats-bar">
    <div class="stat-cell">
      <div class="stat-label">경과 시간</div>
      <div class="stat-value cyan" id="stat-elapsed">00:00</div>
    </div>
    <div class="stat-cell">
      <div class="stat-label">전체 액션</div>
      <div class="stat-value" id="stat-total">0</div>
    </div>
    <div class="stat-cell">
      <div class="stat-label">성공</div>
      <div class="stat-value green" id="stat-success">0</div>
    </div>
    <div class="stat-cell">
      <div class="stat-label">실패</div>
      <div class="stat-value amber" id="stat-failed">0</div>
    </div>
    <div class="stat-cell" style="border-right:none;">
      <div class="stat-label">크래시</div>
      <div class="stat-value red" id="stat-crashes">0</div>
    </div>
  </div>

  <!-- ========== PROGRESS ========== -->
  <div class="progress-row">
    <div class="progress-meta">
      <span id="prog-label">시작 전</span>
      <span id="prog-pct">0%</span>
    </div>
    <div class="progress-track">
      <div class="progress-fill" id="progress-fill"></div>
    </div>
    <div style="height:6px"></div>
  </div>

  <!-- ========== LOG ========== -->
  <div class="main">
    <div class="log-panel">
      <div class="log-header">
        <div class="log-title">
          액션 로그
          <span class="log-count" id="log-count">0</span>
        </div>
        <button class="btn-clear" onclick="clearLog()">비우기</button>
      </div>
      <div class="log-scroll" id="log-scroll">
        <div class="log-empty" id="log-empty">
          <div class="log-empty-icon">🐒</div>
          <div class="log-empty-text">테스트를 시작하면 여기에 액션 로그가 표시됩니다</div>
        </div>
      </div>
    </div>
  </div>
</div>

<div class="toast-container" id="toast-container"></div>

<script>
// =====================================================================
// State
// =====================================================================
let evtSource = null;
let running = false;
let startTime = null;
let durationMs = 0;
let elapsedTimer = null;
let logEntryCount = 0;
let autoScroll = true;

// =====================================================================
// SSE connection
// =====================================================================
function connectSSE() {
  if (evtSource) { evtSource.close(); }
  evtSource = new EventSource('/events');

  evtSource.addEventListener('action', e => {
    const data = JSON.parse(e.data);
    appendLog(data);
    updateStats(data);
  });

  evtSource.addEventListener('status', e => {
    const data = JSON.parse(e.data);
    syncStatusUI(data);
  });

  evtSource.addEventListener('started', e => {
    const data = JSON.parse(e.data);
    onTestStarted(data);
  });

  evtSource.addEventListener('stopped', e => {
    onTestStopped();
  });

  evtSource.addEventListener('error', () => {
    // SSE connection error — will auto-reconnect
  });
}

// =====================================================================
// UI helpers
// =====================================================================
function setRunning(isRunning) {
  running = isRunning;
  document.getElementById('btn-start').disabled = isRunning;
  document.getElementById('btn-stop').disabled = !isRunning;
  document.getElementById('sel-product').disabled = isRunning;
  document.getElementById('inp-duration').disabled = isRunning;
  document.getElementById('inp-seed').disabled = isRunning;

  const badge = document.getElementById('live-badge');
  const label = document.getElementById('live-label');
  if (isRunning) {
    badge.classList.add('active');
    label.textContent = '실행중';
  } else {
    badge.classList.remove('active');
    label.textContent = '대기중';
  }
}

function flashStat(id) {
  const el = document.getElementById(id);
  el.classList.remove('stat-flash');
  void el.offsetWidth; // reflow
  el.classList.add('stat-flash');
  setTimeout(() => el.classList.remove('stat-flash'), 400);
}

function setStatValue(id, value) {
  const el = document.getElementById(id);
  if (el.textContent !== String(value)) {
    el.textContent = value;
    flashStat(id);
  }
}

function formatElapsed(ms) {
  const s = Math.floor(ms / 1000);
  const m = Math.floor(s / 60);
  const h = Math.floor(m / 60);
  if (h > 0) return String(h).padStart(2,'0') + ':' + String(m%60).padStart(2,'0') + ':' + String(s%60).padStart(2,'0');
  return String(m).padStart(2,'0') + ':' + String(s%60).padStart(2,'0');
}

function updateElapsedTick() {
  if (!running || !startTime) return;
  const elapsed = Date.now() - startTime;
  document.getElementById('stat-elapsed').textContent = formatElapsed(elapsed);
  updateProgress(elapsed);
}

function updateProgress(elapsed) {
  if (!durationMs) return;
  const pct = Math.min(100, (elapsed / durationMs) * 100);
  document.getElementById('progress-fill').style.width = pct.toFixed(1) + '%';
  document.getElementById('prog-pct').textContent = pct.toFixed(0) + '%';

  const remaining = Math.max(0, durationMs - elapsed);
  const remS = Math.floor(remaining / 1000);
  const remM = Math.floor(remS / 60);
  document.getElementById('prog-label').textContent =
    '남은 시간: ' + String(remM).padStart(2,'0') + ':' + String(remS%60).padStart(2,'0');
}

// =====================================================================
// Log rendering
// =====================================================================
const BADGE_MAP = {
  click_menu: 'MENU', click_button: 'BTN',
  type_text: 'TYPE', press_keys: 'KEY',
  explore_ui: 'SCAN', handle_dialog: 'DLG',
  session: 'SES', crash: 'CRASH',
  warning: 'WARN', error: 'ERR',
};

function getBadgeClass(type) {
  const known = ['click_menu','click_button','type_text','press_keys',
                 'explore_ui','handle_dialog','session','crash','warning','error'];
  return known.includes(type) ? 'badge-' + type : 'badge-unknown';
}

function buildDetail(entry) {
  const parts = [];
  if (entry.name)          parts.push('<span class="detail-kv"><span class="key">name</span>=<span class="val">"' + esc(entry.name) + '"</span></span>');
  if (entry.keys)          parts.push('<span class="detail-kv"><span class="key">keys</span>=<span class="val">' + esc(entry.keys) + '</span></span>');
  if (entry.text)          parts.push('<span class="detail-kv"><span class="key">text</span>=<span class="val">"' + esc(entry.text.substring(0,40)) + '"</span></span>');
  if (entry.error)         parts.push('<span class="detail-kv" style="color:#fb7185"><span class="key">err</span>=<span class="val">' + esc(entry.error.substring(0,60)) + '</span></span>');
  if (entry.totalElements) parts.push('<span class="detail-kv"><span class="key">elements</span>=<span class="val">' + entry.totalElements + '</span></span>');
  if (entry.crashCount)    parts.push('<span class="detail-kv" style="color:var(--red)"><span class="key">crash#</span><span class="val">' + entry.crashCount + '</span></span>');
  if (entry.product)       parts.push('<span class="detail-kv"><span class="key">product</span>=<span class="val">' + esc(entry.product) + '</span></span>');
  return parts.join(' ');
}

function esc(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function appendLog(entry) {
  const empty = document.getElementById('log-empty');
  if (empty) empty.remove();

  logEntryCount++;
  document.getElementById('log-count').textContent = logEntryCount;

  const time = entry.time ? entry.time.substring(11,19) : '--:--:--';
  const badgeLabel = BADGE_MAP[entry.type] || entry.type.substring(0,5).toUpperCase();
  const badgeClass = getBadgeClass(entry.type);
  const typeClass = 'type-' + entry.type;

  const div = document.createElement('div');
  div.className = 'log-entry new-entry ' + typeClass;
  div.innerHTML =
    '<span class="log-seq">#' + (entry.seq || logEntryCount) + '</span>' +
    '<span class="log-time">' + time + '</span>' +
    '<span class="log-badge ' + badgeClass + '">' + badgeLabel + '</span>' +
    '<span class="log-detail"><span class="action-name">' + esc(entry.action || '') + '</span>' + buildDetail(entry) + '</span>';

  const scroll = document.getElementById('log-scroll');
  scroll.appendChild(div);

  setTimeout(() => div.classList.remove('new-entry'), 700);

  if (autoScroll) {
    scroll.scrollTop = scroll.scrollHeight;
  }
}

// =====================================================================
// Stats update from action events
// =====================================================================
let statsCache = { total: 0, success: 0, failed: 0, crashes: 0 };

function updateStats(entry) {
  if (entry.type !== 'session') statsCache.total++;
  if (entry.type === 'crash') statsCache.crashes++;
  if (entry.action === 'failed' || entry.type === 'error') statsCache.failed++;
  // rough success heuristic: non-error, non-crash actions that ended cleanly
  // actual counts come from server status polls

  setStatValue('stat-total', statsCache.total);
  setStatValue('stat-crashes', statsCache.crashes);
}

function syncStatusUI(status) {
  if (status.totalActions !== undefined)  setStatValue('stat-total',   status.totalActions);
  if (status.crashes !== undefined)       setStatValue('stat-crashes',  status.crashes);
  if (status.failedActions !== undefined) setStatValue('stat-failed',   status.failedActions);
  const success = (status.totalActions || 0) - (status.failedActions || 0) - (status.crashes || 0);
  if (success >= 0) setStatValue('stat-success', success);
}

// =====================================================================
// Test lifecycle
// =====================================================================
function onTestStarted(data) {
  startTime = Date.now();
  durationMs = data.durationMs || 0;
  statsCache = { total: 0, success: 0, failed: 0, crashes: 0 };
  setRunning(true);

  if (elapsedTimer) clearInterval(elapsedTimer);
  elapsedTimer = setInterval(updateElapsedTick, 500);

  document.getElementById('prog-label').textContent = '실행중...';
  document.getElementById('progress-fill').style.width = '0%';
  document.getElementById('prog-pct').textContent = '0%';
}

function onTestStopped() {
  setRunning(false);
  if (elapsedTimer) { clearInterval(elapsedTimer); elapsedTimer = null; }
  document.getElementById('live-badge').classList.remove('active');
  document.getElementById('live-label').textContent = '완료';
  document.getElementById('prog-label').textContent = '테스트 종료';
  showToast('테스트가 종료되었습니다.', 'success');
}

// =====================================================================
// API calls
// =====================================================================
async function startTest() {
  const product  = document.getElementById('sel-product').value;
  const duration = parseInt(document.getElementById('inp-duration').value, 10) || 5;
  const seedRaw  = document.getElementById('inp-seed').value.trim();
  const seed     = seedRaw ? parseInt(seedRaw, 10) : undefined;

  try {
    const res = await fetch('/api/start', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ product, durationMs: duration * 60 * 1000, seed }),
    });
    const data = await res.json();
    if (!res.ok) { showToast(data.error || '시작 실패', 'error'); return; }
    showToast('테스트 시작: ' + data.product + ' (시드 ' + data.seed + ')', 'success');
  } catch (e) {
    showToast('서버 연결 실패: ' + e.message, 'error');
  }
}

async function stopTest() {
  try {
    const res = await fetch('/api/stop', { method: 'POST' });
    const data = await res.json();
    if (!res.ok) { showToast(data.error || '중단 실패', 'error'); return; }
    showToast('중단 요청 전송됨', 'success');
  } catch (e) {
    showToast('서버 연결 실패: ' + e.message, 'error');
  }
}

// =====================================================================
// Toast notifications
// =====================================================================
function showToast(msg, type) {
  const container = document.getElementById('toast-container');
  const toast = document.createElement('div');
  toast.className = 'toast' + (type ? ' ' + type : '');
  const icon = type === 'error' ? '✗' : type === 'success' ? '✓' : 'ℹ';
  toast.innerHTML = '<span>' + icon + '</span><span>' + esc(msg) + '</span>';
  container.appendChild(toast);
  setTimeout(() => {
    toast.style.animation = 'toast-out .25s ease-in forwards';
    setTimeout(() => toast.remove(), 260);
  }, 3500);
}

// =====================================================================
// Auto-scroll toggle on manual scroll
// =====================================================================
document.getElementById('log-scroll').addEventListener('scroll', function() {
  const el = this;
  autoScroll = (el.scrollHeight - el.scrollTop - el.clientHeight) < 40;
});

// =====================================================================
// Clear log
// =====================================================================
function clearLog() {
  const scroll = document.getElementById('log-scroll');
  scroll.innerHTML = '<div class="log-empty" id="log-empty"><div class="log-empty-icon">🐒</div><div class="log-empty-text">테스트를 시작하면 여기에 액션 로그가 표시됩니다</div></div>';
  logEntryCount = 0;
  document.getElementById('log-count').textContent = '0';
}

// =====================================================================
// Init: fetch current status and sync
// =====================================================================
async function init() {
  connectSSE();

  try {
    const res = await fetch('/api/status');
    const status = await res.json();
    if (status.running) {
      startTime = Date.now() - (status.elapsed || 0);
      durationMs = (status.elapsed || 0) + (status.remaining || 0);
      statsCache.total   = status.totalActions || 0;
      statsCache.crashes = status.crashes || 0;
      statsCache.failed  = status.failedActions || 0;
      syncStatusUI(status);
      setRunning(true);
      if (elapsedTimer) clearInterval(elapsedTimer);
      elapsedTimer = setInterval(updateElapsedTick, 500);
      document.getElementById('sel-product').value = status.product || 'hwp';
    }
  } catch (_) { /* server not ready */ }
}

init();
</script>
</body>
</html>`;

// =============================================================================
// Request router
// =============================================================================
function sendJSON(res, code, obj) {
    const body = JSON.stringify(obj);
    res.writeHead(code, {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*',
    });
    res.end(body);
}

function readBody(req) {
    return new Promise((resolve, reject) => {
        let raw = '';
        req.on('data', chunk => { raw += chunk; });
        req.on('end', () => {
            try { resolve(JSON.parse(raw || '{}')); }
            catch (_) { resolve({}); }
        });
        req.on('error', reject);
    });
}

async function handleRequest(req, res) {
    const { method, url } = req;

    // ---- Dashboard HTML ----
    if (method === 'GET' && url === '/') {
        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end(DASHBOARD_HTML);
        return;
    }

    // ---- SSE endpoint ----
    if (method === 'GET' && url === '/events') {
        res.writeHead(200, {
            'Content-Type':  'text/event-stream',
            'Cache-Control': 'no-cache',
            'Connection':    'keep-alive',
            'Access-Control-Allow-Origin': '*',
        });
        res.write(':ok\n\n');

        // Keep-alive ping every 20s
        const ping = setInterval(() => {
            try { res.write(':ping\n\n'); } catch (_) { /* ignore */ }
        }, 20000);

        sseClients.add(res);

        req.on('close', () => {
            clearInterval(ping);
            sseClients.delete(res);
        });

        // Send current status immediately on connect
        if (activeTester) {
            const status = activeTester.getStatus();
            res.write(`event: status\ndata: ${JSON.stringify(status)}\n\n`);
        }
        return;
    }

    // ---- CORS preflight ----
    if (method === 'OPTIONS') {
        res.writeHead(204, {
            'Access-Control-Allow-Origin':  '*',
            'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type',
        });
        res.end();
        return;
    }

    // ---- POST /api/start ----
    if (method === 'POST' && url === '/api/start') {
        if (activeTester && activeTester.running) {
            sendJSON(res, 409, { error: '이미 실행 중입니다. 먼저 중단해 주세요.' });
            return;
        }

        const body = await readBody(req);
        const product    = ['hwp','hword','hshow','hcell'].includes(body.product) ? body.product : 'hwp';
        const durationMs = typeof body.durationMs === 'number' && body.durationMs > 0
            ? body.durationMs : 5 * 60 * 1000;
        const seed = typeof body.seed === 'number' ? body.seed : Date.now();

        activeTester = new MonkeyTester({ product, durationMs, seed });

        // Patch _log to broadcast SSE events
        const originalLog = activeTester._log.bind(activeTester);
        activeTester._log = function(type, action, data) {
            originalLog(type, action, data);
            // Broadcast the new log entry to all SSE clients
            const entry = activeTester.stats.actions[activeTester.stats.actions.length - 1];
            if (entry) broadcast('action', entry);
            // Broadcast summary status every action
            broadcast('status', activeTester.getStatus());
        };

        broadcast('started', { product, durationMs, seed });

        testPromise = activeTester.start()
            .then(report => {
                try { activeTester.exportReport(); } catch (_) {}
                broadcast('stopped', { report: report.summary });
                return report;
            })
            .catch(err => {
                broadcast('stopped', { error: err.message });
            });

        sendJSON(res, 200, { started: true, product, durationMs, seed });
        return;
    }

    // ---- POST /api/stop ----
    if (method === 'POST' && url === '/api/stop') {
        if (!activeTester || !activeTester.running) {
            sendJSON(res, 409, { error: '실행 중인 테스트가 없습니다.' });
            return;
        }
        activeTester.stop();
        sendJSON(res, 200, { stopped: true });
        return;
    }

    // ---- GET /api/status ----
    if (method === 'GET' && url === '/api/status') {
        if (!activeTester) {
            sendJSON(res, 200, { running: false });
            return;
        }
        sendJSON(res, 200, activeTester.getStatus());
        return;
    }

    // ---- 404 ----
    res.writeHead(404, { 'Content-Type': 'text/plain' });
    res.end('Not Found');
}

// =============================================================================
// Server bootstrap
// =============================================================================
const server = http.createServer((req, res) => {
    handleRequest(req, res).catch(err => {
        console.error('[dashboard] unhandled error:', err);
        try {
            res.writeHead(500, { 'Content-Type': 'text/plain' });
            res.end('Internal Server Error');
        } catch (_) {}
    });
});

server.listen(PORT, () => {
    console.log(`[dashboard] 서버 시작: http://localhost:${PORT}`);
    console.log(`[dashboard] 대시보드: http://localhost:${PORT}/`);
});

// =============================================================================
// Graceful shutdown
// =============================================================================
function shutdown(signal) {
    console.log(`\n[dashboard] ${signal} 수신 — 종료 중...`);

    if (activeTester && activeTester.running) {
        console.log('[dashboard] 진행 중인 테스트를 중단합니다...');
        activeTester.stop();
    }

    server.close(() => {
        console.log('[dashboard] 서버 종료 완료.');
        process.exit(0);
    });

    // Force exit after 5s
    setTimeout(() => process.exit(1), 5000).unref();
}

process.on('SIGINT',  () => shutdown('SIGINT'));
process.on('SIGTERM', () => shutdown('SIGTERM'));
process.on('uncaughtException', err => {
    console.error('[dashboard] uncaughtException:', err);
});
process.on('unhandledRejection', (reason) => {
    console.error('[dashboard] unhandledRejection:', reason);
});
