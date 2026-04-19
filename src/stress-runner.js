/**
 * Stress Runner — 단일 문서에서 장시간 랜덤 테스트를 수행한다.
 *
 * 목표: 기능 정합성 검증이 아니라 "한글이 안정적으로 돌아가는가"를 본다.
 *   - 매 iteration마다 맵 기반 랜덤 리본 항목 또는 시나리오 실행
 *   - 프로세스 crash / 응답 없음(hang) 감지 → replay 로그 저장 + 재시작
 *   - 모든 action을 ndjson에 기록 → 재현 스텝 자료
 *
 * 가중치 (iteration당):
 *   55% context-aware 리본 항목 (맵 샘플링)
 *   15% 텍스트 입력 (문서 확장)
 *   10% 사전 정의 시나리오 (region-crossing copy/paste 등)
 *   10% undo/redo (Ctrl+Z / Ctrl+Y)
 *   5%  네비게이션 (PageUp/PageDown/Ctrl+Home/Ctrl+End)
 *   5%  우클릭 컨텍스트 메뉴에서 랜덤 항목
 */
'use strict';

const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');

const controller = require('./hwp-controller');
const sessionModule = require('./session');
const win32 = require('./win32');
const { CancelledError, CancelWatcher } = require('./cancel-watcher');
const { StressContext, STATES } = require('./stress-context');
const { loadMap, mapToCases, weightedForState, pickWeighted } = require('./map-to-cases');
const seeder = require('./stress-seeder');
const { SCENARIOS, runScenario } = require('./stress-scenarios');
const hwpCom = require('./hwp-com');

const { execSync } = require('child_process');

const log = (icon, msg) => process.stderr.write(`${icon} [stress] ${msg}\n`);
const sleep = (ms) => new Promise(r => setTimeout(r, ms));

/** 현재 실행 중인 Hwp.exe PID 목록 */
function listHwpPids() {
    try {
        const out = execSync('tasklist /FI "IMAGENAME eq Hwp.exe" /FO CSV /NH',
            { encoding: 'utf8', windowsHide: true });
        const pids = [];
        for (const line of out.split('\n')) {
            const m = /^"Hwp\.exe","(\d+)"/.exec(line);
            if (m) pids.push(parseInt(m[1], 10));
        }
        return pids;
    } catch (_) { return []; }
}

// =============================================================================
// PRNG (mulberry32) — 시드 기반 재현 가능 PRNG
// =============================================================================

function mulberry32(seed) {
    let s = seed >>> 0;
    return function () {
        s = (s + 0x6D2B79F5) >>> 0;
        let t = Math.imul(s ^ (s >>> 15), 1 | s);
        t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
        return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
}

// =============================================================================
// 가중치 테이블
// =============================================================================

const ACTION_MIX = [
    { type: 'ribbon',     weight: 55 },
    { type: 'typeText',   weight: 15 },
    { type: 'scenario',   weight: 10 },
    { type: 'undoRedo',   weight: 10 },
    { type: 'navigate',   weight: 5  },
    { type: 'contextMenu',weight: 5  },
];

const TEXT_SAMPLES = [
    '이것은 stress 테스트용 입력 문장입니다. ',
    'The quick brown fox jumps over the lazy dog. ',
    '가나다라마바사 1234567890 !@#$%^&*() ',
    '한컴오피스 2027 안정성 검증 — ',
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. ',
];

const NAV_KEYS = ['Ctrl+Home', 'Ctrl+End', 'PageDown', 'PageUp', 'Home', 'End'];
const UNDO_REDO_KEYS = ['Ctrl+Z', 'Ctrl+Z', 'Ctrl+Z', 'Ctrl+Y'];

// =============================================================================
// 가중 선택 유틸
// =============================================================================

function pickFromMix(mix, prng) {
    const total = mix.reduce((s, m) => s + m.weight, 0);
    let r = prng() * total;
    for (const m of mix) {
        r -= m.weight;
        if (r <= 0) return m.type;
    }
    return mix[mix.length - 1].type;
}

function pickOne(arr, prng) {
    return arr[Math.floor(prng() * arr.length)];
}

// =============================================================================
// 상태 전이 힌트 — 리본 항목 실행 후 state를 업데이트
// =============================================================================

function inferStateAfter(caseObj, prevState) {
    const n = caseObj.name || '';
    if (/표\s*만들기|표\s*삽입/.test(n))   return STATES.IN_TABLE_CELL;
    if (/도형|그리기|글상자/.test(n))      return STATES.IN_SHAPE_TEXT;
    if (/각주/.test(n))                    return STATES.IN_FOOTNOTE;
    if (/미주/.test(n))                    return STATES.IN_ENDNOTE;
    if (/메모/.test(n))                    return STATES.IN_MEMO;
    if (/머리말/.test(n) && caseObj.tab === '쪽') return STATES.IN_HEADER;
    if (/꼬리말/.test(n) && caseObj.tab === '쪽') return STATES.IN_FOOTER;
    return prevState; // 변화 없음
}

// =============================================================================
// 건강 검진 — crash / hang 감지
// =============================================================================

/**
 * HWP COM으로 실시간 caret/selection 상태를 읽어 ctx.state 동기화.
 * winax 미사용/HWP COM 실패 시엔 no-op이라 기존 추정 방식이 유지됨.
 * 오버헤드: COM 호출당 ~10~30ms → 매 iter가 아닌 몇 iter 간격으로 호출.
 */
function syncStateFromCOM(ctx) {
    if (!hwpCom.isAvailable()) return null;
    const snap = hwpCom.getStateSnapshot();
    if (!snap) return null;
    const mapped = hwpCom.ctrlIdToContextState(snap);
    // COM이 명확히 특정 state를 감지한 경우에만 덮어쓴다.
    // null이면 "모름"이므로 inferStateAfter/시나리오가 설정한 추정 state를 보존.
    // 이 정책으로 COM(확실) > 추정(덜 확실) > 기본(body) 우선순위가 성립.
    if (mapped && ctx.state !== mapped) {
        ctx.transition(mapped, `com-sync(parent=${snap.parentCtrlId || '-'}, sel=${snap.selectedCtrlId || '-'})`);
    }
    return snap;
}

/**
 * 매 iter 시작 전 호출되는 환경 복구 guard.
 *   1) HWP 포그라운드 — 다른 앱이 포커스 훔쳤으면 되찾기
 *   2) 창 최대화 — 일부 액션 후 최대화가 풀리면 맵 좌표 오클릭 발생
 *   3) 잔여 모달/드롭다운 정리 — Esc 2~3회로 이전 iter 잔재 제거
 */
async function restoreEnvironment() {
    try {
        const session = sessionModule.getSession();
        const hwpPid = session && session.hwpProcess && session.hwpProcess.pid;
        const hwnd = session && session.hwpProcess && session.hwpProcess.hwnd;

        // 포커스 확인·복구
        const fgPid = win32.getForegroundPid();
        if (fgPid !== hwpPid) {
            try { await controller.setForeground(); } catch (_) {}
        }

        // 최대화 재적용 (상태와 무관하게 재설정 — 이미 최대화면 no-op)
        if (hwnd) win32.maximizeWindow(hwnd);

        // 잔여 다이얼로그/팝업 닫기
        try { await controller.pressKeys({ keys: 'Escape' }); } catch (_) {}
        try { await controller.pressKeys({ keys: 'Escape' }); } catch (_) {}
    } catch (_) {}
}

function checkAlive() {
    try {
        const session = sessionModule.getSession();
        if (!session || !session.hwpProcess) return { alive: false, reason: 'no-session' };
        const pid = session.hwpProcess.pid;
        if (!pid) return { alive: false, reason: 'no-pid' };
        if (!win32.isProcessAlive(pid)) return { alive: false, reason: 'process-dead' };
        const hwnd = session.hwpProcess.hwnd;
        if (hwnd && win32.isHungAppWindow(hwnd)) return { alive: false, reason: 'hung' };
        return { alive: true };
    } catch (e) {
        return { alive: false, reason: `check-error: ${e.message}` };
    }
}

// =============================================================================
// Replay 버퍼 — 최근 N개 액션을 링버퍼로 유지
// =============================================================================

class ReplayBuffer {
    constructor(maxSize = 100) {
        this.maxSize = maxSize;
        this.items = [];
    }
    push(entry) {
        this.items.push(entry);
        if (this.items.length > this.maxSize) this.items.shift();
    }
    snapshot() { return [...this.items]; }
}

// =============================================================================
// ProcDump 옵션 (설치돼 있으면 한글 프로세스에 attach)
// =============================================================================

function maybeStartProcDump(procdumpPath, hwpPid, outDir) {
    if (!procdumpPath) return null;
    if (!fs.existsSync(procdumpPath)) {
        log('⚠', `ProcDump 경로 없음: ${procdumpPath} — 스킵`);
        return null;
    }
    try {
        const child = spawn(procdumpPath, [
            '-accepteula', '-ma', '-e', '-g', String(hwpPid), outDir,
        ], { detached: false, stdio: 'ignore', windowsHide: true });
        log('ℹ', `ProcDump attached pid=${hwpPid} → ${outDir}`);
        return child;
    } catch (e) {
        log('⚠', `ProcDump 기동 실패: ${e.message}`);
        return null;
    }
}

// =============================================================================
// Case 실행기들
// =============================================================================

async function execRibbonCase(caseObj, ctx, prng, appendLog) {
    const action = caseObj.action;
    win32.mouseClick(action.x, action.y);
    await sleep(280);

    const expectKind = caseObj.expect && caseObj.expect.kind;
    let interaction = 'none';
    let selectedIndex = -1;

    if (expectKind === 'popupAppears') {
        // dropdown — 70% 확률로 실제 항목 선택, 30% open-only
        const canSelect = caseObj.dropdownItemCount > 0;
        const shouldSelect = canSelect && prng() < 0.7;
        if (shouldSelect) {
            // Down 키로 N번째 항목까지 이동 후 Enter
            selectedIndex = Math.floor(prng() * caseObj.dropdownItemCount);
            for (let i = 0; i <= selectedIndex; i++) {
                try { await controller.pressKeys({ keys: 'Down' }); } catch (_) {}
                await sleep(30);
            }
            try { await controller.pressKeys({ keys: 'Enter' }); } catch (_) {}
            await sleep(200);
            // 선택이 새 다이얼로그를 열었을 가능성 → Esc로 닫음
            try { await controller.pressKeys({ keys: 'Escape' }); } catch (_) {}
            interaction = 'select';
        } else {
            try { await controller.pressKeys({ keys: 'Escape' }); } catch (_) {}
            try { await controller.pressKeys({ keys: 'Escape' }); } catch (_) {}
            interaction = 'open-only';
        }
    } else if (expectKind === 'dialogAppears') {
        // dialog — 70% 확률로 Tab/Space 내부 순회 후 Esc, 30%는 즉시 Esc.
        // Enter는 돌이킬 수 없는 설정 변경 누적 위험이 있어 쓰지 않음.
        // Ctrl+Tab: 다이얼로그 내부 탭 페이지(예: 저장 옵션 탭들) 전환.
        const canInteract = caseObj.dialogControlCount > 0;
        const shouldInteract = canInteract && prng() < 0.7;
        if (shouldInteract) {
            // 다이얼로그 내부 탭이 여러 개면 1~2회 전환
            if (caseObj.dialogTabCount > 1) {
                const tabSwitches = 1 + Math.floor(prng() * Math.min(2, caseObj.dialogTabCount - 1));
                for (let i = 0; i < tabSwitches; i++) {
                    try { await controller.pressKeys({ keys: 'Ctrl+Tab' }); } catch (_) {}
                    await sleep(80);
                }
            }
            // Tab으로 컨트롤 포커스 순회 — 2~6회
            const tabPresses = 2 + Math.floor(prng() * 5);
            for (let i = 0; i < tabPresses; i++) {
                try { await controller.pressKeys({ keys: 'Tab' }); } catch (_) {}
                await sleep(40);
            }
            // Space로 현재 포커스 컨트롤 활성화 — 0~2회
            const spacePresses = Math.floor(prng() * 3);
            for (let i = 0; i < spacePresses; i++) {
                try { await controller.pressKeys({ keys: 'Space' }); } catch (_) {}
                await sleep(80);
            }
            // 정리 — Esc 한 번은 취소, 한 번 더는 잔재 팝업 닫기
            try { await controller.pressKeys({ keys: 'Escape' }); } catch (_) {}
            await sleep(100);
            try { await controller.pressKeys({ keys: 'Escape' }); } catch (_) {}
            interaction = 'navigate-cancel';
        } else {
            for (const key of (caseObj.teardownKeys || ['Escape'])) {
                try { await controller.pressKeys({ keys: key }); } catch (_) {}
                await sleep(100);
            }
            interaction = 'cancel';
        }
    }

    const nextState = inferStateAfter(caseObj, ctx.state);
    if (nextState !== ctx.state) {
        ctx.transition(nextState, `ribbon:${caseObj.id}`);
    }
    appendLog({
        kind: 'ribbon', caseId: caseObj.id, tab: caseObj.tab, name: caseObj.shortName,
        category: caseObj.category, interaction,
        selectedIndex: selectedIndex >= 0 ? selectedIndex : undefined,
    });
}

async function execTypeText(prng, appendLog) {
    // 단락 단위 입력 — 4~6개 샘플을 조합해 ~500~900자를 한 번에 넣어 문서를 빠르게 키움.
    // 끝에 줄바꿈도 넣어 단락이 쌓이도록.
    const sentences = 4 + Math.floor(prng() * 3);
    const parts = [];
    for (let i = 0; i < sentences; i++) parts.push(pickOne(TEXT_SAMPLES, prng));
    const text = parts.join('') + '\n';
    await controller.typeText({ text, useClipboard: true });
    appendLog({ kind: 'typeText', bytes: text.length, preview: text.slice(0, 40) });
}

async function execScenario(map, ctx, prng, appendLog) {
    const scenario = pickOne(SCENARIOS, prng);
    const result = await runScenario(scenario, map, ctx);
    appendLog({ kind: 'scenario', id: scenario.id, status: result.status, error: result.error });
}

async function execUndoRedo(prng, appendLog) {
    const keys = pickOne(UNDO_REDO_KEYS, prng);
    try { await controller.pressKeys({ keys }); } catch (_) {}
    appendLog({ kind: 'undoRedo', keys });
}

async function execNavigate(prng, ctx, appendLog) {
    // 50% 확률로 COM을 이용해 표/도형 내부 list로 직접 caret 이동.
    // 점프 성공 후 즉시 state 확인 — 표면 F5 burst까지 이어서 실행해
    // "selection 상태에서 추가 액션" 코드 경로를 확실히 stress한다.
    if (hwpCom.isAvailable() && prng() < 0.5) {
        const listIdx = 1 + Math.floor(prng() * 4);
        const ok = hwpCom.moveToList(listIdx);
        let landed = null;
        let f5Taps = 0;
        if (ok) {
            await sleep(80);
            const snap = hwpCom.getStateSnapshot();
            landed = hwpCom.ctrlIdToContextState(snap);
            if (landed === STATES.IN_TABLE_CELL) {
                f5Taps = await execTableSelectionBurst(prng);
                ctx.transition(STATES.IN_TABLE_CELL, `com-setpos-list${listIdx}`);
            } else if (landed) {
                ctx.transition(landed, `com-setpos-list${listIdx}`);
            }
        }
        appendLog({ kind: 'navigate', via: 'com-setpos', list: listIdx, ok, landed, f5Taps });
        return;
    }
    const keys = pickOne(NAV_KEYS, prng);
    try { await controller.pressKeys({ keys }); } catch (_) {}
    appendLog({ kind: 'navigate', keys });
}

/**
 * 표 안에 있을 때만 유효한 "선택 모드 burst".
 *   F5 × 1: 셀 선택
 *   F5 × 2: 셀 확장 모드
 *   F5 × 3: 표 전체 선택
 * 이후 실행되는 ribbon/scenario 액션은 선택 상태에서 분기된 코드를 탄다.
 */
async function execTableSelectionBurst(prng) {
    const taps = 1 + Math.floor(prng() * 3);
    for (let i = 0; i < taps; i++) {
        try { await controller.pressKeys({ keys: 'F5' }); } catch (_) {}
        await sleep(60);
    }
    return taps;
}

async function execContextMenu(ctx, prng, appendLog) {
    // 본문 대략 중앙 근처에서 우클릭, 아무것도 고르지 않고 Esc로 닫기.
    // 좌표는 hwpElement.boundingRect에서 읽는다 (win32에 GetWindowRect 래퍼 없음).
    try {
        sessionModule.refreshHwpElement();
        const session = sessionModule.getSession();
        const el = session && session.hwpElement;
        let x = 600, y = 400;
        if (el) {
            try {
                const r = el.boundingRect;
                x = Math.round((r.left + r.right) / 2);
                y = Math.round((r.top + r.bottom) / 2);
            } catch (_) {}
        }
        win32.mouseClick(x, y, { rightClick: true });
        await sleep(300);
        await controller.pressKeys({ keys: 'Escape' });
        appendLog({ kind: 'contextMenu', x, y, action: 'open-close' });
    } catch (e) {
        appendLog({ kind: 'contextMenu', error: e.message });
    }
}

// =============================================================================
// 메인 러너
// =============================================================================

/**
 * @param {object} options
 * @param {string} options.mapPath
 * @param {number} options.durationMs
 * @param {number} [options.seed]
 * @param {boolean} [options.skipSeed]
 * @param {string}  [options.procdumpPath]
 * @param {string}  [options.outDir]   — 결과 디렉터리 (기본 test-results/stress-<ts>)
 * @param {string}  [options.product]  — 기본 'hwp'
 */
async function runStress(options) {
    const {
        mapPath,
        durationMs,
        seed = Date.now() >>> 0,
        skipSeed = false,
        procdumpPath = null,
        product = 'hwp',
        paceMs = 350,  // iter 당 최소 간격 — HWP가 이벤트를 처리할 시간 확보
    } = options;

    const runId = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const outDir = options.outDir || path.join('test-results', `stress-${runId}`);
    fs.mkdirSync(outDir, { recursive: true });
    fs.mkdirSync(path.join(outDir, 'crashes'), { recursive: true });

    const map = loadMap(mapPath);
    const cases = mapToCases(map);
    log('ℹ', `맵 로드: ${cases.length} cases`);

    // 출력 스트림
    const actionsFp = fs.openSync(path.join(outDir, 'actions.ndjson'), 'a');
    const heartbeatFp = fs.openSync(path.join(outDir, 'heartbeat.log'), 'a');

    const appendAction = (entry) => {
        try {
            fs.writeSync(actionsFp, JSON.stringify(entry) + '\n');
        } catch (_) {}
    };
    const heartbeat = (msg) => {
        try {
            fs.writeSync(heartbeatFp, `[${new Date().toISOString()}] ${msg}\n`);
        } catch (_) {}
    };

    // Config 저장
    fs.writeFileSync(path.join(outDir, 'config.json'), JSON.stringify({
        runId, mapPath, durationMs, seed, skipSeed, procdumpPath, product,
        startedAt: new Date().toISOString(),
    }, null, 2));

    // HWP 연결 — winax COM을 먼저 띄워 우리 전용 HWP 인스턴스를 확보한 후
    // UIA를 같은 PID로 attach. 이렇게 해야 winax COM과 UIA가 "같은 HWP"를
    // 조작/조회하므로 state 동기화가 실제로 동작한다.
    const beforePids = listHwpPids();
    log('▶', `${product} — winax COM 기반 HWP 기동 (기존 Hwp.exe ${beforePids.length}개)`);
    const comObj = hwpCom.init();
    let ourPid = null;
    if (comObj) {
        // 새 HWP 프로세스가 올라오길 기다림
        for (let i = 0; i < 25; i++) {
            await sleep(400);
            const after = listHwpPids();
            const diff = after.filter(p => !beforePids.includes(p));
            if (diff.length > 0) { ourPid = diff[0]; break; }
        }
        if (ourPid) {
            log('✓', `winax HWP PID=${ourPid} — UIA attach 시도`);
            try {
                await controller.attach({ product, pid: ourPid });
                log('✓', 'UIA attach 성공 (winax와 동일 인스턴스)');
            } catch (e) {
                log('⚠', `UIA attach 실패(pid=${ourPid}): ${e.message} — 일반 attach 폴백`);
                ourPid = null;
            }
        } else {
            log('⚠', 'winax init 후 새 PID 미탐지 — 일반 attach 폴백');
        }
    }
    if (!ourPid) {
        // winax 미사용 또는 새 PID 탐지 실패 — 기존 attach/launch 경로
        try {
            await controller.attach({ product });
            log('✓', 'attach 성공');
        } catch (e) {
            log('⚠', `attach 실패: ${e.message} — launch로 새 인스턴스 기동`);
            const info = await controller.launch({ product, timeoutMs: 30000 });
            log('✓', `launch 성공: pid=${info.pid}`);
        }
    }
    await controller.setForeground();
    await sleep(400);

    // 창 최대화 — 맵의 clickX/Y는 최대화 상태에서 캡처된 좌표이므로
    // 최대화되지 않으면 리본 버튼이 화면 밖으로 잘려 엉뚱한 곳을 클릭하게 된다.
    try {
        const sess = sessionModule.getSession();
        const hwnd = sess && sess.hwpProcess && sess.hwpProcess.hwnd;
        if (hwnd) {
            win32.maximizeWindow(hwnd);
            await sleep(600);
            log('🗖', '창 최대화');
        }
    } catch (e) {
        log('⚠', `최대화 실패 (계속 진행): ${e.message}`);
    }

    if (hwpCom.isAvailable()) {
        log('🔌', 'HWP COM 연결됨 — 실시간 state 동기화 활성화');
    } else {
        log('ℹ', 'HWP COM 미사용 (winax 초기화 실패) — 추정 방식 fallback');
    }

    // ProcDump
    let procdumpChild = null;
    const session = sessionModule.getSession();
    const hwpPid = session && session.hwpProcess && session.hwpProcess.pid;
    if (procdumpPath && hwpPid) {
        procdumpChild = maybeStartProcDump(procdumpPath, hwpPid, path.join(outDir, 'crashes'));
    }

    // Cancel watcher (ESC×2)
    let cancelled = false;
    const watcher = new CancelWatcher(() => {
        cancelled = true;
        log('🛑', 'ESC × 2 감지 — 다음 체크포인트에서 종료');
    });
    watcher.start();

    // 상태 추적
    const ctx = new StressContext();
    const prng = mulberry32(seed);
    const replay = new ReplayBuffer(100);

    // 시딩
    if (!skipSeed) {
        log('▶', '문서 시딩');
        await seeder.seed(map, ctx, runId);
    } else {
        log('ℹ', '시딩 생략 (--skip-seed)');
        await ctx.forceReturnToBody();
    }

    // 메인 루프
    const stats = {
        iterations: 0, crashes: 0, hangs: 0,
        byKind: {}, errors: 0,
        startedAt: Date.now(),
    };
    const deadline = stats.startedAt + durationMs;
    let lastHeartbeat = Date.now();
    let lastReseedAt = Date.now();

    const saveProgress = () => {
        try {
            fs.writeFileSync(path.join(outDir, 'progress.json'), JSON.stringify({
                ...stats,
                state: ctx.snapshot(),
                elapsedMs: Date.now() - stats.startedAt,
                lastIter: stats.iterations,
            }, null, 2));
        } catch (_) {}
    };

    try {
        while (Date.now() < deadline && !cancelled) {
            stats.iterations++;
            const iter = stats.iterations;
            const t0 = Date.now();

            // 매 iter 시작에 COM state sync — 매번 ~10-30ms이지만 selection
            // 기반 액션(F5 burst 등)이 실제로 의미있게 동작하려면 state가 stale이면 안 된다.
            syncStateFromCOM(ctx);

            // 매 5 iter마다 환경 복구 — 포커스/최대화/잔여 모달 정리.
            if (iter % 5 === 1) {
                await restoreEnvironment();
            }

            // 매 iter 전 health check
            const h = checkAlive();
            if (!h.alive) {
                log('💥', `health check 실패: ${h.reason}`);
                await handleCrash({ reason: h.reason, replay, outDir, iter, ctx });
                stats.crashes++;
                if (!await tryReattach(product)) {
                    log('✋', '재연결 실패 — 루프 종료');
                    break;
                }
                continue;
            }

            // 가중 샘플링으로 액션 타입 선택
            const actionType = pickFromMix(ACTION_MIX, prng);

            try {
                const entry = { iter, t: Date.now(), actionType, state: ctx.state };

                // 표 안에 있고 '기능 실행'류 액션일 때 30% 확률로 F5 선택 모드 burst.
                // 이후 실행되는 액션이 "셀/표 선택됨" 코드 경로를 타게 된다.
                const majorAction = (actionType === 'ribbon' || actionType === 'scenario' || actionType === 'contextMenu');
                if (majorAction && ctx.state === STATES.IN_TABLE_CELL && prng() < 0.3) {
                    const taps = await execTableSelectionBurst(prng);
                    entry.preBurst = { kind: 'table-f5', taps };
                }

                switch (actionType) {
                    case 'ribbon': {
                        const weighted = weightedForState(cases, ctx.state);
                        const c = pickWeighted(weighted, prng);
                        if (c) await execRibbonCase(c, ctx, prng, e => Object.assign(entry, e));
                        else entry.skipped = 'no-weighted-cases';
                        break;
                    }
                    case 'typeText':    await execTypeText(prng, e => Object.assign(entry, e)); break;
                    case 'scenario':    await execScenario(map, ctx, prng, e => Object.assign(entry, e)); break;
                    case 'undoRedo':    await execUndoRedo(prng, e => Object.assign(entry, e)); break;
                    case 'navigate':    await execNavigate(prng, ctx, e => Object.assign(entry, e)); break;
                    case 'contextMenu': await execContextMenu(ctx, prng, e => Object.assign(entry, e)); break;
                }
                entry.durationMs = Date.now() - t0;
                stats.byKind[actionType] = (stats.byKind[actionType] || 0) + 1;
                appendAction(entry);
                replay.push(entry);
            } catch (e) {
                stats.errors++;
                const entry = { iter, t: Date.now(), actionType, state: ctx.state, error: e.message, durationMs: Date.now() - t0 };
                appendAction(entry);
                replay.push(entry);
                // 에러 후 본문 복귀 시도
                try { await ctx.forceReturnToBody(); } catch (_) {}
            }

            // 매 10 iter마다 state drift 보정
            if (iter % 10 === 0 && ctx.state !== STATES.BODY) {
                try { await ctx.forceReturnToBody(); } catch (_) {}
            }

            // 매 100 iter마다 progress 저장
            if (iter % 100 === 0) saveProgress();

            // 5분마다 heartbeat
            if (Date.now() - lastHeartbeat > 5 * 60 * 1000) {
                heartbeat(`alive, iter=${iter}, crashes=${stats.crashes}, state=${ctx.state}`);
                log('💓', `iter=${iter}, 경과 ${Math.round((Date.now()-stats.startedAt)/60000)}분, crashes=${stats.crashes}`);
                lastHeartbeat = Date.now();
            }

            // 30분마다 경량 보강 시딩
            if (Date.now() - lastReseedAt > 30 * 60 * 1000) {
                await seeder.reseed(map, ctx);
                lastReseedAt = Date.now();
            }

            // iter 최소 간격 — HWP 이벤트 처리 여유. --pace-ms로 조절.
            const elapsed = Date.now() - t0;
            if (elapsed < paceMs) await sleep(paceMs - elapsed);
        }
    } finally {
        watcher.stop();
        saveProgress();

        const summary = {
            ...stats,
            endedAt: new Date().toISOString(),
            elapsedMs: Date.now() - stats.startedAt,
            cancelled,
            reason: cancelled ? 'user-cancel' : (Date.now() >= deadline ? 'budget-exhausted' : 'other'),
        };
        fs.writeFileSync(path.join(outDir, 'summary.json'), JSON.stringify(summary, null, 2));

        try { fs.closeSync(actionsFp); } catch (_) {}
        try { fs.closeSync(heartbeatFp); } catch (_) {}
        if (procdumpChild) { try { procdumpChild.kill(); } catch (_) {} }

        log('ℹ', `종료: ${stats.iterations} iters, ${stats.crashes} crashes, ${stats.errors} errors`);
        log('ℹ', `결과 디렉터리: ${outDir}`);
    }

    return outDir;
}

// =============================================================================
// Crash 처리 & 재연결
// =============================================================================

async function handleCrash({ reason, replay, outDir, iter, ctx }) {
    const ts = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const dir = path.join(outDir, 'crashes', `${iter}-${ts}`);
    try {
        fs.mkdirSync(dir, { recursive: true });
        fs.writeFileSync(path.join(dir, 'replay.ndjson'),
            replay.snapshot().map(e => JSON.stringify(e)).join('\n'));
        fs.writeFileSync(path.join(dir, 'context.json'), JSON.stringify({
            reason, iter, state: ctx.snapshot(), at: new Date().toISOString(),
        }, null, 2));
    } catch (_) {}
    log('💾', `크래시 저장: ${dir}`);
}

async function tryReattach(product, maxAttempts = 3) {
    for (let i = 0; i < maxAttempts; i++) {
        await sleep(2000);
        try {
            await controller.attach({ product });
            log('✓', '재연결 성공');
            return true;
        } catch (e) {
            log('⚠', `재연결 ${i + 1}/${maxAttempts} 실패: ${e.message}`);
        }
    }
    return false;
}

module.exports = { runStress, mulberry32 };
