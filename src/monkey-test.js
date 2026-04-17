/**
 * Monkey Testing Engine for Hancom Office 2027
 * 지정된 시간 동안 랜덤하게 UI 기능을 실행하여 크래시/행을 탐지한다.
 *
 * 주요 기능:
 *  - 가중치 기반 랜덤 액션 선택 (메뉴, 버튼, 텍스트, 키보드)
 *  - 블랙리스트로 위험 액션 제외
 *  - 크래시 감지 시 자동 복구 + Jira 리포트
 *  - 시드 기반 PRNG으로 재현 가능
 *  - 상세 액션 로그 + 결과 리포트
 */
'use strict';

const fs = require('fs');
const controller = require('./hwp-controller');
const sessionModule = require('./session');

// =============================================================================
// Seeded PRNG (mulberry32) — 동일 시드로 동일 액션 시퀀스 재현 가능
// =============================================================================
function mulberry32(seed) {
    return function () {
        seed |= 0;
        seed = (seed + 0x6D2B79F5) | 0;
        let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
        t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
        return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
}

// =============================================================================
// 액션 타입 및 가중치
// =============================================================================
const DEFAULT_ACTION_WEIGHTS = [
    { type: 'click_menu',    weight: 30 },
    { type: 'click_button',  weight: 20 },
    { type: 'type_text',     weight: 10 },
    { type: 'press_keys',    weight: 15 },
    { type: 'explore_ui',    weight: 5 },
    { type: 'handle_dialog', weight: 20 },
];

// 안전한 키 조합 목록
const SAFE_KEY_COMBOS = [
    'Ctrl+Z', 'Ctrl+Y', 'Ctrl+A', 'Ctrl+C', 'Ctrl+V',
    'Ctrl+Home', 'Ctrl+End', 'Ctrl+B', 'Ctrl+I', 'Ctrl+U',
    'F1', 'F5', 'F7', 'F9',
    'Tab', 'Enter', 'Escape', 'Space',
    'PageUp', 'PageDown',
    'Up', 'Down', 'Left', 'Right',
    'Home', 'End', 'Delete', 'Backspace',
    'Shift+Left', 'Shift+Right', 'Shift+Up', 'Shift+Down',
    'Ctrl+Shift+Left', 'Ctrl+Shift+Right',
];

// 기본 블랙리스트 (메뉴/버튼 이름에 포함되면 스킵)
const DEFAULT_BLACKLIST = [
    '삭제', '지우기', '포맷', '초기화', '비우기',
    '인쇄', '프린트', 'print', '출력',
    '보내기', '메일', 'mail', 'email',
    '종료', '닫기', 'exit', 'quit', 'close',
    '저장', 'save', '다른 이름',
    '업데이트', '업그레이드', 'update',
    '등록', '구매', '결제', '라이선스',
    '설치', 'install', '제거', 'uninstall',
    '내보내기', 'export', '가져오기', 'import',
    '불러오기', '열기', 'open',
];

// 랜덤 텍스트 샘플
const TEXT_SAMPLES = [
    'Hello World',
    '테스트 문서입니다.',
    '한컴오피스 2027 몽키 테스트',
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit.',
    '가나다라마바사 아자차카타파하',
    '1234567890 !@#$%^&*()',
    'The quick brown fox jumps over the lazy dog.',
    '이 문장은 자동 테스트에 의해 입력되었습니다.',
    '문단 나누기 테스트\n두 번째 줄입니다.',
    'ABC abc XYZ xyz 가나다 라마바',
];

// =============================================================================
// MonkeyTester 클래스
// =============================================================================
class MonkeyTester {
    /**
     * @param {object} options
     * @param {string} [options.product='hwp']         대상 제품
     * @param {number} [options.durationMs=300000]      실행 시간 (기본 5분)
     * @param {object} [options.interval]               액션 간 딜레이 { min, max } ms
     * @param {number} [options.seed]                   PRNG 시드 (재현용)
     * @param {string[]} [options.blacklist]            금지 키워드 목록
     * @param {number} [options.maxCrashes=5]           최대 크래시 횟수 (초과 시 중단)
     * @param {object[]} [options.actionWeights]        액션 가중치 오버라이드
     */
    constructor(options = {}) {
        this.product = options.product || 'hwp';
        this.durationMs = options.durationMs || 5 * 60 * 1000;
        this.interval = options.interval || { min: 300, max: 1000 };
        this.seed = options.seed || Date.now();
        this.blacklist = options.blacklist || DEFAULT_BLACKLIST;
        this.maxCrashes = options.maxCrashes || 5;
        this.actionWeights = options.actionWeights || DEFAULT_ACTION_WEIGHTS;

        this.rng = mulberry32(this.seed);
        this.running = false;
        this.paused = false;

        this.stats = {
            seed: this.seed,
            product: this.product,
            durationMs: this.durationMs,
            totalActions: 0,
            successActions: 0,
            failedActions: 0,
            crashes: 0,
            recoveries: 0,
            actionBreakdown: {},
            startTime: null,
            endTime: null,
            actions: [],
            crashRecords: [],
        };
    }

    // -------------------------------------------------------------------------
    // Public API
    // -------------------------------------------------------------------------

    /**
     * 몽키 테스트 시작 (async, 종료 시 리포트 반환)
     */
    async start() {
        this.running = true;
        this.stats.startTime = new Date().toISOString();

        // 제품 연결
        try {
            await controller.attach({ product: this.product });
            this._log('session', 'attached', { product: this.product });
        } catch (_e) {
            try {
                await controller.launch({ product: this.product });
                this._log('session', 'launched', { product: this.product });
            } catch (e) {
                this._log('session', 'launch_failed', { error: e.message });
                this.running = false;
                this.stats.endTime = new Date().toISOString();
                return this.getReport();
            }
        }

        // 런처 닫기: ESC
        try {
            await controller.setForeground();
            await controller.pressKeys({ keys: 'Escape' });
            await this._sleep(2000);
        } catch (_e) { /* 이미 편집 모드일 수 있음 */ }

        const deadline = Date.now() + this.durationMs;

        while (this.running && Date.now() < deadline) {
            // 일시정지 상태면 대기
            while (this.paused && this.running) {
                await this._sleep(500);
            }
            if (!this.running) break;

            try {
                // 프로세스 상태 확인
                const status = await controller.getStatus();
                if (!status.alive) {
                    this.stats.crashes++;
                    this.stats.crashRecords.push({
                        time: new Date().toISOString(),
                        lastAction: this.stats.actions[this.stats.actions.length - 1] || null,
                    });
                    this._log('crash', 'process_dead', {
                        crashCount: this.stats.crashes,
                    });

                    if (this.stats.crashes >= this.maxCrashes) {
                        this._log('session', 'max_crashes_reached', {
                            maxCrashes: this.maxCrashes,
                        });
                        break;
                    }

                    // 복구: 재실행
                    await this._sleep(3000);
                    try {
                        await controller.launch({ product: this.product });
                        await this._sleep(2000);
                        await controller.pressKeys({ keys: 'Escape' });
                        await this._sleep(2000);
                        this.stats.recoveries++;
                        this._log('session', 'recovered', {});
                    } catch (e) {
                        this._log('session', 'recovery_failed', { error: e.message });
                        break;
                    }
                    continue;
                }

                if (!status.responding) {
                    this._log('warning', 'not_responding', {});
                    await this._sleep(5000);
                    continue;
                }

                // 랜덤 액션 실행
                const actionType = this._pickActionType();
                await this._executeAction(actionType);

                // 랜덤 딜레이
                const delay = this._randomInt(this.interval.min, this.interval.max);
                await this._sleep(delay);

            } catch (e) {
                this.stats.failedActions++;
                this._log('error', 'action_exception', { error: e.message });
                await this._sleep(1000);
            }
        }

        this.running = false;
        this.stats.endTime = new Date().toISOString();
        return this.getReport();
    }

    /**
     * 테스트 중단
     */
    stop() {
        this.running = false;
    }

    /**
     * 일시정지 / 재개
     */
    pause() { this.paused = true; }
    resume() { this.paused = false; }

    /**
     * 현재 상태 조회
     */
    getStatus() {
        return {
            running: this.running,
            paused: this.paused,
            seed: this.seed,
            product: this.product,
            elapsed: this.stats.startTime
                ? Date.now() - new Date(this.stats.startTime).getTime()
                : 0,
            remaining: this.stats.startTime
                ? Math.max(0, this.durationMs - (Date.now() - new Date(this.stats.startTime).getTime()))
                : this.durationMs,
            totalActions: this.stats.totalActions,
            crashes: this.stats.crashes,
            failedActions: this.stats.failedActions,
        };
    }

    /**
     * 결과 리포트 생성
     */
    getReport() {
        const elapsed = this.stats.startTime && this.stats.endTime
            ? new Date(this.stats.endTime).getTime() - new Date(this.stats.startTime).getTime()
            : 0;

        return {
            summary: {
                product: this.product,
                seed: this.seed,
                durationMs: elapsed,
                durationFormatted: this._formatDuration(elapsed),
                totalActions: this.stats.totalActions,
                successActions: this.stats.successActions,
                failedActions: this.stats.failedActions,
                crashes: this.stats.crashes,
                recoveries: this.stats.recoveries,
                actionBreakdown: this.stats.actionBreakdown,
            },
            crashRecords: this.stats.crashRecords,
            actions: this.stats.actions,
            startTime: this.stats.startTime,
            endTime: this.stats.endTime,
        };
    }

    /**
     * 리포트를 JSON 파일로 저장
     */
    exportReport(outputPath) {
        const report = this.getReport();
        const dir = outputPath
            ? require('path').dirname(outputPath)
            : '.';
        const filePath = outputPath ||
            `reports/monkey-test-${this.product}-${Date.now()}.json`;

        try { fs.mkdirSync(require('path').dirname(filePath), { recursive: true }); } catch (_) {}
        fs.writeFileSync(filePath, JSON.stringify(report, null, 2), 'utf8');
        return filePath;
    }

    // -------------------------------------------------------------------------
    // Internal: 액션 선택
    // -------------------------------------------------------------------------

    _pickActionType() {
        const totalWeight = this.actionWeights.reduce((s, a) => s + a.weight, 0);
        let r = this.rng() * totalWeight;
        for (const action of this.actionWeights) {
            r -= action.weight;
            if (r <= 0) return action.type;
        }
        return this.actionWeights[0].type;
    }

    // -------------------------------------------------------------------------
    // Internal: 액션 실행
    // -------------------------------------------------------------------------

    async _executeAction(actionType) {
        this.stats.totalActions++;
        this.stats.actionBreakdown[actionType] =
            (this.stats.actionBreakdown[actionType] || 0) + 1;

        switch (actionType) {
            case 'click_menu':   return this._actionClickMenu();
            case 'click_button': return this._actionClickButton();
            case 'type_text':    return this._actionTypeText();
            case 'press_keys':   return this._actionPressKeys();
            case 'explore_ui':   return this._actionExploreUi();
            case 'handle_dialog': return this._actionHandleDialog();
            default:
                this._log(actionType, 'unknown_action', {});
        }
    }

    async _actionClickMenu() {
        // 한글 2027은 커스텀 리본 UI로 UIA에서 하위 메뉴 항목을 탐지할 수 없음.
        // 키보드 네비게이션으로 메뉴를 탐색한다:
        //   Alt → 메뉴바 활성화 → Right/Left로 탭 이동 → Down으로 하위 진입
        //   → Down으로 항목 이동 → Enter로 실행 또는 Escape로 닫기

        // 메뉴별 액세스 키 (Alt 후 누를 키) — UIA 이름에서 추출
        const SAFE_MENUS = [
            { name: '편집', key: 'E' },
            { name: '보기', key: 'U' },
            { name: '입력', key: 'D' },
            { name: '서식', key: 'J' },
            { name: '쪽',   key: 'W' },
            { name: '검토', key: 'H' },
            { name: '도구', key: 'K' },
        ];

        // 라운드 로빈으로 다양한 메뉴 탐색
        if (!this._menuTabIdx) this._menuTabIdx = 0;
        const menu = SAFE_MENUS[this._menuTabIdx % SAFE_MENUS.length];
        this._menuTabIdx++;
        const tabName = menu.name;

        // 하위 항목 중 몇 번째를 선택할지 (1~8개 아래로)
        const subDepth = this._randomInt(1, 8);

        try {
            // 1. Alt 키로 메뉴바 활성화 후 액세스 키로 직접 이동
            await controller.pressKeys({ keys: 'Alt' });
            await this._sleep(300);
            await controller.pressKeys({ keys: menu.key });
            await this._sleep(400);

            this._log('click_menu', 'navigating', {
                name: `${tabName} > 항목 ${subDepth}번째`,
            });

            // 4. Down으로 원하는 항목까지 이동
            for (let i = 1; i < subDepth; i++) {
                await controller.pressKeys({ keys: 'Down' });
                await this._sleep(100);
            }

            // 5. 70% 확률로 Enter(실행), 30% Escape(닫기)
            if (this.rng() < 0.7) {
                await controller.pressKeys({ keys: 'Enter' });
                this._log('click_menu', 'executed', {
                    name: `${tabName} > 항목 ${subDepth}번째`,
                });
                this.stats.successActions++;
                await this._sleep(800);

                // 다이얼로그가 열렸으면 내부 조작 후 닫기
                await this._actionHandleDialog();
            } else {
                await controller.pressKeys({ keys: 'Escape' });
                this._log('click_menu', 'browsed', {
                    name: `${tabName} > 항목 ${subDepth}번째 (닫음)`,
                });
                this.stats.successActions++;
            }

            await this._sleep(200);
            // 메뉴가 아직 열려있을 수 있으므로 ESC 한번 더
            await controller.pressKeys({ keys: 'Escape' });

        } catch (e) {
            this.stats.failedActions++;
            this._log('click_menu', 'failed', { name: tabName, error: e.message });
            try {
                await controller.pressKeys({ keys: 'Escape' });
                await this._sleep(100);
                await controller.pressKeys({ keys: 'Escape' });
            } catch (_) {}
        }
    }

    async _actionClickButton() {
        const result = await controller.findElement({
            controlType: 'Button',
            maxResults: 50,
        });

        const elements = (result.elements || []).filter(
            el => el.name && !this._isBlacklisted(el.name) && el.isEnabled
        );

        if (elements.length === 0) {
            this._log('click_button', 'no_buttons', {});
            return;
        }

        const target = elements[this._randomInt(0, elements.length - 1)];
        this._log('click_button', 'clicking', { name: target.name, index: target.index });

        try {
            await controller.clickElement({ index: target.index });
            this.stats.successActions++;
        } catch (e) {
            this.stats.failedActions++;
            this._log('click_button', 'failed', { name: target.name, error: e.message });
        }
    }

    async _actionTypeText() {
        const text = TEXT_SAMPLES[this._randomInt(0, TEXT_SAMPLES.length - 1)];
        this._log('type_text', 'typing', { text: text.substring(0, 30), length: text.length });

        try {
            // 클립보드 방식으로 유니코드 텍스트 입력
            await controller.typeText({ text, useClipboard: true });
            this.stats.successActions++;
        } catch (e) {
            this.stats.failedActions++;
            this._log('type_text', 'failed', { error: e.message });
        }
    }

    async _actionPressKeys() {
        const combo = SAFE_KEY_COMBOS[this._randomInt(0, SAFE_KEY_COMBOS.length - 1)];
        this._log('press_keys', 'pressing', { keys: combo });

        try {
            await controller.pressKeys({ keys: combo });
            this.stats.successActions++;
        } catch (e) {
            this.stats.failedActions++;
            this._log('press_keys', 'failed', { keys: combo, error: e.message });
        }
    }

    async _actionExploreUi() {
        this._log('explore_ui', 'scanning', {});

        try {
            const tree = await controller.getUiTree({ depth: 2 });
            this.stats.successActions++;
            this._log('explore_ui', 'scanned', { totalElements: tree.totalElements });
        } catch (e) {
            this.stats.failedActions++;
            this._log('explore_ui', 'failed', { error: e.message });
        }
    }

    async _actionHandleDialog() {
        this._log('handle_dialog', 'checking', {});

        try {
            // 다이얼로그가 실제로 열려있는지 확인:
            // 포커스된 요소가 메인 편집기(HwpMainEditWnd/paragraph)가 아니면 다이얼로그로 판단
            const focused = await controller.getFocusedElement();
            const inMainEditor = focused && (
                (focused.className || '').includes('HwpMainEditWnd') ||
                focused.name === 'paragraph' ||
                (focused.className || '').includes('FrameWindowImpl')
            );

            if (inMainEditor || !focused || !focused.controlType) {
                this._log('handle_dialog', 'no_dialog_open', {});
                this.stats.successActions++;
                return;
            }

            this._log('handle_dialog', 'dialog_detected', {
                name: focused.name, type: focused.controlType,
            });

            // 다이얼로그 내부를 Tab 키로 순회하며 컨트롤 조작
            const maxTabs = this._randomInt(3, 10);
            let interactions = 0;

            for (let i = 0; i < maxTabs; i++) {
                // Tab으로 다음 컨트롤 이동
                await controller.pressKeys({ keys: 'Tab' });
                await this._sleep(200);

                // 현재 포커스된 요소 확인
                const focused = await controller.getFocusedElement();
                if (!focused || !focused.controlType) continue;

                const ctName = focused.controlType;
                const elName = focused.name || '(unnamed)';

                // 블랙리스트 체크
                if (this._isBlacklisted(elName)) continue;

                try {
                    switch (ctName) {
                        case 'CheckBox':
                            // Space로 토글
                            await controller.pressKeys({ keys: 'Space' });
                            this._log('handle_dialog', 'toggled', { name: elName, type: ctName });
                            interactions++;
                            break;

                        case 'RadioButton':
                            // Space 또는 화살표로 선택
                            await controller.pressKeys({ keys: 'Space' });
                            this._log('handle_dialog', 'selected', { name: elName, type: ctName });
                            interactions++;
                            break;

                        case 'ComboBox':
                        case 'List':
                            // 화살표로 항목 변경
                            const moves = this._randomInt(1, 4);
                            for (let j = 0; j < moves; j++) {
                                await controller.pressKeys({ keys: this.rng() < 0.5 ? 'Down' : 'Up' });
                                await this._sleep(100);
                            }
                            this._log('handle_dialog', 'combo_changed', { name: elName, moves });
                            interactions++;
                            break;

                        case 'Edit':
                        case 'Spinner':
                            // 값 입력/변경
                            if (ctName === 'Spinner' || elName.includes('pt') || elName.includes('mm') || elName.includes('%')) {
                                // 숫자 스피너: 화살표로 값 조정
                                const spins = this._randomInt(1, 3);
                                for (let j = 0; j < spins; j++) {
                                    await controller.pressKeys({ keys: this.rng() < 0.5 ? 'Up' : 'Down' });
                                    await this._sleep(100);
                                }
                                this._log('handle_dialog', 'spinner_changed', { name: elName, spins });
                            } else {
                                // 텍스트 필드: 값 입력
                                await controller.pressKeys({ keys: 'Ctrl+A' });
                                const val = String(this._randomInt(1, 100));
                                await controller.typeText({ text: val, useClipboard: false });
                                this._log('handle_dialog', 'text_entered', { name: elName, value: val });
                            }
                            interactions++;
                            break;

                        case 'Tab':
                        case 'TabItem':
                            // 탭 클릭
                            await controller.pressKeys({ keys: 'Space' });
                            this._log('handle_dialog', 'tab_switched', { name: elName });
                            interactions++;
                            break;

                        case 'Button':
                            // 버튼은 확인/취소일 수 있으므로 스킵 (마지막에 처리)
                            break;

                        default:
                            // 기타 컨트롤은 Space로 활성화 시도
                            if (this.rng() < 0.3) {
                                await controller.pressKeys({ keys: 'Space' });
                                this._log('handle_dialog', 'activated', { name: elName, type: ctName });
                                interactions++;
                            }
                    }
                    await this._sleep(200);
                } catch (e) {
                    this._log('handle_dialog', 'control_failed', { name: elName, error: e.message });
                }
            }

            if (interactions === 0) {
                this._log('handle_dialog', 'no_dialog_controls', {});
            } else {
                this._log('handle_dialog', 'interacted', { count: interactions });
            }

            // 조작 완료 후 50% 확률로 확인/취소
            await this._sleep(300);
            if (this.rng() < 0.5) {
                await controller.pressKeys({ keys: 'Enter' });
                this._log('handle_dialog', 'confirmed', {});
            } else {
                await controller.pressKeys({ keys: 'Escape' });
                this._log('handle_dialog', 'cancelled', {});
            }

            this.stats.successActions++;
        } catch (e) {
            this.stats.failedActions++;
            this._log('handle_dialog', 'failed', { error: e.message });
            try { await controller.pressKeys({ keys: 'Escape' }); } catch (_) {}
        }
    }

    // -------------------------------------------------------------------------
    // Internal: 유틸리티
    // -------------------------------------------------------------------------

    _isBlacklisted(name) {
        if (!name) return false;
        const lower = name.toLowerCase();
        return this.blacklist.some(kw => lower.includes(kw.toLowerCase()));
    }

    _randomInt(min, max) {
        return Math.floor(this.rng() * (max - min + 1)) + min;
    }

    _sleep(ms) {
        return new Promise(r => setTimeout(r, ms));
    }

    _log(type, action, data) {
        const entry = {
            seq: this.stats.actions.length + 1,
            time: new Date().toISOString(),
            type,
            action,
            ...data,
        };
        this.stats.actions.push(entry);

        // 실시간 콘솔 출력
        const elapsed = this.stats.startTime
            ? Math.floor((Date.now() - new Date(this.stats.startTime).getTime()) / 1000)
            : 0;
        const mm = String(Math.floor(elapsed / 60)).padStart(2, '0');
        const ss = String(elapsed % 60).padStart(2, '0');
        const tag = `[${mm}:${ss}]`;

        let detail = '';
        if (data.name)  detail = ` "${data.name}"`;
        if (data.keys)  detail = ` ${data.keys}`;
        if (data.text)  detail = ` "${data.text}"`;
        if (data.error) detail += ` ERROR: ${data.error}`;
        if (data.totalElements) detail += ` (${data.totalElements} elements)`;
        if (data.crashCount) detail += ` (crash #${data.crashCount})`;

        const icon = {
            click_menu: 'MENU', click_button: 'BTN ', type_text: 'TYPE',
            press_keys: 'KEY ', explore_ui: 'SCAN', handle_dialog: 'DLG ',
            crash: 'CRASH', warning: 'WARN', error: 'ERR ', session: 'SES ',
        }[type] || type.substring(0, 4).toUpperCase();

        process.stderr.write(`${tag} #${entry.seq} ${icon} ${action}${detail}\n`);
    }

    _formatDuration(ms) {
        const s = Math.floor(ms / 1000);
        const m = Math.floor(s / 60);
        const h = Math.floor(m / 60);
        if (h > 0) return `${h}h ${m % 60}m ${s % 60}s`;
        if (m > 0) return `${m}m ${s % 60}s`;
        return `${s}s`;
    }
}

// =============================================================================
// Singleton instance for MCP tool control
// =============================================================================
let _activeTester = null;

/**
 * 몽키 테스트 시작 (MCP 도구용)
 * 백그라운드에서 실행되며, 완료 시 리포트를 파일로 저장
 */
async function startMonkeyTest(options = {}) {
    if (_activeTester && _activeTester.running) {
        throw new Error('Monkey test is already running. Stop it first.');
    }

    _activeTester = new MonkeyTester(options);

    // 백그라운드에서 실행 (Promise 반환하지 않음)
    const testPromise = _activeTester.start().then(report => {
        try {
            const filePath = _activeTester.exportReport();
            report.reportFile = filePath;
        } catch (_) {}
        return report;
    });

    // 시작 정보 즉시 반환
    return {
        started: true,
        seed: _activeTester.seed,
        product: _activeTester.product,
        durationMs: _activeTester.durationMs,
        durationFormatted: _activeTester._formatDuration(_activeTester.durationMs),
        message: `Monkey test started. Use hwp_monkey_status to check progress.`,
        _promise: testPromise, // internal: for awaiting in tests
    };
}

/**
 * 몽키 테스트 중단
 */
function stopMonkeyTest() {
    if (!_activeTester || !_activeTester.running) {
        return { stopped: false, message: 'No monkey test is running.' };
    }

    _activeTester.stop();
    const report = _activeTester.getReport();
    try {
        const filePath = _activeTester.exportReport();
        report.reportFile = filePath;
    } catch (_) {}
    return { stopped: true, report };
}

/**
 * 몽키 테스트 상태 조회
 */
function getMonkeyStatus() {
    if (!_activeTester) {
        return { running: false, message: 'No monkey test has been started.' };
    }
    return _activeTester.getStatus();
}

module.exports = {
    MonkeyTester,
    startMonkeyTest,
    stopMonkeyTest,
    getMonkeyStatus,
    // 상수 (외부 커스터마이즈용)
    DEFAULT_ACTION_WEIGHTS,
    DEFAULT_BLACKLIST,
    SAFE_KEY_COMBOS,
    TEXT_SAMPLES,
};
