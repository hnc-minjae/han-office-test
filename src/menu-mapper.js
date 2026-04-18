/**
 * Menu Mapper — 한컴오피스 2027 메뉴 맵 크롤러
 *
 * UIA로 각 탭의 리본 항목을 직접 읽고, 각 항목을 invoke하여
 * 다이얼로그/컨트롤을 수집한다. 키보드 이동 없이 UIA 직접 접근 방식.
 */
'use strict';

const fs = require('fs');
const path = require('path');
const controller = require('./hwp-controller');
const sessionModule = require('./session');
const { TreeScope } = require('./uia');
const win32 = require('./win32');

const { PRODUCTS } = sessionModule;

const DELAY = { short: 150, medium: 300, long: 600, dialog: 1000 };
const MAX_TAB_CONTROLS = 60;

// 다이얼로그 프로빙에서 전체 스킵할 탭 — 보기 탭은 창 생성/분할/리본 토글 등
// 파괴적 UI 상태 변경 액션이 많고, 설정 다이얼로그가 거의 없음
const SKIP_PROBING_TABS = new Set(['보기']);

// 리본 항목 이름이 이 패턴 중 하나와 일치하면 다이얼로그 프로빙에서 스킵.
// 창 분할·새 창 생성·리본 최소화처럼 UI 상태를 영구 변경하는 액션을 차단.
const DANGEROUS_ITEM_PATTERNS = [
    /^새 창/, /^편집 화면 나누기/, /^창 배열/, /^창 전환/,
    /리본 최소화/, /리본 숨기/,
];

// 다이얼로그가 아닌 메인 윈도우 클래스
const MAIN_CLASSES = ['HwpMainEditWnd', 'FrameWindowImpl', 'HwpRulerWnd', 'Popup',
    'ToolBoxImpl', 'ToolBarImpl', 'StatusBarImpl', 'MenuBarImpl'];

// =============================================================================
// Dropdown probing — plan §5, 7계층 방어
// =============================================================================

const DD_LIMITS = {
    closeAttempts: 5,
    popupTreeMaxNodes: 500,
    popupTreeMaxDepth: 8,
    popupWaitMs: 1500,
    perDropdownBudgetMs: 10_000,
    perTabBudgetMs: 180_000,         // 일반 탭 3 min. 도형 탭은 재선택 오버헤드 감안.
    totalProbingBudgetMs: 900_000,
};

// 드롭다운 이름이 이 패턴 중 하나와 일치하면 프로빙 전체 스킵
const DANGEROUS_DROPDOWN_PATTERNS = [
    // Discovery 후 필요 시 추가
];

// Dialog-opener 판정 (plan §3에서 갱신: 한글 UI는 '...' 대신 '설정/모양/사용자 정의')
const DIALOG_OPENER_RE = /설정\s*:|사용자 정의|모양\s*:/;

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function log(icon, msg) { process.stderr.write(`${icon} ${msg}\n`); }

// =============================================================================
// MenuMapper
// =============================================================================

class MenuMapper {
    constructor(options = {}) {
        this.product = options.product || 'hwp';
        this.probeDialogs = options.probeDialogs !== false;
        this.probeDropdowns = options.probeDropdowns !== false;
        this.probeShapeTab = options.probeShapeTab !== false;
        this._recoveryInProgress = false; // plan §5 Layer 5

        const info = PRODUCTS[this.product];
        if (!info) throw new Error(`Unknown product: "${this.product}"`);
        this.productName = info.name;

        this.map = {
            product: this.product,
            productName: this.productName,
            version: '2027',
            mappedAt: null,
            tabs: {},
            toolbar: [],
            statusbar: [],
            stats: {
                totalTabs: 0,
                totalRibbonItems: 0,
                totalDialogs: 0,
                totalControls: 0,
                totalDropdowns: 0,
                totalDropdownItems: 0,
                dropdownErrors: 0,
                errors: 0,
            },
        };
    }

    // =========================================================================
    // Public API
    // =========================================================================

    async run() {
        this.map.mappedAt = new Date().toISOString();

        log('▶', `${this.productName} 연결 중...`);
        try {
            await controller.attach({ product: this.product });
        } catch (_) {
            throw new Error(`${this.productName}이(가) 실행 중이어야 합니다.`);
        }

        await controller.setForeground();
        await sleep(DELAY.medium);
        await controller.pressKeys({ keys: 'Escape' });
        await sleep(DELAY.long);

        // Step 1: 메뉴 탭 목록 수집
        log('▶', 'Step 1: 메뉴 탭 수집');
        const menuTabs = await this._collectMenuTabs();
        log('ℹ', `${menuTabs.length}개 탭: ${menuTabs.map(t => t.name).join(', ')}`);

        // Step 2: 각 탭의 리본 항목 수집
        for (const tab of menuTabs) {
            // "파일" 탭은 Backstage 패널을 열므로 리본 탐색에서 제외
            if (tab.name === '파일') {
                log('ℹ', `  "${tab.name}" — Backstage 패널 (리본 탐색 제외)`);
                this.map.tabs[tab.name] = { accessKey: tab.accessKey, uiaName: tab.uiaName, ribbonItems: [], note: 'backstage' };
                continue;
            }
            // 비활성 탭 건너뛰기
            if (!tab.isEnabled) {
                log('ℹ', `  "${tab.name}" — 비활성 탭 (건너뜀)`);
                this.map.tabs[tab.name] = { accessKey: tab.accessKey, uiaName: tab.uiaName, ribbonItems: [], note: 'disabled' };
                continue;
            }

            log('▶', `Step 2: "${tab.name}" 리본 항목 수집`);
            try {
                const ribbonItems = await this._collectRibbonItems(tab);
                this.map.tabs[tab.name] = {
                    accessKey: tab.accessKey,
                    uiaName: tab.uiaName,
                    ribbonItems,
                };
                this.map.stats.totalTabs++;
                this.map.stats.totalRibbonItems += ribbonItems.length;
                log('ℹ', `  "${tab.name}": ${ribbonItems.length}개 항목`);
            } catch (e) {
                log('⚠', `  "${tab.name}" 실패: ${e.message}`);
                this.map.stats.errors++;
                await this._recover();
            }
        }

        // Step 3: 고정 ToolBar / StatusBar 수집
        log('▶', 'Step 3: ToolBar + StatusBar 수집');
        await this._collectFixedBars();

        // Step 4: 다이얼로그 프로빙
        if (this.probeDialogs) {
            log('▶', 'Step 4: 다이얼로그 프로빙');
            await this._probeAllDialogs();
        }

        // Step 5: 드롭다운 프로빙 (plan §6.4)
        if (this.probeDropdowns) {
            log('▶', 'Step 5: 드롭다운 프로빙');
            // Step 4 직후 HWP 포그라운드가 불안정할 수 있음 — 복구 단계
            try { await controller.setForeground(); } catch (_) {}
            await sleep(DELAY.long);
            await this._ensureNoDialog();
            await this._ensureUiHealthy();
            await this._probeAllDropdowns();
        }

        // Step 6: Selection 상태 프로빙 — 도형 탭 (plan-selection Phase A)
        if (this.probeShapeTab) {
            log('▶', 'Step 6: 도형 탭 프로빙 (Selection 상태)');
            try { await controller.setForeground(); } catch (_) {}
            await sleep(DELAY.long);
            await this._ensureNoDialog();
            await this._ensureUiHealthy();
            await this._probeShapeTab();
        }

        log('▶', '완료!');
        log('ℹ', JSON.stringify(this.map.stats));
        return this.map;
    }

    save(outputPath) {
        const filePath = outputPath || path.join('maps', `${this.product}-menu-map.json`);
        try { fs.mkdirSync(path.dirname(filePath), { recursive: true }); } catch (_) {}
        fs.writeFileSync(filePath, JSON.stringify(this.map, null, 2), 'utf8');
        log('ℹ', `저장: ${filePath}`);
        return filePath;
    }

    // =========================================================================
    // Step 1: 메뉴 탭 수집 (UIA 직접 접근)
    // =========================================================================

    async _collectMenuTabs() {
        const session = sessionModule.getSession();
        sessionModule.refreshHwpElement();
        if (!session.hwpElement) throw new Error('한컴 윈도우를 찾을 수 없습니다');

        // MenuBarImpl에서 MainMenuItemImpl 자식을 직접 읽기
        const menuBar = session.hwpElement.findAllChildren()
            .find(c => { try { return c.className === 'MenuBarImpl'; } catch (_) { return false; } });
        if (!menuBar) throw new Error('MenuBarImpl을 찾을 수 없습니다');

        const menuChildren = menuBar.findAllChildren();
        const tabs = [];

        for (const el of menuChildren) {
            try {
                if (el.className !== 'MainMenuItemImpl' || !el.name) { el.release(); continue; }

                const match = el.name.match(/^(.+?)\s+([A-Z0-9]+)$/);
                const rect = el.boundingRect;
                const w = rect.right - rect.left;
                tabs.push({
                    name: match ? match[1] : el.name,
                    accessKey: match ? match[2] : null,
                    uiaName: el.name,
                    isEnabled: el.isEnabled,
                    // 왼쪽 1/3 지점 클릭 — 오른쪽의 드롭다운 화살표 영역 회피
                    clickX: rect.left + Math.round(w * 0.3),
                    clickY: Math.round((rect.top + rect.bottom) / 2),
                });
            } catch (_) {}
            el.release();
        }
        menuBar.release();

        return tabs;
    }

    // =========================================================================
    // Step 2: 탭별 리본 항목 수집 (마우스 클릭으로 탭 전환)
    // =========================================================================

    async _collectRibbonItems(tab) {
        // 마우스 클릭으로 탭 전환 (버튼 영역 클릭 → 리본 패널 교체)
        win32.mouseClick(tab.clickX, tab.clickY);
        await sleep(DELAY.long);

        // ToolBox 자식을 UIA로 직접 읽기
        const session = sessionModule.getSession();
        sessionModule.refreshHwpElement();
        if (!session.hwpElement) return [];

        const topChildren = session.hwpElement.findAllChildren();
        const toolbox = topChildren.find(c => {
            try { return c.className === 'ToolBoxImpl'; } catch (_) { return false; }
        });
        topChildren.filter(c => c !== toolbox).forEach(c => { try { c.release(); } catch (_) {} });

        if (!toolbox) return [];

        const tbChildren = toolbox.findAllChildren();
        const items = [];

        for (const child of tbChildren) {
            try {
                const name = child.name;
                if (!name || name === 'ToolBoxGallery' || name.includes('스크롤')) continue;

                const r = child.boundingRect;
                items.push({
                    name,
                    controlType: child.controlTypeName,
                    hasDropdown: child.controlTypeName === 'MenuItem',
                    type: 'unknown',
                    clickX: Math.round((r.left + r.right) / 2),
                    clickY: r.top + Math.round((r.bottom - r.top) * 0.3),
                });
            } catch (_) {}
            child.release();
        }
        toolbox.release();

        return items;
    }

    /**
     * 마우스 클릭으로 탭을 전환하는 헬퍼 (다이얼로그 프로빙 등에서 재사용)
     */
    async _switchTab(tab) {
        await this._ensureForeground();
        win32.mouseClick(tab.clickX, tab.clickY);
        await sleep(DELAY.medium);
        sessionModule.refreshHwpElement();
    }

    // =========================================================================
    // Step 3: 고정 바 수집
    // =========================================================================

    async _collectFixedBars() {
        const tree = await controller.getUiTree({ depth: 2 });

        // ToolBar
        const toolbar = tree.tree[0].children.find(c => c.className === 'ToolBarImpl');
        if (toolbar) {
            this.map.toolbar = toolbar.children
                .filter(c => c.name)
                .map(c => ({ name: c.name, controlType: c.controlType }));
            log('ℹ', `  ToolBar: ${this.map.toolbar.length}개`);
        }

        // StatusBar
        const statusbar = tree.tree[0].children.find(c => c.className === 'StatusBarImpl');
        if (statusbar) {
            this.map.statusbar = statusbar.children
                .filter(c => c.name)
                .map(c => ({ name: c.name, controlType: c.controlType }));
            log('ℹ', `  StatusBar: ${this.map.statusbar.length}개`);
        }
    }

    // =========================================================================
    // Step 4: 다이얼로그 프로빙 (각 항목 직접 invoke)
    // =========================================================================

    async _probeAllDialogs() {
        // 탭 좌표 재수집 (클릭 위치가 필요)
        const menuTabs = await this._collectMenuTabs();
        const tabLookup = {};
        for (const t of menuTabs) tabLookup[t.name] = t;

        // 프로빙 내내 재사용할 기준 텍스트를 문서에 한 번 삽입
        await this._seedBaselineContent();

        for (const [tabName, tabData] of Object.entries(this.map.tabs)) {
            if (tabData.note) continue; // backstage/disabled

            // 다이얼로그 없음이 확실하고 위험 액션이 많은 탭은 전체 스킵
            if (SKIP_PROBING_TABS.has(tabName)) {
                log('ℹ', `  "${tabName}" 다이얼로그 프로빙 스킵 (설정 다이얼로그 없음)`);
                for (const item of tabData.ribbonItems) {
                    if (item.hasDropdown) item.type = 'dropdown';
                    else item.type = 'action';
                }
                continue;
            }

            log('▶', `  "${tabName}" 다이얼로그 프로빙`);
            const tabInfo = tabLookup[tabName];
            if (!tabInfo) { log('⚠', `  "${tabName}" 탭 좌표 없음 — 스킵`); continue; }

            // 탭 시작 전 UI 건강 상태 체크 & 복구
            await this._ensureUiHealthy();

            for (const item of tabData.ribbonItems) {
                if (item.hasDropdown) { item.type = 'dropdown'; continue; }

                // 위험 액션(창 생성/분할/리본 토글)은 스킵
                if (this._isDangerousItem(item.name)) {
                    log('🚫', `    "${item.name}" 위험 액션 — 스킵`);
                    item.type = 'skipped';
                    continue;
                }

                try {
                    await this._switchTab(tabInfo);
                    // 매 버튼 프로빙 전에 캐럿 컨텍스트를 "문서 맨 앞 + 전체 선택"으로 통일
                    await this._resetCaretContext();

                    const dialog = await this._probeItem(item);
                    if (dialog) {
                        item.type = 'dialog';
                        item.dialog = dialog;
                        this.map.stats.totalDialogs++;
                        const ctrlCount = Object.values(dialog.controls).flat().length;
                        this.map.stats.totalControls += ctrlCount;
                        log('💬', `    "${item.name}" → ${dialog.title} (${dialog.tabs.length}탭, ${ctrlCount}컨트롤)`);
                    } else {
                        item.type = 'action';
                    }
                } catch (e) {
                    log('⚠', `    "${item.name}" 프로빙 실패: ${e.message}`);
                    item.type = 'error';
                    this.map.stats.errors++;
                    await this._recover();
                }

                await this._ensureNoDialog();
                // 주기적 건강 체크 — Ctrl+F1 의도치 않은 입력 등으로 리본이 사라졌는지 확인
                await this._ensureUiHealthy();
            }
        }
    }

    /**
     * 항목 이름이 DANGEROUS_ITEM_PATTERNS 중 하나와 일치하면 true
     */
    _isDangerousItem(itemName) {
        if (!itemName) return false;
        return DANGEROUS_ITEM_PATTERNS.some(pat => pat.test(itemName));
    }

    /**
     * 프로빙 전용 기준 텍스트를 문서에 삽입.
     * 이후 모든 _resetCaretContext()는 이 텍스트를 전체 선택 상태로 복원.
     * 이미 내용이 있으면 덮어쓰지 않도록 문서 맨 끝에 한 번만 추가.
     */
    async _seedBaselineContent() {
        // 문서 맨 끝으로 이동 후 테스트 텍스트 한 줄 추가
        await controller.pressKeys({ keys: 'Ctrl+End' });
        await sleep(DELAY.short);
        win32.clipboardPaste('\n가나다 ABC 1234');
        await sleep(DELAY.medium);
    }

    /**
     * 매 버튼 프로빙 전에 호출.
     * 캐럿을 문서 처음으로 보내고 전체 선택하여 "선택된 텍스트 있음"
     * 상태를 일관되게 유지.  선택 유무에 따라 동작이 갈리는 서식/편집
     * 관련 버튼도 동일한 컨텍스트에서 평가됨.
     */
    async _resetCaretContext() {
        await this._safeKeys('Escape');
        await sleep(DELAY.short);
        await this._safeKeys('Ctrl+Home');
        await sleep(DELAY.short);
        await this._safeKeys('Ctrl+A');
        await sleep(DELAY.short);
    }

    /**
     * UI 상태가 정상인지 확인하고 오염됐으면 복구.
     *   - ToolBoxImpl 없음 → Ctrl+F1 (리본 복구)
     *   - 윈도우 제목에 "(2)" 등 보조 창 표시 → Ctrl+F4 (보조 창 닫기)
     */
    async _ensureUiHealthy() {
        const session = sessionModule.getSession();
        sessionModule.refreshHwpElement();
        if (!session.hwpElement) return;

        // 보조 창 감지: 제목에 " (숫자)" 패턴
        const title = session.hwpElement.name || '';
        if (/\(\d+\)/.test(title)) {
            log('🔧', `  보조 창 감지 "${title}" — Ctrl+F4로 닫기`);
            await this._safeKeys('Ctrl+F4');
            await sleep(DELAY.long);
            sessionModule.refreshHwpElement();
        }

        // 리본 감지: ToolBoxImpl 존재 여부
        const children = session.hwpElement.findAllChildren();
        const hasToolbox = children.some(c => {
            try { return c.className === 'ToolBoxImpl'; } catch (_) { return false; }
        });
        children.forEach(c => { try { c.release(); } catch (_) {} });

        if (!hasToolbox) {
            log('🔧', '  리본 최소화 감지 — Ctrl+F1로 복구');
            await this._safeKeys('Ctrl+F1');
            await sleep(DELAY.long);
        }
    }

    /**
     * 단일 리본 버튼을 마우스 클릭으로 프로빙.
     * DialogImpl 감지 시 구조 수집 후 취소 버튼으로 닫는다.
     * 다이얼로그가 안 뜨면 즉시 실행된 것으로 간주하고 Ctrl+Z로 롤백.
     * @returns {object|null} 다이얼로그 정보 or null (즉시 실행)
     */
    async _probeItem(item) {
        if (typeof item.clickX !== 'number') return null;

        win32.mouseClick(item.clickX, item.clickY);
        await sleep(DELAY.dialog);

        const dialog = this._findDialog();
        if (!dialog) {
            // 다이얼로그 없음 → 즉시 실행으로 간주, 롤백
            await controller.pressKeys({ keys: 'Ctrl+Z' });
            await sleep(DELAY.short);
            return null;
        }

        // 다이얼로그 구조 수집
        const info = this._crawlDialog(dialog);
        dialog.release();

        // 닫기 (취소 버튼 우선, 없으면 ESC)
        await this._closeDialog();

        return info;
    }

    /**
     * FrameWindowImpl의 직계 자식에서 DialogImpl 찾기
     */
    _findDialog() {
        const session = sessionModule.getSession();
        sessionModule.refreshHwpElement();
        if (!session.hwpElement) return null;

        const topCh = session.hwpElement.findAllChildren();
        let dialog = null;
        for (const c of topCh) {
            try { if (c.className === 'DialogImpl') { dialog = c; break; } } catch (_) {}
        }
        topCh.filter(c => c !== dialog).forEach(c => { try { c.release(); } catch (_) {} });
        return dialog;
    }

    /**
     * 다이얼로그 구조를 UIA 트리에서 한 번에 수집.
     * 탭이 있으면 각 탭을 클릭해 전환한 뒤 컨트롤 목록 수집.
     */
    _crawlDialog(dialog) {
        const title = dialog.name || '';
        const result = { title, tabs: [], controls: {} };

        // TabItem 목록과 좌표 수집
        const descs = dialog.findAll(TreeScope.Descendants);
        const tabItems = [];
        for (const el of descs) {
            try {
                if (el.controlTypeName === 'TabItem') {
                    tabItems.push({ name: el.name, rect: el.boundingRect });
                }
            } catch (_) {}
        }
        descs.forEach(el => { try { el.release(); } catch (_) {} });

        if (tabItems.length === 0) {
            // 탭 없는 다이얼로그
            result.tabs = ['(기본)'];
            result.controls['(기본)'] = this._collectControlsFromDialog();
        } else {
            result.tabs = tabItems.map(t => t.name);
            for (const tab of tabItems) {
                const cx = Math.round((tab.rect.left + tab.rect.right) / 2);
                const cy = Math.round((tab.rect.top + tab.rect.bottom) / 2);
                win32.mouseClick(cx, cy);
                // 탭 전환 대기 (동기 함수지만 딜레이는 불가, 수집만)
                const start = Date.now();
                while (Date.now() - start < DELAY.medium) { /* spin */ }
                result.controls[tab.name] = this._collectControlsFromDialog();
            }
        }
        return result;
    }

    /**
     * 현재 열린 다이얼로그의 컨트롤을 타입별로 수집 (다이얼로그 버튼 제외)
     */
    _collectControlsFromDialog() {
        const dialog = this._findDialog();
        if (!dialog) return [];
        const descs = dialog.findAll(TreeScope.Descendants);
        const controls = [];
        for (const el of descs) {
            try {
                const t = el.controlTypeName;
                const n = el.name || '';
                // 다이얼로그 기본 버튼(확인/취소/도움말)과 TabItem 자체는 제외
                if (el.className === 'DialogButtonImpl') continue;
                if (t === 'TabItem') continue;
                if (t === 'Button' || t === 'CheckBox' || t === 'Edit' ||
                    t === 'ComboBox' || t === 'RadioButton') {
                    controls.push({ name: n, type: t, className: el.className });
                }
            } catch (_) {}
        }
        descs.forEach(el => { try { el.release(); } catch (_) {} });
        dialog.release();
        return controls;
    }

    /**
     * "취소"(DialogButtonImpl) 클릭으로 다이얼로그 닫기.
     * 없거나 실패하면 ESC로 폴백.
     */
    async _closeDialog() {
        const dialog = this._findDialog();
        if (!dialog) return;

        const descs = dialog.findAll(TreeScope.Descendants);
        let cancelRect = null;
        for (const el of descs) {
            try {
                if (el.className === 'DialogButtonImpl' && el.name === '취소') {
                    cancelRect = el.boundingRect;
                    break;
                }
            } catch (_) {}
        }
        descs.forEach(el => { try { el.release(); } catch (_) {} });
        dialog.release();

        if (cancelRect) {
            const cx = Math.round((cancelRect.left + cancelRect.right) / 2);
            const cy = Math.round((cancelRect.top + cancelRect.bottom) / 2);
            win32.mouseClick(cx, cy);
            await sleep(DELAY.medium);
        } else {
            await controller.pressKeys({ keys: 'Escape' });
            await sleep(DELAY.medium);
        }
    }

    /**
     * 다이얼로그가 아직 열려 있으면 ESC로 연속 닫기 시도.
     */
    async _ensureNoDialog() {
        for (let i = 0; i < 3; i++) {
            const d = this._findDialog();
            if (!d) return;
            d.release();
            await controller.pressKeys({ keys: 'Escape' });
            await sleep(DELAY.short);
        }
    }

    // =========================================================================
    // Foreground 가드 — 이벤트가 잘못된 앱으로 가지 않도록 방어
    // =========================================================================

    /**
     * 현재 foreground 창의 소유 PID가 HWP와 일치하는지 확인.
     * 불일치 시 setForeground 재시도. 2회 실패 시 throw하여 잘못된 이벤트 발사 차단.
     */
    async _ensureForeground() {
        const expectedPid = sessionModule.getSession().hwpProcess.pid;
        if (!expectedPid) return;
        const fg = win32.getForegroundPid();
        if (fg === expectedPid) return;
        log('🔆', `  foreground=${fg} ≠ HWP=${expectedPid} — 복원 중`);
        try { await controller.setForeground(); } catch (_) {}
        await sleep(DELAY.medium);
        const fg2 = win32.getForegroundPid();
        if (fg2 !== expectedPid) {
            // 한 번 더 시도
            try { await controller.setForeground(); } catch (_) {}
            await sleep(DELAY.long);
            const fg3 = win32.getForegroundPid();
            if (fg3 !== expectedPid) {
                throw new Error(`HWP foreground 복원 실패: fg=${fg3}, expected=${expectedPid}`);
            }
        }
    }

    async _safeClick(x, y) {
        await this._ensureForeground();
        win32.mouseClick(x, y);
    }

    async _safeKeys(keys) {
        await this._ensureForeground();
        return controller.pressKeys({ keys });
    }

    // =========================================================================
    // Step 5: 드롭다운 프로빙 (plan §6)
    // =========================================================================

    async _probeAllDropdowns() {
        // Step 5 진입 시 _collectMenuTabs가 hwpElement 없어서 실패할 수 있으므로 재시도
        let menuTabs;
        for (let retry = 0; retry < 3; retry++) {
            try {
                menuTabs = await this._collectMenuTabs();
                break;
            } catch (e) {
                log('⚠', `  _collectMenuTabs 실패 (${retry+1}/3): ${e.message} — 재시도`);
                try { await controller.setForeground(); } catch (_) {}
                await sleep(DELAY.long);
                sessionModule.refreshHwpElement();
            }
        }
        if (!menuTabs) {
            log('⚠', '  Step 5 진입 실패 — 드롭다운 프로빙 중단');
            return;
        }
        const tabLookup = {};
        for (const t of menuTabs) tabLookup[t.name] = t;

        const totalDeadline = Date.now() + DD_LIMITS.totalProbingBudgetMs;

        for (const [tabName, tabData] of Object.entries(this.map.tabs)) {
            if (Date.now() > totalDeadline) {
                log('⏰', '  전체 드롭다운 프로빙 예산 초과 — 중단');
                break;
            }
            if (tabData.note) continue; // backstage/disabled
            if (SKIP_PROBING_TABS.has(tabName)) continue;

            const tabInfo = tabLookup[tabName];
            if (!tabInfo) { log('⚠', `  "${tabName}" 탭 좌표 없음 — 드롭다운 스킵`); continue; }

            const dropdownItems = (tabData.ribbonItems || []).filter(i => i.hasDropdown);
            if (dropdownItems.length === 0) continue;

            log('▶', `  "${tabName}" 드롭다운 프로빙 (${dropdownItems.length}개)`);
            const tabDeadline = Date.now() + DD_LIMITS.perTabBudgetMs;

            await this._ensureUiHealthy();

            for (const item of dropdownItems) {
                if (Date.now() > tabDeadline) {
                    log('⏰', `  "${tabName}" 탭 예산 초과 — 남은 드롭다운 스킵`);
                    break;
                }
                if (this._isDangerousDropdown(item.name)) {
                    log('🚫', `    "${item.name}" 드롭다운 블랙리스트 — 스킵`);
                    item.dropdown = { classification: 'blacklisted', itemCount: 0, items: [], notes: ['blacklisted'] };
                    continue;
                }

                try {
                    await this._switchTab(tabInfo);
                    await this._resetCaretContext();
                    const dd = await this._probeDropdown(item);
                    item.dropdown = dd;
                    if (dd.classification !== 'no-popup' && dd.classification !== 'error') {
                        this.map.stats.totalDropdowns++;
                        this.map.stats.totalDropdownItems += dd.itemCount;
                    }
                    log('📂', `    "${item.name}" → ${dd.classification} (${dd.itemCount}개)`);
                } catch (e) {
                    log('⚠', `    "${item.name}" 드롭다운 실패: ${e.message}`);
                    item.dropdown = {
                        classification: 'error',
                        itemCount: 0,
                        items: [],
                        notes: [`error:${e.message}`],
                        probeDurationMs: 0,
                    };
                    this.map.stats.dropdownErrors++;
                    // Layer 5: 재귀 복구 차단
                    if (!this._recoveryInProgress) {
                        this._recoveryInProgress = true;
                        try { await this._recover(); }
                        finally { this._recoveryInProgress = false; }
                    }
                }

                await this._ensureNoPopup();
                await this._ensureUiHealthy();
            }
        }
    }

    /**
     * 단일 드롭다운 프로빙.
     * plan §6.4 — 우하단 1/4 클릭 + 팝업 snapshot diff + 트리 walk + Escape close.
     */
    async _probeDropdown(item) {
        const start = Date.now();
        const deadline = start + DD_LIMITS.perDropdownBudgetMs;
        const check = () => { if (Date.now() > deadline) throw new Error('BudgetExceeded'); };

        const result = {
            classification: 'no-popup',
            popupClass: null,
            itemCount: 0,
            items: [],
            probeDurationMs: 0,
            notes: [],
        };

        // Rect 재수집 (pre-sweep 좌표가 stale 할 수 있음)
        check();
        const fresh = this._refreshItem(item.name);
        let clickX, clickY;
        if (fresh) {
            const r = fresh.rect;
            clickX = Math.round(r.left + (r.right - r.left) * 0.8);
            clickY = Math.round(r.top + (r.bottom - r.top) * 0.8);
            try { fresh.element.release(); } catch (_) {}
        } else {
            clickX = item.clickX;
            clickY = item.clickY;
            result.notes.push('rect-refresh-failed');
        }

        // Before snapshot
        check();
        sessionModule.refreshHwpElement();
        const session = sessionModule.getSession();
        if (!session.hwpElement) return { ...result, notes: [...result.notes, 'no-hwp-element'] };
        const before = this._snapshotChildren(session.hwpElement);
        this._releaseAll(before.children);

        // Click — 한컴 ExpandCollapse 패턴은 no-op 확인됨, 마우스 클릭만 사용
        check();
        await this._safeClick(clickX, clickY);
        await sleep(DD_LIMITS.popupWaitMs);

        // After snapshot
        check();
        sessionModule.refreshHwpElement();
        const frameAfter = sessionModule.getSession().hwpElement;
        const after = this._snapshotChildren(frameAfter);

        // Diff로 새 Popup 감지
        const added = this._findAddedChildren(before, after);
        let popupEl = null;
        let popupInfo = null;
        for (const a of added) {
            if (a.info.className === 'Popup' && a.info.controlType === 'Window') {
                popupEl = a.element;
                popupInfo = a.info;
                break;
            }
        }

        if (popupEl) {
            check();
            result.popupClass = popupInfo.className;
            try {
                const visited = new Set();
                const nodes = [];
                this._walkPopupTree(popupEl, 0, visited, nodes);
                if (nodes.length >= DD_LIMITS.popupTreeMaxNodes) result.notes.push('tree-truncated');
                result.classification = this._classifyDropdown(nodes);
                result.itemCount = nodes.length;
                result.items = this._summarizeItems(nodes);
            } catch (e) {
                result.notes.push(`walk-error:${e.message}`);
                result.classification = 'error';
            }
        } else {
            result.notes.push('no-popup-detected');
        }

        // 모든 after 자식 해제 (popupEl 포함)
        this._releaseAll(after.children);

        // Close popup (Layer 1 + 7)
        for (let i = 0; i < DD_LIMITS.closeAttempts; i++) {
            check();
            await this._safeKeys('Escape');
            await sleep(DELAY.short);
        }

        result.probeDurationMs = Date.now() - start;
        return result;
    }

    /**
     * 수집된 node 리스트를 바탕으로 드롭다운 유형 결정.
     * Discovery 결과 기반 — plan §3 갱신.
     */
    _classifyDropdown(nodes) {
        if (nodes.length <= 1) return 'empty';
        const d1 = nodes.filter(n => n.depth === 1);
        const hasScroll = d1.some(n => n.className === 'ScrollImpl');
        const hasGrid = d1.some(n => n.className === 'RowColImpl');
        const hasList = d1.some(n => n.controlType === 'List');
        const hasMenuItem = d1.some(n => n.controlType === 'MenuItem');

        if (hasScroll) return 'scrollable-gallery';
        if (hasGrid) return 'visual-grid';
        if (hasList && hasMenuItem) return 'gallery-mixed';
        if (hasMenuItem) return 'menu';
        if (hasList) return 'gallery-pure';
        return 'unknown';
    }

    _summarizeItems(nodes) {
        return nodes
            .filter(n => n.name && (n.controlType === 'MenuItem' || n.controlType === 'ListItem'))
            .map(n => ({
                name: n.name,
                controlType: n.controlType,
                depth: n.depth,
                isDialogOpener: DIALOG_OPENER_RE.test(n.name),
            }));
    }

    _snapshotChildren(parent) {
        const children = parent.findAllChildren();
        const infos = children.map(c => {
            const info = { className: '', controlType: '', name: '', rect: null };
            try { info.className = c.className || ''; } catch (_) {}
            try { info.controlType = c.controlTypeName || ''; } catch (_) {}
            try { info.name = c.name || ''; } catch (_) {}
            try { info.rect = c.boundingRect; } catch (_) {}
            return info;
        });
        return { children, infos };
    }

    _findAddedChildren(before, after) {
        const fp = (info) => {
            const r = info.rect;
            const rs = r ? `${r.left},${r.top},${r.right},${r.bottom}` : '';
            return `${info.className}|${info.controlType}|${info.name}|${rs}`;
        };
        const beforeSet = new Set(before.infos.map(fp));
        const added = [];
        for (let i = 0; i < after.infos.length; i++) {
            if (!beforeSet.has(fp(after.infos[i]))) {
                added.push({ info: after.infos[i], element: after.children[i] });
            }
        }
        return added;
    }

    /**
     * 팝업 트리 재귀 순회. Layer 3: visited + depth + count.
     * visited 키는 (className, name, rect) 튜플 — koffi pointer String 변환 불가 회피.
     */
    _walkPopupTree(el, depth, visited, nodes) {
        if (depth > DD_LIMITS.popupTreeMaxDepth) return;
        if (nodes.length >= DD_LIMITS.popupTreeMaxNodes) throw new Error('TreeSizeExceeded');

        const node = { depth, name: '', controlType: '', className: '', rect: null };
        try { node.name = el.name || ''; } catch (_) {}
        try { node.controlType = el.controlTypeName || ''; } catch (_) {}
        try { node.className = el.className || ''; } catch (_) {}
        try { node.rect = el.boundingRect; } catch (_) {}

        const r = node.rect;
        const key = `${node.className}|${node.name}|${r ? `${r.left},${r.top},${r.right},${r.bottom}` : ''}`;
        if (visited.has(key)) return;
        visited.add(key);
        nodes.push(node);

        if (depth < DD_LIMITS.popupTreeMaxDepth) {
            const children = el.findAllChildren();
            try {
                for (const child of children) {
                    if (nodes.length >= DD_LIMITS.popupTreeMaxNodes) break;
                    this._walkPopupTree(child, depth + 1, visited, nodes);
                }
            } finally {
                this._releaseAll(children);
            }
        }
    }

    _releaseAll(list) {
        for (const c of list) { try { c.release(); } catch (_) {} }
    }

    /**
     * 현재 활성 탭에서 리본 항목을 이름으로 찾아 살아있는 element + rect 반환.
     * 호출자는 사용 후 element.release() 필요.
     */
    _refreshItem(itemName) {
        const session = sessionModule.getSession();
        sessionModule.refreshHwpElement();
        if (!session.hwpElement) return null;

        const topChildren = session.hwpElement.findAllChildren();
        const toolbox = topChildren.find(c => {
            try { return c.className === 'ToolBoxImpl'; } catch (_) { return false; }
        });
        topChildren.filter(c => c !== toolbox).forEach(c => { try { c.release(); } catch (_) {} });
        if (!toolbox) return null;

        let hit = null;
        const tbChildren = toolbox.findAllChildren();
        for (const c of tbChildren) {
            try {
                if (c.name === itemName) {
                    hit = { element: c, rect: c.boundingRect };
                    break;
                }
            } catch (_) {}
        }
        tbChildren.filter(c => !hit || c !== hit.element).forEach(c => { try { c.release(); } catch (_) {} });
        toolbox.release();
        return hit;
    }

    /**
     * FrameWindowImpl 자식 중 Popup(Window)가 여러 개 있으면 Escape로 정리.
     * 한컴은 baseline으로 내부 Popup(32×32)을 항상 가지므로 count>1 일 때만 정리.
     */
    async _ensureNoPopup() {
        for (let i = 0; i < 3; i++) {
            sessionModule.refreshHwpElement();
            const session = sessionModule.getSession();
            if (!session.hwpElement) return;
            const children = session.hwpElement.findAllChildren();
            const popupCount = children.filter(c => {
                try { return c.className === 'Popup' && c.controlTypeName === 'Window'; }
                catch (_) { return false; }
            }).length;
            children.forEach(c => { try { c.release(); } catch (_) {} });
            if (popupCount <= 1) return;
            await this._safeKeys('Escape');
            await sleep(DELAY.short);
        }
    }

    _isDangerousDropdown(name) {
        if (!name) return false;
        return DANGEROUS_DROPDOWN_PATTERNS.some(pat => pat.test(name));
    }

    // =========================================================================
    // Step 6: 도형 탭 프로빙 (Selection-dependent 컨텍스트 탭)
    // =========================================================================

    /**
     * 편집 탭의 "도형" 드롭다운을 열고 "직사각형"을 선택한 뒤 캔버스를 단일 클릭
     * 하여 기본 크기 도형을 삽입. 삽입 후 자동 선택됨.
     * @returns {{clickX:number, clickY:number}|null} 삽입 위치 (재선택용) 또는 실패 시 null
     */
    async _insertShapeForProbing() {
        await this._safeKeys('Ctrl+End');
        await sleep(DELAY.short);

        // 편집 탭으로 전환
        const menuTabs = await this._collectMenuTabs();
        const editTab = menuTabs.find(t => t.name === '편집');
        if (!editTab) { log('⚠', '  편집 탭 없음'); return null; }
        await this._switchTab(editTab);
        await sleep(DELAY.long);

        // 도형 드롭다운 버튼 rect 재수집 (pre-built 좌표가 stale 할 수 있음)
        const shapeBtn = this._refreshItem('도형 : ALT+H');
        if (!shapeBtn) { log('⚠', '  "도형" 드롭다운 버튼 없음'); return null; }
        const sr = shapeBtn.rect;
        const bx = Math.round(sr.left + (sr.right - sr.left) * 0.8);
        const by = Math.round(sr.top + (sr.bottom - sr.top) * 0.8);
        try { shapeBtn.element.release(); } catch (_) {}

        // 드롭다운 열기
        await this._safeClick(bx, by);
        await sleep(DD_LIMITS.popupWaitMs);

        // "직사각형" 아이템 찾기
        sessionModule.refreshHwpElement();
        const session = sessionModule.getSession();
        if (!session.hwpElement) { log('⚠', '  HWP element 없음'); return null; }

        const target = this._findInLatestPopup(session.hwpElement, '직사각형');
        if (!target) {
            await this._safeKeys('Escape');
            await sleep(DELAY.short);
            log('⚠', '  드롭다운에서 "직사각형" 찾지 못함');
            return null;
        }

        const rcx = Math.round((target.rect.left + target.rect.right) / 2);
        const rcy = Math.round((target.rect.top + target.rect.bottom) / 2);
        await this._safeClick(rcx, rcy);
        await sleep(DELAY.long);

        // 캔버스(HwpMainEditWnd) 중앙 찾기
        sessionModule.refreshHwpElement();
        const frame = sessionModule.getSession().hwpElement;
        const children = frame.findAllChildren();
        let canvas = null;
        for (const c of children) {
            try { if (c.className === 'HwpMainEditWnd') { canvas = c; break; } } catch (_) {}
        }
        children.filter(c => c !== canvas).forEach(c => { try { c.release(); } catch (_) {} });
        if (!canvas) { log('⚠', '  캔버스 없음'); return null; }

        const cr = canvas.boundingRect;
        const canvasCx = Math.round((cr.left + cr.right) / 2);
        const canvasCy = Math.round((cr.top + cr.bottom) / 2);
        try { canvas.release(); } catch (_) {}

        // 단일 클릭으로 기본 크기 도형 삽입 (HWP 기본 동작)
        await this._safeClick(canvasCx, canvasCy);
        await sleep(DELAY.dialog);

        return { clickX: canvasCx, clickY: canvasCy };
    }

    /**
     * 가장 큰 Popup 자식 내부를 descendants 탐색하여 name으로 요소 검색.
     * (baseline Popup은 32×32로 작으므로 크기로 구분)
     * @returns {{rect: object}|null}
     */
    _findInLatestPopup(frame, targetName) {
        const topChildren = frame.findAllChildren();
        let largest = null;
        let largestArea = 0;
        for (const c of topChildren) {
            try {
                if (c.className === 'Popup' && c.controlTypeName === 'Window') {
                    const r = c.boundingRect;
                    const area = (r.right - r.left) * (r.bottom - r.top);
                    if (area > largestArea) {
                        if (largest) largest.release();
                        largest = c;
                        largestArea = area;
                        continue;
                    }
                }
            } catch (_) {}
            try { c.release(); } catch (_) {}
        }
        if (!largest) return null;

        const descs = largest.findAll(TreeScope.Descendants);
        let hit = null;
        for (const el of descs) {
            try {
                const name = (el.name || '').trim();
                if (name === targetName) {
                    hit = { rect: el.boundingRect };
                    break;
                }
            } catch (_) {}
        }
        descs.forEach(el => { try { el.release(); } catch (_) {} });
        largest.release();
        return hit;
    }

    /**
     * 삽입한 도형을 Escape + Ctrl+Z×4로 롤백. 실패해도 조용히 종료.
     */
    async _cleanupShape() {
        try {
            await this._safeKeys('Escape');
            await sleep(DELAY.short);
            for (let i = 0; i < 4; i++) {
                await this._safeKeys('Ctrl+Z');
                await sleep(DELAY.short);
            }
        } catch (e) {
            log('⚠', `  도형 정리 실패 (무시): ${e.message}`);
        }
    }

    /**
     * 도형 탭(컨텍스트 탭) 프로빙.
     * 도형 삽입 → 탭 활성화 확인 → 리본 수집 → 드롭다운 프로빙 → 정리.
     */
    async _probeShapeTab() {
        await this._ensureUiHealthy();

        const shapeLoc = await this._insertShapeForProbing();
        if (!shapeLoc) {
            log('⚠', '  도형 삽입 실패 — Step 6 스킵');
            await this._cleanupShape();
            return;
        }

        try {
            const menuTabs = await this._collectMenuTabs();
            const shapeTab = menuTabs.find(t => t.name === '도형');
            if (!shapeTab || !shapeTab.isEnabled) {
                log('⚠', '  도형 탭 비활성 — 프로빙 스킵');
                return;
            }

            // 도형 탭으로 전환 + 리본 수집
            await this._switchTab(shapeTab);
            await sleep(DELAY.long);
            const ribbonItems = await this._collectRibbonItems(shapeTab);
            for (const it of ribbonItems) it.type = it.hasDropdown ? 'dropdown' : 'action';

            // map 업데이트 (기존 'disabled' note 덮어쓰기)
            this.map.tabs['도형'] = {
                accessKey: shapeTab.accessKey,
                uiaName: shapeTab.uiaName,
                ribbonItems,
                contextState: 'shape-selected',
            };
            this.map.stats.totalTabs++;
            this.map.stats.totalRibbonItems += ribbonItems.length;
            log('ℹ', `  도형 탭: ${ribbonItems.length}개 리본 항목 수집`);

            // 드롭다운 프로빙
            const dropdownItems = ribbonItems.filter(i => i.hasDropdown);
            if (dropdownItems.length === 0) return;
            log('▶', `  도형 탭 드롭다운 프로빙 (${dropdownItems.length}개)`);
            const tabDeadline = Date.now() + DD_LIMITS.perTabBudgetMs;

            for (const item of dropdownItems) {
                if (Date.now() > tabDeadline) { log('⏰', '  도형 탭 예산 초과'); break; }
                if (this._isDangerousDropdown(item.name)) {
                    log('🚫', `    "${item.name}" 블랙리스트`);
                    item.dropdown = { classification: 'blacklisted', itemCount: 0, items: [], notes: ['blacklisted'] };
                    continue;
                }

                // 도형 선택이 유지되는지 확인 — 풀리면 재선택
                if (!(await this._isShapeTabActive())) {
                    log('🔁', '    도형 선택 풀림 — 재선택');
                    await this._safeClick(shapeLoc.clickX, shapeLoc.clickY);
                    await sleep(DELAY.long);
                    if (!(await this._isShapeTabActive())) {
                        log('⚠', `    "${item.name}" 도형 재선택 실패 — 스킵`);
                        item.dropdown = { classification: 'shape-deselected', itemCount: 0, items: [], notes: ['shape-deselected'] };
                        continue;
                    }
                    await this._switchTab(shapeTab);
                    await sleep(DELAY.long);
                }

                try {
                    const dd = await this._probeDropdown(item);
                    item.dropdown = dd;
                    if (dd.classification !== 'no-popup' && dd.classification !== 'error') {
                        this.map.stats.totalDropdowns++;
                        this.map.stats.totalDropdownItems += dd.itemCount;
                    }
                    log('📂', `    "${item.name}" → ${dd.classification} (${dd.itemCount}개)`);
                } catch (e) {
                    log('⚠', `    "${item.name}" 실패: ${e.message}`);
                    item.dropdown = { classification: 'error', itemCount: 0, items: [], notes: [`error:${e.message}`] };
                    this.map.stats.dropdownErrors++;
                    if (!this._recoveryInProgress) {
                        this._recoveryInProgress = true;
                        try { await this._recover(); } finally { this._recoveryInProgress = false; }
                    }
                }

                await this._ensureNoPopup();
            }
        } finally {
            await this._cleanupShape();
        }
    }

    /**
     * 현재 도형 탭이 활성 상태인지 빠르게 확인.
     */
    async _isShapeTabActive() {
        try {
            const menuTabs = await this._collectMenuTabs();
            const shapeTab = menuTabs.find(t => t.name === '도형');
            return !!(shapeTab && shapeTab.isEnabled);
        } catch (_) { return false; }
    }

    // =========================================================================
    // 유틸리티
    // =========================================================================

    async _recover() {
        // 포커스 먼저 복원 — Escape가 잘못된 앱으로 가지 않도록
        try { await controller.setForeground(); } catch (_) {}
        await sleep(DELAY.medium);
        for (let i = 0; i < 5; i++) {
            try { await this._safeKeys('Escape'); }
            catch (_) { break; } // foreground 회복 불가 시 조용히 종료
            await sleep(DELAY.short);
        }
        await sleep(DELAY.long);
    }

    _isDialogButton(name) {
        if (!name) return false;
        const l = name.toLowerCase();
        return ['확인', '취소', '닫기', 'ok', 'cancel', 'close', '적용', '도움말'].some(k => l.includes(k));
    }
}

// =============================================================================
// CLI
// =============================================================================

async function main() {
    const product = process.argv[2] || 'hwp';
    const noDialogs = process.argv.includes('--no-dialogs');

    const mapper = new MenuMapper({ product, probeDialogs: !noDialogs });
    await mapper.run();
    mapper.save();
}

if (require.main === module) {
    main().catch(e => { console.error('Error:', e.message); process.exit(1); });
}

module.exports = { MenuMapper };
