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

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function log(icon, msg) { process.stderr.write(`${icon} ${msg}\n`); }

// =============================================================================
// MenuMapper
// =============================================================================

class MenuMapper {
    constructor(options = {}) {
        this.product = options.product || 'hwp';
        this.probeDialogs = options.probeDialogs !== false;

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
            stats: { totalTabs: 0, totalRibbonItems: 0, totalDialogs: 0, totalControls: 0, errors: 0 },
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
        await controller.pressKeys({ keys: 'Escape' });
        await sleep(DELAY.short);
        await controller.pressKeys({ keys: 'Ctrl+Home' });
        await sleep(DELAY.short);
        await controller.pressKeys({ keys: 'Ctrl+A' });
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
            await controller.pressKeys({ keys: 'Ctrl+F4' });
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
            await controller.pressKeys({ keys: 'Ctrl+F1' });
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
    // 유틸리티
    // =========================================================================

    async _recover() {
        for (let i = 0; i < 5; i++) {
            await controller.pressKeys({ keys: 'Escape' });
            await sleep(DELAY.short);
        }
        await sleep(DELAY.long);
        try { await controller.setForeground(); } catch (_) {}
        await sleep(DELAY.medium);
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
