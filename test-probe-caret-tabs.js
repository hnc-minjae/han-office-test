'use strict';
/**
 * Step 9.5/9.6/9.7 단독 검증.
 *   - 메모 탭 (입력/메모 클릭 → "메모" contextual tab 기대)
 *   - 주석 탭 (Ctrl+N,N → "주석" contextual tab 기대)
 *   - 머리말/꼬리말 탭 (쪽/머리말 → "머리말/꼬리말" contextual tab 기대)
 *
 * 전체 맵 생성(~15분)을 건너뛰고, 의존하는 '입력'/'쪽' 탭 리본만 수집한 뒤
 * 3개 probe를 순차 실행.
 */
const { MenuMapper } = require('./src/menu-mapper');
const controller = require('./src/hwp-controller');

async function run() {
    await controller.attach({ product: 'hwp' });
    await controller.setForeground();
    for (let i = 0; i < 3; i++) {
        await controller.pressKeys({ keys: 'Escape' });
        await new Promise(r => setTimeout(r, 200));
    }

    const mapper = new MenuMapper({
        product: 'hwp',
        probeDialogs: false,
        probeDropdowns: false,
        probeShapeTab: false,
        probeTableTabs: false,
        probeChartTabs: false,
        probeImageTab: false,
        probeContextMenus: false,
        probeFileBackstage: false,
    });
    mapper.map.mappedAt = new Date().toISOString();

    await controller.pressKeys({ keys: 'Escape' });
    await new Promise(r => setTimeout(r, 300));

    // Step 1: 탭 목록 수집
    const menuTabs = await mapper._collectMenuTabs();
    process.stderr.write(`\nℹ 탭 ${menuTabs.length}개:\n`);
    for (const t of menuTabs) {
        process.stderr.write(`  "${t.name}" accessKey=${t.accessKey} uiaName="${t.uiaName}" enabled=${t.isEnabled} @(${t.clickX},${t.clickY})\n`);
    }

    // Step 2: '입력', '쪽' 탭만 리본 수집 (나머지는 placeholder)
    for (const tab of menuTabs) {
        if (tab.name === '파일' || !tab.isEnabled) {
            mapper.map.tabs[tab.name] = { accessKey: tab.accessKey, uiaName: tab.uiaName, ribbonItems: [], note: tab.isEnabled ? 'backstage' : 'disabled' };
            continue;
        }
        if (tab.name !== '입력' && tab.name !== '쪽') {
            mapper.map.tabs[tab.name] = { accessKey: tab.accessKey, uiaName: tab.uiaName, ribbonItems: [], note: 'skipped-for-test' };
            continue;
        }
        const items = await mapper._collectRibbonItems(tab);
        for (const it of items) it.type = it.hasDropdown ? 'dropdown' : 'action';
        mapper.map.tabs[tab.name] = { accessKey: tab.accessKey, uiaName: tab.uiaName, ribbonItems: items };
        mapper.map.stats.totalTabs++;
        mapper.map.stats.totalRibbonItems += items.length;
        process.stderr.write(`ℹ "${tab.name}" 리본: ${items.length}개\n`);
    }

    // enterFn만 실행하고 탭 상태를 덤프하는 진단 헬퍼
    const sleep = (ms) => new Promise(r => setTimeout(r, ms));
    async function diagnose(label, enterFn) {
        process.stderr.write(`\n▶ ${label} 진단 시작\n`);
        await mapper._ensureUiHealthy();

        const before = await mapper._collectMenuTabs();
        process.stderr.write(`  before: ${before.map(t => `${t.name}${t.isEnabled ? '' : '*'}`).join(', ')}\n`);

        try { await enterFn(); }
        catch (e) { process.stderr.write(`  enterFn 예외: ${e.message}\n`); }
        await sleep(700);

        const after = await mapper._collectMenuTabs();
        process.stderr.write(`  after : ${after.map(t => `${t.name}${t.isEnabled ? '' : '*'}`).join(', ')}\n`);

        const beforeNames = new Set(before.map(t => t.name));
        const afterNames = new Set(after.map(t => t.name));
        const added = [...afterNames].filter(n => !beforeNames.has(n));
        const removed = [...beforeNames].filter(n => !afterNames.has(n));
        process.stderr.write(`  diff: +[${added.join(', ')}] -[${removed.join(', ')}]\n`);
        const enabledChanges = [];
        for (const a of after) {
            const b = before.find(x => x.name === a.name);
            if (b && b.isEnabled !== a.isEnabled) {
                enabledChanges.push(`${a.name}: ${b.isEnabled}→${a.isEnabled}`);
            }
        }
        if (enabledChanges.length) process.stderr.write(`  enabled 변화: ${enabledChanges.join(' | ')}\n`);

        // 복구
        try { await mapper._recover(); } catch (_) {}
        for (let i = 0; i < 4; i++) { try { await mapper._safeKeys('Ctrl+Z'); } catch (_) {} await sleep(100); }
        try { await mapper._ensureUiHealthy(); } catch (_) {}
    }

    // 각 probe의 enterFn만 실행하여 실제로 탭이 추가되는지 확인
    const memoItem = (mapper.map.tabs['입력']?.ribbonItems || []).find(i => i.name && i.name.includes('메모'));
    const tab쪽 = mapper.map.tabs['쪽'];
    const headerItem = (tab쪽?.ribbonItems || []).find(i => i.name && /머리말/.test(i.name));
    const controller2 = require('./src/hwp-controller');
    const win32 = require('./src/win32');

    // 입력 탭 리본 덤프
    process.stderr.write(`\n=== 입력 탭 ribbon 덤프 (메모/각주·미주 관련) ===\n`);
    for (const it of mapper.map.tabs['입력'].ribbonItems) {
        if (it.name && (/메모|각주|미주|주석/.test(it.name))) {
            process.stderr.write(`  "${it.name}" hasDropdown=${it.hasDropdown} ctrl=${it.controlType} @(${it.clickX},${it.clickY})\n`);
        }
    }
    process.stderr.write(`\n=== 쪽 탭 ribbon 덤프 (머리말/꼬리말) ===\n`);
    for (const it of mapper.map.tabs['쪽'].ribbonItems) {
        if (it.name && /머리말|꼬리말/.test(it.name)) {
            process.stderr.write(`  "${it.name}" hasDropdown=${it.hasDropdown} ctrl=${it.controlType} @(${it.clickX},${it.clickY})\n`);
        }
    }

    // 입력·쪽 탭의 실제 MenuTab 객체 (_switchTab용)
    const inputMenuTab = menuTabs.find(t => t.name === '입력');
    const sideMenuTab = menuTabs.find(t => t.name === '쪽');

    await diagnose('memo (mouse switch 입력 → 메모 click)', async () => {
        if (!memoItem) throw new Error('memoItem 없음');
        process.stderr.write(`  memoItem: "${memoItem.name}" hasDropdown=${memoItem.hasDropdown} ctrl=${memoItem.controlType}\n`);
        await mapper._switchTab(inputMenuTab);
        await sleep(500);
        await mapper._safeClick(memoItem.clickX, memoItem.clickY);
        process.stderr.write(`  메모 클릭 좌표: (${memoItem.clickX}, ${memoItem.clickY})\n`);
        await sleep(1200);
    });

    await diagnose('annotation (Ctrl+N, N)', async () => {
        await controller2.pressKeys({ keys: 'Ctrl+N' });
        await sleep(150);
        await controller2.pressKeys({ keys: 'N' });
    });

    // 검증된 _probeDropdown 로직으로 머리말 dropdown이 열리는지 먼저 확인
    process.stderr.write('\n▶ _probeDropdown 직접 호출로 머리말 dropdown 테스트\n');
    await mapper._switchTab(sideMenuTab);
    await sleep(500);
    try {
        const dd = await mapper._probeDropdown(headerItem);
        process.stderr.write(`  _probeDropdown 결과: classification=${dd.classification}, itemCount=${dd.itemCount}\n`);
        if (dd.items && dd.items.length) {
            for (const it of dd.items.slice(0, 10)) {
                process.stderr.write(`    - "${it.name}" (${it.controlType}, depth=${it.depth})\n`);
            }
        }
    } catch (e) { process.stderr.write(`  _probeDropdown 예외: ${e.message}\n`); }
    try { await mapper._recover(); } catch (_) {}
    await sleep(500);

    // Strategy SAFE: dropdown 열고 여러 키조합 테스트
    for (const trial of [
        { label: 'A 바로',                     keys: ['A'] },
        { label: 'Alt+A',                      keys: ['Alt+A'] },
        { label: 'A Enter',                    keys: ['A', 'Enter'] },
        { label: 'A A',                        keys: ['A', 'A'] },
        { label: 'Right Enter',                keys: ['Right', 'Enter'] },
    ]) {
        await diagnose(`headerFooter dropdown→ ${trial.label}`, async () => {
            await mapper._switchTab(sideMenuTab);
            await sleep(500);
            const r = mapper._refreshItem(headerItem.name);
            if (!r) return;
            const rect = r.rect;
            const rW = rect.right - rect.left, rH = rect.bottom - rect.top;
            const bx = rW > rH ? rect.right - 8 : Math.round(rect.left + rW * 0.8);
            const by = rW > rH ? Math.round(rect.top + rH * 0.5) : Math.round(rect.top + rH * 0.8);
            try { r.element.release(); } catch (_) {}
            await mapper._safeClick(bx, by);
            await sleep(700);
            for (const k of trial.keys) {
                await mapper._safeKeys(k);
                await sleep(400);
            }
            await sleep(1000);
        });
    }

    // Ctrl+N 단축키 계열 시도 (각주는 Ctrl+N,N; 머리말은 Ctrl+N,H 추정)
    await diagnose('headerFooter Ctrl+N, H', async () => {
        await mapper._safeKeys('Ctrl+N');
        await sleep(250);
        await mapper._safeKeys('H');
        await sleep(1200);
    });

    await diagnose('headerFooter Ctrl+N, K', async () => {
        await mapper._safeKeys('Ctrl+N');
        await sleep(250);
        await mapper._safeKeys('K');
        await sleep(1200);
    });

    // 머리말/꼬리말 item (dropdown 3rd item) — 개체 만들기 대화상자로 진입할 가능성
    await diagnose('headerFooter popup→ 머리말/꼬리말 item invoke', async () => {
        const sessionModule = require('./src/session');
        const { TreeScope } = require('./src/uia');
        await mapper._switchTab(sideMenuTab);
        await sleep(500);
        const r = mapper._refreshItem(headerItem.name);
        if (!r) return;
        const rect = r.rect;
        const bx = rect.right - 8;
        const by = Math.round(rect.top + (rect.bottom - rect.top) * 0.8);
        try { r.element.release(); } catch (_) {}
        await mapper._safeClick(bx, by);
        await sleep(700);

        sessionModule.refreshHwpElement();
        const topChildren = sessionModule.getSession().hwpElement.findAllChildren();
        let popup = null;
        for (const c of topChildren) {
            try { if (c.className === 'Popup' && !popup) { popup = c; continue; } } catch (_) {}
            try { c.release(); } catch (_) {}
        }
        if (!popup) return;
        const descs = popup.findAll(TreeScope.Descendants);
        let target = null;
        for (const d of descs) {
            try { if (d.controlTypeName === 'MenuItem' && /머리말\/꼬리말/.test(d.name || '')) { target = d; break; } } catch (_) {}
        }
        if (target) {
            const ir = target.boundingRect;
            const cx = Math.round((ir.left + ir.right) / 2);
            const cy = Math.round((ir.top + ir.bottom) / 2);
            process.stderr.write(`  머리말/꼬리말 item @(${cx},${cy})\n`);
            await mapper._safeClick(cx, cy);
        }
        descs.forEach(d => { try { d.release(); } catch (_) {} });
        try { popup.release(); } catch (_) {}
        await sleep(1500);
    });

    // 위쪽 invoke 후 서브메뉴 덤프
    await diagnose('headerFooter popup→ UIA invoke 위쪽 + submenu dump', async () => {
        const sessionModule = require('./src/session');
        const { TreeScope } = require('./src/uia');
        await mapper._switchTab(sideMenuTab);
        await sleep(500);
        const r = mapper._refreshItem(headerItem.name);
        if (!r) return;
        const rect = r.rect;
        const bx = rect.right - 8;
        const by = Math.round(rect.top + (rect.bottom - rect.top) * 0.8);
        try { r.element.release(); } catch (_) {}
        await mapper._safeClick(bx, by);
        await sleep(700);

        sessionModule.refreshHwpElement();
        let topChildren = sessionModule.getSession().hwpElement.findAllChildren();
        let popup1 = null;
        for (const c of topChildren) {
            try { if (c.className === 'Popup' && !popup1) { popup1 = c; continue; } } catch (_) {}
            try { c.release(); } catch (_) {}
        }
        if (!popup1) { process.stderr.write('  popup1 미탐지\n'); return; }

        const descs1 = popup1.findAll(TreeScope.Descendants);
        let target = null;
        for (const d of descs1) {
            try { if (d.controlTypeName === 'MenuItem' && /위쪽/.test(d.name || '')) { target = d; break; } } catch (_) {}
        }
        if (target) { try { target.invoke(); } catch (_) {} }
        descs1.forEach(d => { try { if (d !== target) d.release(); } catch (_) {} });
        try { popup1.release(); } catch (_) {}
        await sleep(1200); // 서브메뉴 렌더링 대기

        // 모든 Popup 자식 덤프
        sessionModule.refreshHwpElement();
        topChildren = sessionModule.getSession().hwpElement.findAllChildren();
        process.stderr.write(`  후속 Popups:\n`);
        for (const c of topChildren) {
            try {
                if (c.className === 'Popup') {
                    const pr = c.boundingRect;
                    const descs = c.findAll(TreeScope.Descendants);
                    const items = descs.map(d => { try { return `${d.controlTypeName}:${d.name || '-'}`; } catch (_) { return '?'; } })
                        .filter(s => s.startsWith('MenuItem:'));
                    process.stderr.write(`    rect=${JSON.stringify(pr)} items=${items.slice(0, 8).join(' | ')}\n`);
                    descs.forEach(d => { try { d.release(); } catch (_) {} });
                }
            } catch (_) {}
            try { c.release(); } catch (_) {}
        }
    });

    // UIA invoke() 패턴으로 위쪽 항목 실행
    await diagnose('headerFooter popup→ UIA invoke 위쪽', async () => {
        const sessionModule = require('./src/session');
        const { TreeScope } = require('./src/uia');
        await mapper._switchTab(sideMenuTab);
        await sleep(500);
        const r = mapper._refreshItem(headerItem.name);
        if (!r) return;
        const rect = r.rect;
        const bx = rect.right - 8;
        const by = Math.round(rect.top + (rect.bottom - rect.top) * 0.8);
        try { r.element.release(); } catch (_) {}
        await mapper._safeClick(bx, by);
        await sleep(700);

        sessionModule.refreshHwpElement();
        const frame = sessionModule.getSession().hwpElement;
        const topChildren = frame.findAllChildren();
        let popup = null, largest = 0;
        for (const c of topChildren) {
            try {
                if (c.className === 'Popup') {
                    const pr = c.boundingRect;
                    const a = (pr.right - pr.left) * (pr.bottom - pr.top);
                    if (a > largest) { if (popup) popup.release(); popup = c; largest = a; continue; }
                }
            } catch (_) {}
            try { c.release(); } catch (_) {}
        }
        if (!popup) return;

        const descs = popup.findAll(TreeScope.Descendants);
        let invoked = false;
        for (const d of descs) {
            try {
                if (d.controlTypeName === 'MenuItem' && d.name && /위쪽/.test(d.name)) {
                    process.stderr.write(`  invoke("${d.name}")\n`);
                    try { d.invoke(); invoked = true; process.stderr.write('  invoke OK\n'); }
                    catch (e) { process.stderr.write(`  invoke 실패: ${e.message}\n`); }
                    break;
                }
            } catch (_) {}
        }
        descs.forEach(d => { try { d.release(); } catch (_) {} });
        try { popup.release(); } catch (_) {}
        await sleep(1000);
        if (invoked) {
            // 서브메뉴가 열렸을 수 있음 — 첫 항목 Enter
            await mapper._safeKeys('Enter');
            await sleep(1200);
        }
    });

    // UIA로 popup에서 첫 MenuItem rect를 찾아 직접 클릭
    await diagnose('headerFooter popup→ UIA find & click first item', async () => {
        const sessionModule = require('./src/session');
        const { TreeScope } = require('./src/uia');
        await mapper._switchTab(sideMenuTab);
        await sleep(500);
        const r = mapper._refreshItem(headerItem.name);
        if (!r) return;
        const rect = r.rect;
        const rW = rect.right - rect.left, rH = rect.bottom - rect.top;
        const bx = rW > rH ? rect.right - 8 : Math.round(rect.left + rW * 0.8);
        const by = rW > rH ? Math.round(rect.top + rH * 0.5) : Math.round(rect.top + rH * 0.8);
        try { r.element.release(); } catch (_) {}
        await mapper._safeClick(bx, by);
        await sleep(700);

        // popup에서 첫 번째 MenuItem을 찾아서 rect 확인
        sessionModule.refreshHwpElement();
        const frame = sessionModule.getSession().hwpElement;
        const topChildren = frame.findAllChildren();
        let popup = null;
        let largest = 0;
        for (const c of topChildren) {
            try {
                if (c.className === 'Popup') {
                    const pr = c.boundingRect;
                    const area = (pr.right - pr.left) * (pr.bottom - pr.top);
                    if (area > largest) { if (popup) popup.release(); popup = c; largest = area; continue; }
                }
            } catch (_) {}
            try { c.release(); } catch (_) {}
        }
        if (!popup) { process.stderr.write('  popup 미탐지\n'); return; }
        process.stderr.write(`  popup rect: ${JSON.stringify(popup.boundingRect)}\n`);
        const descs = popup.findAll(TreeScope.Descendants);
        let clicked = false;
        for (const d of descs) {
            try {
                if (d.controlTypeName === 'MenuItem' && d.name && /위쪽/.test(d.name)) {
                    const ir = d.boundingRect;
                    const cx = Math.round((ir.left + ir.right) / 2);
                    const cy = Math.round((ir.top + ir.bottom) / 2);
                    process.stderr.write(`  "위쪽" item @(${cx},${cy}) rect=${JSON.stringify(ir)}\n`);
                    await mapper._safeClick(cx, cy);
                    clicked = true;
                    break;
                }
            } catch (_) {}
        }
        descs.forEach(d => { try { d.release(); } catch (_) {} });
        try { popup.release(); } catch (_) {}
        if (!clicked) { process.stderr.write('  위쪽 미탐지\n'); return; }
        await sleep(800);
        // 위쪽 서브메뉴가 열렸을 것 — 첫 항목 Enter
        await mapper._safeKeys('Enter');
        await sleep(1200);
    });

    const results = { diagnose: 'done' };

    // 요약
    process.stderr.write('\n=== 검증 결과 ===\n');
    process.stderr.write(`run 상태: ${JSON.stringify(results)}\n`);
    process.stderr.write(`stats: ${JSON.stringify(mapper.map.stats)}\n`);
    const knownTabs = new Set(menuTabs.filter(t => t.isEnabled).map(t => t.name));
    const newTabs = Object.keys(mapper.map.tabs).filter(n => !knownTabs.has(n) && mapper.map.tabs[n].contextState);
    process.stderr.write(`신규 contextual tabs: ${newTabs.length === 0 ? '(없음)' : newTabs.join(', ')}\n`);
    for (const n of newTabs) {
        const t = mapper.map.tabs[n];
        process.stderr.write(`  "${n}": contextState=${t.contextState}, ribbonItems=${t.ribbonItems?.length || 0}\n`);
        if (t.ribbonItems?.length) {
            const names = t.ribbonItems.map(i => i.name).filter(Boolean).slice(0, 10);
            process.stderr.write(`    샘플: ${names.join(' | ')}\n`);
        }
    }

    // JSON 덤프 (신규 탭만)
    const dump = {};
    for (const n of newTabs) dump[n] = mapper.map.tabs[n];
    console.log(JSON.stringify(dump, null, 2));
}

run().catch(e => {
    process.stderr.write(`FATAL: ${e.message}\n${e.stack}\n`);
    process.exit(1);
});
