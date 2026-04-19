'use strict';
/**
 * 스타일 드롭다운을 열고 각 아이템의 UIA 속성을 모두 덤프.
 * 짧은 이름("바탕글", "본문", "개요 1")이 어느 속성에 들어있는지 확인.
 */
const { MenuMapper } = require('./src/menu-mapper');
const controller = require('./src/hwp-controller');
const sessionModule = require('./src/session');
const { TreeScope } = require('./src/uia');

async function run() {
    await controller.attach({ product: 'hwp' });
    await controller.setForeground();
    for (let i = 0; i < 3; i++) {
        await controller.pressKeys({ keys: 'Escape' });
        await new Promise(r => setTimeout(r, 200));
    }

    const mapper = new MenuMapper({ product: 'hwp' });
    mapper.map.mappedAt = new Date().toISOString();

    await controller.pressKeys({ keys: 'Escape' });
    await new Promise(r => setTimeout(r, 300));

    const menuTabs = await mapper._collectMenuTabs();
    const editTab = menuTabs.find(t => t.name === '편집');
    if (!editTab) throw new Error('편집 탭 없음');

    // 편집 탭으로 전환
    await mapper._switchTab(editTab);
    await new Promise(r => setTimeout(r, 500));

    const items = await mapper._collectRibbonItems(editTab);
    const styleItem = items.find(i => i.name && /스타일/.test(i.name) && i.hasDropdown);
    if (!styleItem) {
        process.stderr.write('스타일 드롭다운 항목 없음. 전체 ribbon:\n');
        items.forEach(i => process.stderr.write(`  - "${i.name}" hasDropdown=${i.hasDropdown}\n`));
        throw new Error('스타일 드롭다운 없음');
    }
    process.stderr.write(`발견: "${styleItem.name}" @(${styleItem.clickX},${styleItem.clickY})\n`);

    // 드롭다운 열기 (_probeDropdown의 클릭 로직 재사용)
    const r = mapper._refreshItem(styleItem.name);
    if (!r) throw new Error('_refreshItem 실패');
    const rect = r.rect;
    const rW = rect.right - rect.left, rH = rect.bottom - rect.top;
    const bx = rW > rH ? rect.right - 8 : Math.round(rect.left + rW * 0.8);
    const by = rW > rH ? Math.round(rect.top + rH * 0.5) : Math.round(rect.top + rH * 0.8);
    try { r.element.release(); } catch (_) {}
    process.stderr.write(`드롭다운 열기: (${bx},${by}), rect=${rect.left},${rect.top},${rect.right},${rect.bottom}\n`);

    await mapper._safeClick(bx, by);
    await new Promise(r => setTimeout(r, 600));

    // 가장 큰 Popup 찾기
    sessionModule.refreshHwpElement();
    const frame = sessionModule.getSession().hwpElement;
    const topChildren = frame.findAllChildren();
    let popup = null, largest = 0;
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
    if (!popup) throw new Error('popup 미탐지');
    process.stderr.write(`popup rect: ${JSON.stringify(popup.boundingRect)}, area=${largest}\n`);

    // 모든 descendant를 덤프. MenuItem/ListItem 필터링.
    const descs = popup.findAll(TreeScope.Descendants);
    process.stderr.write(`popup descendants: ${descs.length}개\n\n`);

    const rows = [];
    for (const d of descs) {
        try {
            const name = d.name || '';
            const ct = d.controlTypeName || '';
            if (ct !== 'MenuItem' && ct !== 'ListItem' && ct !== 'Button') continue;
            rows.push({
                ct,
                name: name.replace(/\n/g, ' ¶ ').slice(0, 70),
                automationId: d.automationId || '',
                className: d.className || '',
                acceleratorKey: d.acceleratorKey || '',
                localizedControlType: d.localizedControlType || '',
            });
        } catch (_) {}
    }
    descs.forEach(d => { try { d.release(); } catch (_) {} });
    try { popup.release(); } catch (_) {}

    // Escape to close
    await mapper._safeKeys('Escape');
    await new Promise(r => setTimeout(r, 200));
    await mapper._safeKeys('Escape');

    // 덤프
    process.stderr.write(`총 ${rows.length}개 아이템\n\n`);
    for (const row of rows) {
        process.stderr.write(`[${row.ct}]\n`);
        process.stderr.write(`  name:             "${row.name}"\n`);
        process.stderr.write(`  automationId:     "${row.automationId}"\n`);
        process.stderr.write(`  className:        "${row.className}"\n`);
        process.stderr.write(`  acceleratorKey:   "${row.acceleratorKey}"\n`);
        process.stderr.write(`  localizedCtrlTy:  "${row.localizedControlType}"\n`);
        process.stderr.write('\n');
    }

    console.log(JSON.stringify(rows, null, 2));
}

run().catch(e => {
    process.stderr.write(`FATAL: ${e.message}\n${e.stack}\n`);
    process.exit(1);
});
