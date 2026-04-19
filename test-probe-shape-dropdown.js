'use strict';
/**
 * 편집 탭의 도형 드롭다운을 열고 각 아이템의 UIA 속성을 덤프.
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
    await mapper._switchTab(editTab);
    await new Promise(r => setTimeout(r, 500));

    const items = await mapper._collectRibbonItems(editTab);
    const shapeItem = items.find(i => i.name && /^도형/.test(i.name) && i.hasDropdown);
    if (!shapeItem) {
        process.stderr.write('도형 드롭다운 미탐지. 편집 탭 항목:\n');
        items.forEach(i => process.stderr.write(`  - "${i.name}"\n`));
        throw new Error('no shape');
    }
    process.stderr.write(`도형 드롭다운: "${shapeItem.name}" @(${shapeItem.clickX},${shapeItem.clickY})\n`);

    // 드롭다운 열기
    const r = mapper._refreshItem(shapeItem.name);
    if (!r) throw new Error('_refreshItem 실패');
    const rect = r.rect;
    const rW = rect.right - rect.left, rH = rect.bottom - rect.top;
    const bx = rW > rH ? rect.right - 8 : Math.round(rect.left + rW * 0.8);
    const by = rW > rH ? Math.round(rect.top + rH * 0.5) : Math.round(rect.top + rH * 0.8);
    try { r.element.release(); } catch (_) {}
    await mapper._safeClick(bx, by);
    await new Promise(r => setTimeout(r, 600));

    // popup 찾기
    sessionModule.refreshHwpElement();
    const topChildren = sessionModule.getSession().hwpElement.findAllChildren();
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
    process.stderr.write(`popup rect: ${JSON.stringify(popup.boundingRect)}\n\n`);

    const descs = popup.findAll(TreeScope.Descendants);
    const rows = [];
    for (const d of descs) {
        try {
            const ct = d.controlTypeName || '';
            if (ct !== 'MenuItem' && ct !== 'ListItem' && ct !== 'Button') continue;
            rows.push({
                ct,
                name: (d.name || '').slice(0, 50),
                automationId: d.automationId || '',
                className: d.className || '',
                acceleratorKey: d.acceleratorKey || '',
            });
        } catch (_) {}
    }
    descs.forEach(d => { try { d.release(); } catch (_) {} });
    try { popup.release(); } catch (_) {}

    await mapper._safeKeys('Escape');
    await new Promise(r => setTimeout(r, 200));
    await mapper._safeKeys('Escape');

    process.stderr.write(`총 ${rows.length}개:\n`);
    for (const row of rows) {
        process.stderr.write(`  [${row.ct}] name="${row.name}" automationId="${row.automationId}" cls="${row.className}" acc="${row.acceleratorKey}"\n`);
    }
    console.log(JSON.stringify(rows, null, 2));
}

run().catch(e => { process.stderr.write(`FATAL: ${e.message}\n`); process.exit(1); });
