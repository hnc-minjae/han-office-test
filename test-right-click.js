'use strict';
/**
 * Phase E smoke test: 우클릭 컨텍스트 메뉴 감지.
 * 선택된 텍스트 위에서 우클릭 → 팝업 감지 + 트리 수집 + Escape.
 */
const controller = require('./src/hwp-controller');
const sessionModule = require('./src/session');
const win32 = require('./src/win32');
const { MenuMapper } = require('./src/menu-mapper');
const { TreeScope } = require('./src/uia');

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function snapshot(el) {
    const children = el.findAllChildren();
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

function fp(i) { const r = i.rect; return `${i.className}|${i.controlType}|${i.name}|${r ? `${r.left},${r.top},${r.right},${r.bottom}` : ''}`; }

async function run() {
    await controller.attach({ product: 'hwp' });
    await controller.setForeground();
    await sleep(500);
    for (let i = 0; i < 3; i++) { await controller.pressKeys({ keys: 'Escape' }); await sleep(200); }

    // 텍스트 삽입
    await controller.pressKeys({ keys: 'Ctrl+End' });
    await sleep(200);
    win32.clipboardPaste('\n테스트 텍스트 가나다 ABC');
    await sleep(400);
    await controller.pressKeys({ keys: 'Ctrl+Home' });
    await sleep(200);
    await controller.pressKeys({ keys: 'Ctrl+A' });
    await sleep(400);

    // 캔버스 중앙 좌표
    sessionModule.refreshHwpElement();
    const session = sessionModule.getSession();
    const children = session.hwpElement.findAllChildren();
    let canvas = null;
    for (const c of children) {
        try { if (c.className === 'HwpMainEditWnd') { canvas = c; break; } } catch (_) {}
    }
    children.filter(c => c !== canvas).forEach(c => { try { c.release(); } catch (_) {} });
    const cr = canvas.boundingRect;
    const cx = Math.round((cr.left + cr.right) / 2);
    const cy = Math.round((cr.top + cr.bottom) / 2);
    try { canvas.release(); } catch (_) {}

    process.stderr.write(`▶ 위치 (${cx}, ${cy})\n`);

    // 먼저 좌클릭으로 커서 위치 확정
    await controller.setForeground();
    await sleep(200);
    win32.mouseClick(cx, cy);
    await sleep(300);

    // Before snapshot
    sessionModule.refreshHwpElement();
    const frame = sessionModule.getSession().hwpElement;
    const before = snapshot(frame);
    before.children.forEach(c => { try { c.release(); } catch (_) {} });
    const uia = sessionModule.getUia();
    const rootBefore = uia.getRootElement();
    const rootBeforeSnap = snapshot(rootBefore);
    rootBeforeSnap.children.forEach(c => { try { c.release(); } catch (_) {} });
    try { rootBefore.release(); } catch (_) {}

    // Right-click
    await controller.setForeground();
    await sleep(200);
    process.stderr.write('▶ 우클릭 시작\n');
    win32.mouseClick(cx, cy, { rightClick: true });
    await sleep(2000);

    // After snapshot
    sessionModule.refreshHwpElement();
    const after = snapshot(sessionModule.getSession().hwpElement);
    const rootAfter = uia.getRootElement();
    const rootAfterSnap = snapshot(rootAfter);

    const beforeSet = new Set(before.infos.map(fp));
    const added = [];
    for (let i = 0; i < after.infos.length; i++) {
        if (!beforeSet.has(fp(after.infos[i]))) added.push({ info: after.infos[i], el: after.children[i], src: 'frame' });
    }
    const rootBeforeSet = new Set(rootBeforeSnap.infos.map(fp));
    for (let i = 0; i < rootAfterSnap.infos.length; i++) {
        if (!rootBeforeSet.has(fp(rootAfterSnap.infos[i]))) added.push({ info: rootAfterSnap.infos[i], el: rootAfterSnap.children[i], src: 'desktop' });
    }
    process.stderr.write(`새 요소: ${added.length}개 (frame+desktop)\n`);
    for (const a of added) {
        process.stderr.write(`  [${a.src}] ${a.info.className} | ${a.info.controlType} | "${a.info.name}" | ${JSON.stringify(a.info.rect)}\n`);
    }

    // Popup 내부 walk
    const popup = added.find(a => a.info.className === 'Popup' && a.info.controlType === 'Window');
    if (popup) {
        process.stderr.write('▶ Popup 내부 descendants:\n');
        const descs = popup.el.findAll(TreeScope.Descendants);
        let shown = 0;
        for (const d of descs) {
            try {
                if (shown < 40) {
                    process.stderr.write(`  | ${d.className} | ${d.controlTypeName} | "${d.name || ''}"\n`);
                    shown++;
                }
            } catch (_) {}
        }
        process.stderr.write(`  (total descendants: ${descs.length})\n`);
        descs.forEach(d => { try { d.release(); } catch (_) {} });
    } else {
        process.stderr.write('❌ Popup 미감지\n');
    }

    after.children.forEach(c => { try { c.release(); } catch (_) {} });
    rootAfterSnap.children.forEach(c => { try { c.release(); } catch (_) {} });
    try { rootAfter.release(); } catch (_) {}

    // Close
    for (let i = 0; i < 3; i++) { await controller.pressKeys({ keys: 'Escape' }); await sleep(200); }
    await controller.pressKeys({ keys: 'Ctrl+Z' });
}

run().catch((e) => { process.stderr.write(`FATAL: ${e.message}\n${e.stack}\n`); process.exit(1); });
