'use strict';
/**
 * Phase C 진단: 도형 "채우기" 드롭다운 클릭 후 실제로 어디에 팝업이 나타나는지 조사.
 * 현재 _probeDropdown은 FrameWindowImpl 직계 자식 중 className='Popup'만 감지.
 * 이 팝업이 다른 className이거나 다른 위치에 있을 수 있음.
 */
const controller = require('./src/hwp-controller');
const sessionModule = require('./src/session');
const win32 = require('./src/win32');
const { MenuMapper } = require('./src/menu-mapper');
const { TreeScope } = require('./src/uia');

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function walk(el, depth, maxDepth, seen, out) {
    if (depth > maxDepth) return;
    if (out.length >= 300) return;
    const node = { depth, className: '', controlType: '', name: '', rect: null };
    try { node.className = el.className || ''; } catch (_) {}
    try { node.controlType = el.controlTypeName || ''; } catch (_) {}
    try { node.name = el.name || ''; } catch (_) {}
    try { node.rect = el.boundingRect; } catch (_) {}
    const r = node.rect;
    const key = `${node.className}|${node.name}|${r ? `${r.left},${r.top},${r.right},${r.bottom}` : ''}`;
    if (seen.has(key)) return;
    seen.add(key);
    out.push(node);
    if (depth >= maxDepth) return;
    const kids = el.findAllChildren();
    try {
        for (const c of kids) walk(c, depth + 1, maxDepth, seen, out);
    } finally {
        for (const c of kids) { try { c.release(); } catch (_) {} }
    }
}

function snapshotRoot(uia, depthFromRoot = 2) {
    const root = uia.getRootElement();
    const out = [];
    const seen = new Set();
    walk(root, 0, depthFromRoot, seen, out);
    try { root.release(); } catch (_) {}
    return out;
}

function snapshotFrame(frame, depth = 2) {
    const out = [];
    const seen = new Set();
    walk(frame, 0, depth, seen, out);
    return out;
}

function fingerprint(node) {
    const r = node.rect;
    return `${node.className}|${node.controlType}|${node.name}|${r ? `${r.left},${r.top},${r.right},${r.bottom}` : ''}|d${node.depth}`;
}

function diff(before, after) {
    const bSet = new Set(before.map(fingerprint));
    return after.filter(n => !bSet.has(fingerprint(n)));
}

async function run() {
    await controller.attach({ product: 'hwp' });
    await controller.setForeground();
    await sleep(500);
    for (let i = 0; i < 3; i++) {
        await controller.pressKeys({ keys: 'Escape' });
        await sleep(200);
    }
    await controller.pressKeys({ keys: 'Ctrl+End' });
    await sleep(300);

    const mapper = new MenuMapper({ product: 'hwp', probeDialogs: false, probeDropdowns: false, probeShapeTab: false, probeTableTabs: false });

    // 도형 삽입
    process.stderr.write('▶ 도형 삽입\n');
    const shapeLoc = await mapper._insertShapeForProbing();
    if (!shapeLoc) throw new Error('도형 삽입 실패');

    const menuTabs = await mapper._collectMenuTabs();
    const shapeTab = menuTabs.find(t => t.name === '도형');
    if (!shapeTab || !shapeTab.isEnabled) throw new Error('도형 탭 비활성');
    await mapper._switchTab(shapeTab);
    await sleep(800);

    // "도형 채우기" 드롭다운 rect 재수집
    const fillBtn = mapper._refreshItem('도형 채우기 : ALT+F');
    if (!fillBtn) throw new Error('"도형 채우기" 버튼 없음');
    const fr = fillBtn.rect;
    const bx = Math.round(fr.left + (fr.right - fr.left) * 0.8);
    const by = Math.round(fr.top + (fr.bottom - fr.top) * 0.8);
    try { fillBtn.element.release(); } catch (_) {}
    process.stderr.write(`▶ "도형 채우기" 버튼 rect=${JSON.stringify(fr)}, click=(${bx},${by})\n`);

    // BEFORE snapshot — Desktop root depth 2, Frame children depth 2
    sessionModule.refreshHwpElement();
    const session = sessionModule.getSession();
    const uia = sessionModule.getUia();

    const beforeRoot = snapshotRoot(uia, 2);
    const beforeFrame = snapshotFrame(session.hwpElement, 2);
    process.stderr.write(`BEFORE: desktop=${beforeRoot.length}, frame=${beforeFrame.length}\n`);

    // Click
    await controller.setForeground();
    await sleep(200);
    win32.mouseClick(bx, by);
    await sleep(1500);

    // AFTER snapshot
    sessionModule.refreshHwpElement();
    const afterRoot = snapshotRoot(uia, 2);
    const afterFrame = snapshotFrame(sessionModule.getSession().hwpElement, 2);
    process.stderr.write(`AFTER: desktop=${afterRoot.length}, frame=${afterFrame.length}\n`);

    // Diff
    const newRoot = diff(beforeRoot, afterRoot);
    const newFrame = diff(beforeFrame, afterFrame);
    process.stderr.write(`\n=== 새로 등장한 요소 ===\n`);
    process.stderr.write(`[Desktop diff] ${newRoot.length}개:\n`);
    for (const n of newRoot) {
        process.stderr.write(`  d=${n.depth} | ${n.className} | ${n.controlType} | "${n.name}" | ${JSON.stringify(n.rect)}\n`);
    }
    process.stderr.write(`[Frame diff] ${newFrame.length}개:\n`);
    for (const n of newFrame) {
        process.stderr.write(`  d=${n.depth} | ${n.className} | ${n.controlType} | "${n.name}" | ${JSON.stringify(n.rect)}\n`);
    }

    // 정리
    await controller.setForeground();
    for (let i = 0; i < 3; i++) { await controller.pressKeys({ keys: 'Escape' }); await sleep(200); }
    await mapper._cleanupShape();
}

run().catch((e) => {
    process.stderr.write(`FATAL: ${e.message}\n${e.stack}\n`);
    process.exit(1);
});
