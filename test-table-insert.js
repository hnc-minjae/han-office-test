'use strict';
/**
 * Phase B-1 smoke test: 표 삽입 + 새 컨텍스트 탭 감지 + 정리.
 */
const controller = require('./src/hwp-controller');
const sessionModule = require('./src/session');
const win32 = require('./src/win32');
const { MenuMapper } = require('./src/menu-mapper');
const { TreeScope } = require('./src/uia');

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function findInPopup(frame, targetName) {
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
    descs.forEach((el) => { try { el.release(); } catch (_) {} });
    largest.release();
    return hit;
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

    const mapper = new MenuMapper({ product: 'hwp', probeDialogs: false, probeDropdowns: false });
    const initialTabs = await mapper._collectMenuTabs();
    const initialTabNames = new Set(initialTabs.filter(t => t.isEnabled).map(t => t.name));
    process.stderr.write(`초기 활성 탭: ${Array.from(initialTabNames).join(', ')}\n`);

    // 입력 탭으로 전환
    const inputTab = initialTabs.find(t => t.name === '입력');
    if (!inputTab) throw new Error('입력 탭 없음');
    await mapper._switchTab(inputTab);
    await sleep(500);

    // 표 드롭다운 열기
    const tableBtn = mapper._refreshItem('표 : ALT+T');
    if (!tableBtn) throw new Error('입력/표 버튼 없음');
    const tr = tableBtn.rect;
    const bx = Math.round(tr.left + (tr.right - tr.left) * 0.8);
    const by = Math.round(tr.top + (tr.bottom - tr.top) * 0.8);
    try { tableBtn.element.release(); } catch (_) {}

    process.stderr.write(`▶ 표 드롭다운 열기 (${bx}, ${by})\n`);
    await controller.setForeground();
    await sleep(200);
    win32.mouseClick(bx, by);
    await sleep(1500);

    sessionModule.refreshHwpElement();
    const session = sessionModule.getSession();
    const target = await findInPopup(session.hwpElement, '표 만들기 : ALT+T');
    if (!target) {
        await controller.pressKeys({ keys: 'Escape' });
        throw new Error('"표 만들기" 메뉴 없음');
    }
    process.stderr.write(`▶ "표 만들기" 발견: rect=${JSON.stringify(target.rect)}\n`);

    const rcx = Math.round((target.rect.left + target.rect.right) / 2);
    const rcy = Math.round((target.rect.top + target.rect.bottom) / 2);
    await controller.setForeground();
    await sleep(200);
    win32.mouseClick(rcx, rcy);
    await sleep(1200);

    // 다이얼로그 열림 예상 - Enter로 기본값 삽입
    process.stderr.write('▶ Enter로 표 만들기 다이얼로그 확인\n');
    await controller.pressKeys({ keys: 'Enter' });
    await sleep(1500);

    // 새 탭 감지
    const afterTabs = await mapper._collectMenuTabs();
    const afterEnabled = new Set(afterTabs.filter(t => t.isEnabled).map(t => t.name));
    const newTabs = Array.from(afterEnabled).filter(n => !initialTabNames.has(n));
    process.stderr.write(`표 삽입 후 활성 탭: ${Array.from(afterEnabled).join(', ')}\n`);
    process.stderr.write(`새로 등장한 탭: ${newTabs.length > 0 ? newTabs.join(', ') : '없음'}\n`);

    if (newTabs.length > 0) {
        for (const tabName of newTabs) {
            const tabInfo = afterTabs.find(t => t.name === tabName);
            if (!tabInfo) continue;
            await mapper._switchTab(tabInfo);
            await sleep(500);
            const items = await mapper._collectRibbonItems(tabInfo);
            process.stderr.write(`  "${tabName}" 리본: ${items.length}개\n`);
            for (const it of items.slice(0, 6)) {
                process.stderr.write(`    - ${it.name} (${it.controlType})\n`);
            }
        }
    }

    // 정리
    process.stderr.write('▶ 정리 (Escape + Ctrl+Z × 6)\n');
    await controller.setForeground();
    await controller.pressKeys({ keys: 'Escape' });
    await sleep(300);
    for (let i = 0; i < 6; i++) {
        await controller.pressKeys({ keys: 'Ctrl+Z' });
        await sleep(300);
    }

    const finalTabs = await mapper._collectMenuTabs();
    const finalEnabled = new Set(finalTabs.filter(t => t.isEnabled).map(t => t.name));
    const remaining = Array.from(finalEnabled).filter(n => !initialTabNames.has(n));
    process.stderr.write(`정리 후 남은 새 탭: ${remaining.length > 0 ? remaining.join(', ') : '없음 ✅'}\n`);

    if (remaining.length > 0) process.exit(1);
}

run().catch((e) => {
    process.stderr.write(`FATAL: ${e.message}\n${e.stack}\n`);
    process.exit(1);
});
