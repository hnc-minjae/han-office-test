'use strict';
/**
 * Phase A-1 smoke test: 도형 삽입 + 도형 탭 활성화 확인 + 정리(Ctrl+Z).
 * 성공 기준:
 *   - 도형 삽입 후 "도형" 탭 isEnabled=true
 *   - Ctrl+Z 후 "도형" 탭 isEnabled=false로 복귀
 *   - 문서 상태 정상 (빈 문서)
 */

const controller = require('./src/hwp-controller');
const sessionModule = require('./src/session');
const win32 = require('./src/win32');
const { MenuMapper } = require('./src/menu-mapper');
const { TreeScope } = require('./src/uia');

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/**
 * 가장 최근에 생성된 Popup(가장 큰 것) 내부에서 targetName과 일치하는 요소를 찾음.
 * 이름 앞뒤 공백/trailing space 허용.
 */
async function findInLatestPopup(frame, targetName) {
    const topChildren = frame.findAllChildren();
    // 가장 큰 Popup 선택 (baseline popup은 32×32 등 작음)
    let largestPopup = null;
    let largestArea = 0;
    for (const c of topChildren) {
        try {
            if (c.className === 'Popup' && c.controlTypeName === 'Window') {
                const r = c.boundingRect;
                const area = (r.right - r.left) * (r.bottom - r.top);
                if (area > largestArea) {
                    if (largestPopup) largestPopup.release();
                    largestPopup = c;
                    largestArea = area;
                    continue;
                }
            }
        } catch (_) {}
        try { c.release(); } catch (_) {}
    }
    if (!largestPopup) return null;

    const descs = largestPopup.findAll(TreeScope.Descendants);
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
    largestPopup.release();
    return hit;
}

async function run() {
    process.stderr.write('▶ HWP attach\n');
    await controller.attach({ product: 'hwp' });
    await controller.setForeground();
    await sleep(500);
    for (let i = 0; i < 3; i++) {
        await controller.pressKeys({ keys: 'Escape' });
        await sleep(200);
    }

    // 문서 맨 끝으로 이동
    await controller.pressKeys({ keys: 'Ctrl+End' });
    await sleep(300);

    const mapper = new MenuMapper({ product: 'hwp', probeDialogs: false, probeDropdowns: false });

    // 초기 도형 탭 상태 확인
    process.stderr.write('▶ 초기 도형 탭 상태 확인\n');
    const initialTabs = await mapper._collectMenuTabs();
    const initialShape = initialTabs.find((t) => t.name === '도형');
    process.stderr.write(`  도형 탭: ${initialShape ? (initialShape.isEnabled ? 'enabled' : 'disabled') : 'not found'}\n`);

    // 편집 탭으로 전환
    const editTab = initialTabs.find((t) => t.name === '편집');
    if (!editTab) throw new Error('편집 탭 없음');
    await mapper._switchTab(editTab);
    await sleep(500);

    // 도형 드롭다운 버튼 찾기
    const shapeButton = mapper._refreshItem('도형 : ALT+H');
    if (!shapeButton) throw new Error('도형 드롭다운 버튼 없음');
    const sr = shapeButton.rect;
    const bx = Math.round(sr.left + (sr.right - sr.left) * 0.8);
    const by = Math.round(sr.top + (sr.bottom - sr.top) * 0.8);
    try { shapeButton.element.release(); } catch (_) {}

    process.stderr.write(`▶ 도형 드롭다운 열기 (${bx}, ${by})\n`);
    await controller.setForeground();
    await sleep(200);
    win32.mouseClick(bx, by);
    await sleep(1500);

    // 드롭다운에서 "직사각형" 찾기
    sessionModule.refreshHwpElement();
    const session = sessionModule.getSession();
    if (!session.hwpElement) throw new Error('HWP element 없음');

    const target = await findInLatestPopup(session.hwpElement, '직사각형');
    if (!target) {
        // 닫고 실패
        await controller.pressKeys({ keys: 'Escape' });
        throw new Error('드롭다운에서 "직사각형" 찾지 못함');
    }
    process.stderr.write(`▶ "직사각형" 발견: rect=${JSON.stringify(target.rect)}\n`);

    // "직사각형" 클릭
    const rcx = Math.round((target.rect.left + target.rect.right) / 2);
    const rcy = Math.round((target.rect.top + target.rect.bottom) / 2);
    await controller.setForeground();
    await sleep(200);
    win32.mouseClick(rcx, rcy);
    await sleep(800);

    // 이제 캔버스 클릭으로 기본 크기 도형 삽입 시도
    sessionModule.refreshHwpElement();
    const frame = sessionModule.getSession().hwpElement;
    const children = frame.findAllChildren();
    let canvas = null;
    for (const c of children) {
        try {
            if (c.className === 'HwpMainEditWnd') { canvas = c; break; }
        } catch (_) {}
    }
    children.filter((c) => c !== canvas).forEach((c) => { try { c.release(); } catch (_) {} });
    if (!canvas) throw new Error('캔버스(HwpMainEditWnd) 없음');

    const cr = canvas.boundingRect;
    const canvasCx = Math.round((cr.left + cr.right) / 2);
    const canvasCy = Math.round((cr.top + cr.bottom) / 2);
    try { canvas.release(); } catch (_) {}

    process.stderr.write(`▶ 캔버스 클릭 (${canvasCx}, ${canvasCy}) — 단일 클릭 시도\n`);
    await controller.setForeground();
    await sleep(200);
    win32.mouseClick(canvasCx, canvasCy);
    await sleep(1200);

    // 도형 탭 활성화 확인
    process.stderr.write('▶ 도형 탭 상태 재확인\n');
    const afterTabs = await mapper._collectMenuTabs();
    const afterShape = afterTabs.find((t) => t.name === '도형');
    const ok = afterShape && afterShape.isEnabled;
    process.stderr.write(`  도형 탭: ${afterShape ? (afterShape.isEnabled ? '✅ ENABLED' : '❌ disabled') : 'not found'}\n`);

    if (ok) {
        // 리본 항목 몇 개만 프리뷰
        await mapper._switchTab(afterShape);
        await sleep(500);
        const items = await mapper._collectRibbonItems(afterShape);
        process.stderr.write(`  도형 탭 리본 항목: ${items.length}개\n`);
        for (const it of items.slice(0, 8)) {
            process.stderr.write(`    - ${it.name} (${it.controlType}, dropdown=${it.hasDropdown})\n`);
        }
    } else {
        // 단일 클릭으로 삽입이 안 됐으면 drag 폴백 필요 — 이 smoke 에선 실패로 종료
        process.stderr.write('  ⚠ 단일 클릭으로 도형 삽입 안 됨 — drag 폴백 필요\n');
    }

    // 정리
    process.stderr.write('▶ 정리 (Escape + Ctrl+Z × 4)\n');
    await controller.setForeground();
    await controller.pressKeys({ keys: 'Escape' });
    await sleep(300);
    for (let i = 0; i < 4; i++) {
        await controller.pressKeys({ keys: 'Ctrl+Z' });
        await sleep(300);
    }

    // 최종 상태 확인
    const finalTabs = await mapper._collectMenuTabs();
    const finalShape = finalTabs.find((t) => t.name === '도형');
    process.stderr.write(`▶ 정리 후 도형 탭: ${finalShape ? (finalShape.isEnabled ? '⚠ STILL ENABLED' : '✅ disabled') : 'not found'}\n`);

    if (!ok) process.exit(1);
}

run().catch((e) => {
    process.stderr.write(`FATAL: ${e.message}\n${e.stack}\n`);
    process.exit(1);
});
