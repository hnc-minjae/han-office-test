'use strict';

const controller = require('./src/hwp-controller');
const sessionModule = require('./src/session');
const { TreeScope } = require('./src/uia');
const win32 = require('./src/win32');

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function findDialog(session) {
    sessionModule.refreshHwpElement();
    const topCh = session.hwpElement.findAllChildren();
    let dialog = null;
    for (const c of topCh) {
        try { if (c.className === 'DialogImpl') { dialog = c; break; } } catch (_) {}
    }
    topCh.filter(c => c !== dialog).forEach(c => { try { c.release(); } catch (_) {} });
    return dialog;
}

function collectDialogControls(dialog) {
    const descs = dialog.findAll(TreeScope.Descendants);
    const result = { TabItem: [], Button: [], Edit: [], ComboBox: [], CheckBox: [], RadioButton: [] };
    for (const el of descs) {
        try {
            const t = el.controlTypeName;
            const n = el.name || '';
            if (result[t] !== undefined) result[t].push(n);
        } catch (_) {}
        el.release();
    }
    return result;
}

function findRectInDialog(dialog, predicate) {
    const descs = dialog.findAll(TreeScope.Descendants);
    let rect = null, name = null;
    for (const el of descs) {
        try {
            if (predicate(el)) {
                rect = el.boundingRect;
                name = el.name;
                break;
            }
        } catch (_) {}
    }
    descs.forEach(el => { try { el.release(); } catch (_) {} });
    return rect ? { rect, name } : null;
}

async function openCharFormatDialog(session) {
    sessionModule.refreshHwpElement();
    const ch = session.hwpElement.findAllChildren();
    const mb = ch.find(c => { try { return c.className === 'MenuBarImpl'; } catch (_) { return false; } });
    ch.filter(c => c !== mb).forEach(c => c.release());
    const mbc = mb.findAllChildren();
    const et = mbc.find(el => { try { return el.name === '편집 E'; } catch (_) { return false; } });
    const tr = et.boundingRect;
    win32.mouseClick(tr.left + Math.round((tr.right - tr.left) * 0.3), Math.round((tr.top + tr.bottom) / 2));
    mbc.forEach(el => el.release());
    mb.release();
    await sleep(600);

    sessionModule.refreshHwpElement();
    const d = session.hwpElement.findAll(TreeScope.Descendants);
    let r = null;
    for (const el of d) {
        try {
            if (el.name && el.name.startsWith('글자 모양') && el.controlTypeName === 'Button') {
                r = el.boundingRect;
                break;
            }
        } catch (_) {}
    }
    d.forEach(el => { try { el.release(); } catch (_) {} });

    const cx = Math.round((r.left + r.right) / 2);
    const cy = r.top + Math.round((r.bottom - r.top) * 0.3);
    win32.mouseClick(cx, cy);
    await sleep(1500);
}

async function runTest() {
    await controller.attach({ product: 'hwp' });
    await controller.setForeground();
    await sleep(300);
    await controller.pressKeys({ keys: 'Escape' });
    await sleep(400);

    const session = sessionModule.getSession();

    console.log('=== Step 1: 테스트 문자 입력 + 전체 선택 ===');
    win32.clipboardPaste('가나다');
    await sleep(300);
    await controller.pressKeys({ keys: 'Ctrl+A' });
    await sleep(300);
    console.log('OK');

    console.log('\n=== Step 2: 글자 모양 다이얼로그 열기 ===');
    await openCharFormatDialog(session);

    let dialog = findDialog(session);
    if (!dialog) { console.log('실패'); process.exit(1); }
    console.log('OK — 다이얼로그 "' + dialog.name + '" 열림');

    console.log('\n=== Step 3: 탭별 컨트롤 수집 ===');
    const descs = dialog.findAll(TreeScope.Descendants);
    const tabRects = [];
    for (const el of descs) {
        try {
            if (el.controlTypeName === 'TabItem') {
                tabRects.push({ name: el.name, rect: el.boundingRect });
            }
        } catch (_) {}
    }
    descs.forEach(el => { try { el.release(); } catch (_) {} });
    dialog.release();

    for (const tab of tabRects) {
        const tx = Math.round((tab.rect.left + tab.rect.right) / 2);
        const ty = Math.round((tab.rect.top + tab.rect.bottom) / 2);
        win32.mouseClick(tx, ty);
        await sleep(500);

        dialog = findDialog(session);
        const ctrls = collectDialogControls(dialog);
        console.log('  [' + tab.name + '] Button=' + ctrls.Button.length + ' Edit=' + ctrls.Edit.length + ' ComboBox=' + ctrls.ComboBox.length + ' CheckBox=' + ctrls.CheckBox.length);
        if (ctrls.Edit.length > 0) {
            console.log('    Edit 항목: ' + ctrls.Edit.slice(0, 4).join(' | '));
        }
        dialog.release();
    }

    console.log('\n=== Step 4: 기본 탭 복귀 + 기준 크기 변경 (10 -> 20) ===');
    const basicTab = tabRects.find(t => t.name === '기본');
    win32.mouseClick(
        Math.round((basicTab.rect.left + basicTab.rect.right) / 2),
        Math.round((basicTab.rect.top + basicTab.rect.bottom) / 2)
    );
    await sleep(400);

    dialog = findDialog(session);
    const sizeCombo = findRectInDialog(dialog,
        el => el.controlTypeName === 'ComboBox' && el.name && el.name.startsWith('기준 크기'));
    dialog.release();

    if (!sizeCombo) { console.log('기준 크기 ComboBox 못찾음'); process.exit(1); }
    console.log('  대상: "' + sizeCombo.name + '" (rect=' + JSON.stringify(sizeCombo.rect) + ')');

    const cbCx = Math.round((sizeCombo.rect.left + sizeCombo.rect.right) / 2);
    const cbCy = Math.round((sizeCombo.rect.top + sizeCombo.rect.bottom) / 2);
    // 더블클릭으로 ComboBox 값 전체 선택
    win32.mouseClick(cbCx, cbCy);
    await sleep(150);
    win32.mouseClick(cbCx, cbCy);
    await sleep(200);
    await controller.pressKeys({ keys: 'Ctrl+A' });
    await sleep(200);
    win32.typeAsciiText('20');
    await sleep(300);
    console.log('  "20" 입력 완료 (Enter 미송신 — 설정 버튼이 커밋)');

    console.log('\n=== Step 5: 설정 버튼 클릭 ===');
    dialog = findDialog(session);
    const okBtn = findRectInDialog(dialog, el => el.name && el.name.startsWith('설정') && el.className === 'DialogButtonImpl');
    dialog.release();

    if (!okBtn) { console.log('설정 버튼 못찾음'); process.exit(1); }
    const okCx = Math.round((okBtn.rect.left + okBtn.rect.right) / 2);
    const okCy = Math.round((okBtn.rect.top + okBtn.rect.bottom) / 2);
    win32.mouseClick(okCx, okCy);
    await sleep(1000);

    const dialogAfter = findDialog(session);
    if (dialogAfter) {
        console.log('⚠ 다이얼로그 안 닫힘');
        dialogAfter.release();
    } else {
        console.log('  다이얼로그 정상 닫힘 ✓');
    }

    console.log('\n=== Step 6: 검증 - 재오픈해서 기준 크기 확인 ===');
    await controller.pressKeys({ keys: 'Ctrl+A' });
    await sleep(300);
    await openCharFormatDialog(session);

    dialog = findDialog(session);
    if (!dialog) { console.log('재오픈 실패'); process.exit(1); }

    // 다이얼로그 로딩 안정화 대기
    await sleep(500);
    const descs2 = dialog.findAll(TreeScope.Descendants);
    let currentSize = null;
    const all = [];
    for (const el of descs2) {
        try {
            if (el.name && el.name.startsWith('기준 크기')) {
                all.push('[' + el.controlTypeName + '] "' + el.name + '"');
                if (el.controlTypeName === 'ComboBox' && !currentSize) {
                    currentSize = el.name;
                }
            }
        } catch (_) {}
    }
    descs2.forEach(el => { try { el.release(); } catch (_) {} });
    console.log('  "기준 크기" 매칭 요소 전체: ' + all.join(', '));
    console.log('  ComboBox 값: "' + currentSize + '"');

    if (currentSize && currentSize.includes('20')) {
        console.log('  ★ 검증 성공 — 기준 크기가 20으로 변경되었습니다!');
    } else {
        console.log('  ⚠ 검증 실패 — 변경이 적용되지 않음');
    }

    const cancelBtn = findRectInDialog(dialog, el => el.name === '취소' && el.className === 'DialogButtonImpl');
    dialog.release();
    if (cancelBtn) {
        win32.mouseClick(
            Math.round((cancelBtn.rect.left + cancelBtn.rect.right) / 2),
            Math.round((cancelBtn.rect.top + cancelBtn.rect.bottom) / 2)
        );
        await sleep(500);
    }

    await controller.pressKeys({ keys: 'Ctrl+Z' });
    await sleep(200);
    await controller.pressKeys({ keys: 'Ctrl+Z' });
    await sleep(200);
    await controller.pressKeys({ keys: 'Ctrl+Z' });
    await sleep(300);
    console.log('\n완료 - 변경 원복');
}

runTest().catch(e => { console.error('ERROR:', e.message); process.exit(1); });
