'use strict';
/** SpinImpl(Edit) 값 변경 방법 디버깅 */
const controller = require('./src/hwp-controller');
const sessionModule = require('./src/session');
const { TreeScope } = require('./src/uia');
const win32 = require('./src/win32');
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function findDialog(session) {
    sessionModule.refreshHwpElement();
    const topCh = session.hwpElement.findAllChildren();
    let d = null;
    for (const c of topCh) { try { if (c.className === 'DialogImpl') { d = c; break; } } catch (_) {} }
    topCh.filter(c => c !== d).forEach(c => { try { c.release(); } catch (_) {} });
    return d;
}

async function readXName(session) {
    const dialog = findDialog(session);
    const descs = dialog.findAll(TreeScope.Descendants);
    let name = null;
    for (const el of descs) {
        try {
            if (el.controlTypeName === 'Edit' && el.name && el.name.startsWith('X 방향')) {
                name = el.name;
                break;
            }
        } catch (_) {}
    }
    descs.forEach(el => { try { el.release(); } catch (_) {} });
    dialog.release();
    return name;
}

async function openAndGoExt(session) {
    sessionModule.refreshHwpElement();
    const ch = session.hwpElement.findAllChildren();
    const mb = ch.find(c => { try { return c.className === 'MenuBarImpl'; } catch (_) { return false; } });
    ch.filter(c => c !== mb).forEach(c => c.release());
    const mbc = mb.findAllChildren();
    const et = mbc.find(el => { try { return el.name === '편집 E'; } catch (_) { return false; } });
    const tr = et.boundingRect;
    win32.mouseClick(tr.left + Math.round((tr.right - tr.left) * 0.3), Math.round((tr.top + tr.bottom) / 2));
    mbc.forEach(el => el.release()); mb.release();
    await sleep(600);
    sessionModule.refreshHwpElement();
    const d = session.hwpElement.findAll(TreeScope.Descendants);
    let r = null;
    for (const el of d) {
        try { if (el.name && el.name.startsWith('글자 모양') && el.controlTypeName === 'Button') { r = el.boundingRect; break; } } catch (_) {}
    }
    d.forEach(el => { try { el.release(); } catch (_) {} });
    win32.mouseClick(Math.round((r.left + r.right) / 2), r.top + Math.round((r.bottom - r.top) * 0.3));
    await sleep(1500);

    // 확장 탭
    const dialog = findDialog(session);
    const descs = dialog.findAll(TreeScope.Descendants);
    let ext = null;
    for (const el of descs) { try { if (el.controlTypeName === 'TabItem' && el.name === '확장') { ext = el.boundingRect; break; } } catch (_) {} }
    descs.forEach(el => { try { el.release(); } catch (_) {} });
    dialog.release();
    win32.mouseClick(Math.round((ext.left + ext.right) / 2), Math.round((ext.top + ext.bottom) / 2));
    await sleep(500);
}

async function tryMethod(session, label, fn) {
    // 초기화: 취소로 닫고 재오픈
    const dialog = findDialog(session);
    if (dialog) {
        const descs = dialog.findAll(TreeScope.Descendants);
        let cancel = null;
        for (const el of descs) { try { if (el.name === '취소' && el.className === 'DialogButtonImpl') { cancel = el.boundingRect; break; } } catch (_) {} }
        descs.forEach(el => { try { el.release(); } catch (_) {} });
        dialog.release();
        if (cancel) {
            win32.mouseClick(Math.round((cancel.left + cancel.right) / 2), Math.round((cancel.top + cancel.bottom) / 2));
            await sleep(500);
        }
    }

    await controller.pressKeys({ keys: 'Ctrl+A' });
    await sleep(200);
    await openAndGoExt(session);

    const before = await readXName(session);
    console.log('[' + label + '] 변경 전: "' + before + '"');

    await fn();
    await sleep(300);

    // 포커스 확인
    const f = await controller.getFocusedElement();
    console.log('  타이핑 후 focused: [' + f.controlType + '] "' + f.name + '" class=' + f.className);

    // 설정 클릭
    const d2 = findDialog(session);
    const descs2 = d2.findAll(TreeScope.Descendants);
    let okBtn = null;
    for (const el of descs2) { try { if (el.name && el.name.startsWith('설정') && el.className === 'DialogButtonImpl') { okBtn = el.boundingRect; break; } } catch (_) {} }
    descs2.forEach(el => { try { el.release(); } catch (_) {} });
    d2.release();
    win32.mouseClick(Math.round((okBtn.left + okBtn.right) / 2), Math.round((okBtn.top + okBtn.bottom) / 2));
    await sleep(800);

    // 재오픈 + 값 확인
    await controller.pressKeys({ keys: 'Ctrl+A' });
    await sleep(200);
    await openAndGoExt(session);
    const after = await readXName(session);
    console.log('  변경 후: "' + after + '"');
    const ok = after && (after.includes(' 30 ') || after.includes(' 25 '));
    console.log('  ' + (ok ? '★ 성공' : '⚠ 실패'));

    // 원복
    const d3 = findDialog(session);
    const descs3 = d3.findAll(TreeScope.Descendants);
    let cancelBtn = null;
    for (const el of descs3) { try { if (el.name === '취소' && el.className === 'DialogButtonImpl') { cancelBtn = el.boundingRect; break; } } catch (_) {} }
    descs3.forEach(el => { try { el.release(); } catch (_) {} });
    d3.release();
    if (cancelBtn) {
        win32.mouseClick(Math.round((cancelBtn.left + cancelBtn.right) / 2), Math.round((cancelBtn.top + cancelBtn.bottom) / 2));
        await sleep(500);
    }
    for (let i = 0; i < 4; i++) { await controller.pressKeys({ keys: 'Ctrl+Z' }); await sleep(200); }
    return ok;
}

async function getXRect(session) {
    const dialog = findDialog(session);
    const descs = dialog.findAll(TreeScope.Descendants);
    let rect = null;
    for (const el of descs) {
        try { if (el.controlTypeName === 'Edit' && el.name && el.name.startsWith('X 방향')) { rect = el.boundingRect; break; } } catch (_) {}
    }
    descs.forEach(el => { try { el.release(); } catch (_) {} });
    dialog.release();
    return rect;
}

async function run() {
    await controller.attach({ product: 'hwp' });
    await controller.setForeground();
    await sleep(300);
    for (let i = 0; i < 3; i++) { await controller.pressKeys({ keys: 'Escape' }); await sleep(200); }
    const session = sessionModule.getSession();

    win32.clipboardPaste('가나다');
    await sleep(300);
    await controller.pressKeys({ keys: 'Ctrl+A' });
    await sleep(300);
    await openAndGoExt(session);

    // 방법 1: Alt+X (액세스 키)
    await tryMethod(session, '방법1 Alt+X', async () => {
        await controller.pressKeys({ keys: 'Alt+X' });
        await sleep(200);
        win32.typeAsciiText('30');
    });

    // 방법 2: 클릭 + End + Shift+Home + 타이핑
    await tryMethod(session, '방법2 클릭+End+Shift+Home', async () => {
        const r = await getXRect(session);
        const cx = Math.round((r.left + r.right) / 2);
        const cy = Math.round((r.top + r.bottom) / 2);
        win32.mouseClick(cx, cy);
        await sleep(200);
        await controller.pressKeys({ keys: 'End' });
        await sleep(100);
        await controller.pressKeys({ keys: 'Shift+Home' });
        await sleep(100);
        win32.typeAsciiText('30');
    });

    // 방법 3: 트리플 클릭 + 타이핑
    await tryMethod(session, '방법3 트리플클릭', async () => {
        const r = await getXRect(session);
        const cx = Math.round((r.left + r.right) / 2);
        const cy = Math.round((r.top + r.bottom) / 2);
        win32.mouseClick(cx, cy); await sleep(80);
        win32.mouseClick(cx, cy); await sleep(80);
        win32.mouseClick(cx, cy); await sleep(200);
        win32.typeAsciiText('30');
    });

    // 방법 4: 클릭 + Backspace 여러 번 + 타이핑
    await tryMethod(session, '방법4 클릭+Backspace', async () => {
        const r = await getXRect(session);
        const cx = Math.round((r.left + r.right) / 2);
        const cy = Math.round((r.top + r.bottom) / 2);
        win32.mouseClick(cx, cy);
        await sleep(200);
        await controller.pressKeys({ keys: 'End' });
        await sleep(100);
        for (let i = 0; i < 5; i++) { await controller.pressKeys({ keys: 'Backspace' }); await sleep(50); }
        win32.typeAsciiText('30');
    });

    // 방법 5: 스피너 위쪽 화살표 여러번 (10 → 25로 증감)
    await tryMethod(session, '방법5 UpArrow x15', async () => {
        const r = await getXRect(session);
        const cx = Math.round((r.left + r.right) / 2);
        const cy = Math.round((r.top + r.bottom) / 2);
        win32.mouseClick(cx, cy);
        await sleep(200);
        for (let i = 0; i < 15; i++) { await controller.pressKeys({ keys: 'Up' }); await sleep(40); }
    });
}

run().catch(e => { console.error('ERROR:', e.message); process.exit(1); });
