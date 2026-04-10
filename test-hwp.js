/**
 * 한글 (Hwp.exe) 실제 실행 + Node.js UIA 조작 테스트
 */
const { UIAutomation, ControlTypeId } = require('./src/uia');
const { spawn } = require('child_process');
const koffi = require('koffi');

const HWP_PATH = String.raw`C:\Program Files (x86)\Hnc\Office 2024\HOffice130\Bin\Hwp.exe`;

// === Win32 API ===
const user32 = koffi.load('user32.dll');
const FindWindowW = user32.func('void * __stdcall FindWindowW(void *, void *)');
const SetForegroundWindow = user32.func('int32 __stdcall SetForegroundWindow(void *)');
const PostMessageW = user32.func('int32 __stdcall PostMessageW(void *, uint32, uint64, int64)');
const SendMessageW = user32.func('int64 __stdcall SendMessageW(void *, uint32, uint64, int64)');
const GetWindowThreadProcessId = user32.func('uint32 __stdcall GetWindowThreadProcessId(void *, _Out_ uint32 *)');
const BringWindowToTop = user32.func('int32 __stdcall BringWindowToTop(void *)');
const ShowWindow = user32.func('int32 __stdcall ShowWindow(void *, int32)');
const AttachThreadInput = user32.func('int32 __stdcall AttachThreadInput(uint32, uint32, int32)');
const GetForegroundWindow = user32.func('void * __stdcall GetForegroundWindow()');
const GetWindowThreadProcessId2 = user32.func('uint32 __stdcall GetWindowThreadProcessId(void *, void *)');
const GetCurrentThreadId = koffi.load('kernel32.dll').func('uint32 __stdcall GetCurrentThreadId()');

// SendInput INPUT 구조체 (x64: 40 bytes)
// type(4) + pad(4) + KEYBDINPUT(wVk:2 + wScan:2 + dwFlags:4 + time:4 + pad:4 + dwExtraInfo:8) + pad(8) = 40
const INPUT_KEYBOARD = koffi.struct('INPUT_KEYBOARD', {
    type: 'uint32',        // offset 0
    _pad0: 'uint32',       // offset 4 (union alignment padding)
    wVk: 'uint16',         // offset 8
    wScan: 'uint16',       // offset 10
    dwFlags: 'uint32',     // offset 12
    time: 'uint32',        // offset 16
    _pad1: 'uint32',       // offset 20
    dwExtraInfo: 'uint64', // offset 24
    _pad2: koffi.array('uint8', 8),  // offset 32, total=40
});
const SendInput = user32.func('uint32 __stdcall SendInput(uint32, INPUT_KEYBOARD *, int32)');

const WM_KEYDOWN = 0x0100;
const WM_KEYUP = 0x0101;
const WM_CLOSE = 0x0010;
const VK_ESCAPE = 0x1B;
const VK_RETURN = 0x0D;

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

/**
 * UIA를 통해 PID로 윈도우 요소와 HWND를 찾기
 */
function findWindowByPid(uia, targetPid) {
    const root = uia.getRootElement();
    if (!root) return null;
    const windows = root.findAllChildren();
    let found = null;
    for (const win of windows) {
        try {
            if (win.processId === targetPid && win.className === 'FrameWindowImpl') {
                found = win;
                break;
            }
        } catch (e) {}
        if (!found) win.release();
    }
    root.release();
    return found;
}

/**
 * UIElement에서 네이티브 HWND 포인터 추출
 */
function getHwnd(element) {
    const hwndVal = element.nativeWindowHandle;
    if (hwndVal === 0n) return null;
    const buf = Buffer.alloc(8);
    buf.writeBigUInt64LE(hwndVal);
    return koffi.decode(buf, 'void *');
}

/**
 * 윈도우를 포그라운드로 가져오기
 */
function forceSetForeground(hwnd) {
    const fgWnd = GetForegroundWindow();
    const fgThread = GetWindowThreadProcessId2(fgWnd, null);
    const myThread = GetCurrentThreadId();
    if (fgThread !== myThread) {
        AttachThreadInput(fgThread, myThread, 1);
    }
    ShowWindow(hwnd, 9); // SW_RESTORE
    BringWindowToTop(hwnd);
    SetForegroundWindow(hwnd);
    if (fgThread !== myThread) {
        AttachThreadInput(fgThread, myThread, 0);
    }
}

/**
 * SendInput으로 키 전송 (포그라운드 윈도우에 전송)
 */
function makeInput(vk, flags) {
    return { type: 1, _pad0: 0, wVk: vk, wScan: 0, dwFlags: flags, time: 0, _pad1: 0, dwExtraInfo: 0, _pad2: new Array(8).fill(0) };
}

function sendKey(vk) {
    const r1 = SendInput(1, [makeInput(vk, 0)], 40);
    const r2 = SendInput(1, [makeInput(vk, 0x0002)], 40);
    return r1 > 0 && r2 > 0;
}

function sendKeyCombo(modVk, vk) {
    SendInput(1, [makeInput(modVk, 0)], 40);
    SendInput(1, [makeInput(vk, 0)], 40);
    SendInput(1, [makeInput(vk, 0x0002)], 40);
    SendInput(1, [makeInput(modVk, 0x0002)], 40);
}

/**
 * PostMessage로 특정 윈도우에 키 직접 전송 (SendInput 실패 시 백업)
 */
function postKey(hwnd, vk) {
    PostMessageW(hwnd, WM_KEYDOWN, vk, 0);
    PostMessageW(hwnd, WM_KEYUP, vk, 0);
}

function typeText(text) {
    for (const ch of text) {
        const vk = ch.toUpperCase().charCodeAt(0);
        sendKey(vk);
    }
}

// =============================================
async function main() {
    const uia = new UIAutomation();
    uia.init();
    console.log('[1] UIA initialized\n');

    // 1. 한글 실행
    console.log('[2] Launching Hwp.exe...');
    const proc = spawn(HWP_PATH, [], { detached: true, stdio: 'ignore' });
    proc.unref();
    const pid = proc.pid;
    console.log(`    PID: ${pid}`);

    // UIA로 윈도우 찾기
    console.log('    Waiting for window...');
    let hwpElem = null;
    for (let i = 0; i < 20; i++) {
        await sleep(1000);
        hwpElem = findWindowByPid(uia, pid);
        if (hwpElem) break;
        process.stdout.write('.');
    }

    if (!hwpElem) {
        console.log('\n    ERROR: Could not find Hwp window!');
        process.exit(1);
    }
    const hwnd = getHwnd(hwpElem);
    console.log(`    Found! HWND=${hwnd ? 'OK' : 'null'}\n`);
    console.log(`[3] Window: "${hwpElem.name}" class=${hwpElem.className}`);

    // Launcher 확인
    const launcherChildren = hwpElem.findAllChildren();
    console.log(`    Children: ${launcherChildren.map(c => `[${c.controlTypeName}] "${c.name}"`).join(', ')}`);
    launcherChildren.forEach(c => c.release());

    // 2. 포그라운드로 가져온 후 ESC로 런처 닫기
    console.log('\n[4] Closing launcher...');
    if (hwnd) {
        forceSetForeground(hwnd);
        await sleep(500);
        // SendInput 시도
        const sent = sendKey(VK_ESCAPE);
        console.log(`    SendInput ESC: ${sent ? 'OK' : 'FAILED'}`);
        await sleep(1500);

        // 아직 런처가 열려있으면 PostMessage로 재시도
        hwpElem.release();
        hwpElem = findWindowByPid(uia, pid);
        if (hwpElem && !hwpElem.name.includes('빈 문서')) {
            console.log('    Launcher still open, trying PostMessage...');
            postKey(hwnd, VK_ESCAPE);
            await sleep(1500);

            // 그래도 안되면 체크박스 언체크 후 닫기 버튼 시도 없이 그냥 기다림
            hwpElem.release();
            hwpElem = findWindowByPid(uia, pid);
        }
    }
    await sleep(1000);

    // UI 다시 읽기
    hwpElem.release();
    hwpElem = findWindowByPid(uia, pid);
    console.log(`    Window after ESC: "${hwpElem.name}"`);

    // 메인 편집기가 로드되었는지 확인
    const afterChildren = hwpElem.findAllChildren();
    console.log(`    Children: ${afterChildren.map(c => `[${c.controlTypeName}] "${c.name}" class=${c.className}`).join('\n              ')}`);
    afterChildren.forEach(c => c.release());

    // 3. 메뉴 아이템 탐색
    console.log('\n[5] Exploring menus...');
    const menuItems = hwpElem.findByControlType('MenuItem');
    if (menuItems.length > 0) {
        console.log(`    MenuItems (${menuItems.length}):`);
        for (const mi of menuItems.slice(0, 15)) {
            console.log(`      "${mi.name}"`);
        }
        if (menuItems.length > 15) console.log(`      ... and ${menuItems.length - 15} more`);
    } else {
        console.log('    No MenuItems found (launcher may still be open)');
    }
    menuItems.forEach(m => m.release());

    // 4. 버튼 탐색
    console.log('\n[6] Exploring buttons...');
    const buttons = hwpElem.findByControlType('Button');
    console.log(`    Buttons (${buttons.length}):`);
    for (const btn of buttons.slice(0, 10)) {
        console.log(`      "${btn.name}"`);
    }
    buttons.forEach(b => b.release());

    // 5. 텍스트 입력 테스트
    console.log('\n[7] Typing text...');
    forceSetForeground(hwnd);
    await sleep(300);
    typeText('Hello');
    await sleep(500);
    sendKey(0x20); // Space
    typeText('Hwp');
    sendKey(0x20);
    typeText('Test');
    await sleep(1000);
    console.log('    Typed "Hello Hwp Test"');

    // 6. 파일 메뉴 열기
    console.log('\n[8] Opening File menu via UIA invoke...');
    const fileMenus = hwpElem.findByName('파일 F');
    if (fileMenus.length > 0) {
        const invoked = fileMenus[0].invoke();
        console.log(`    Invoke result: ${invoked}`);
        if (!invoked) {
            // invoke 실패시 클릭으로 시도
            fileMenus[0].setFocus();
            await sleep(200);
            sendKey(VK_RETURN);
        }
        await sleep(1000);

        // 서브메뉴 탐색
        hwpElem.release();
        hwpElem = findWindowByPid(uia, pid);
        const subMenus = hwpElem.findByControlType('MenuItem');
        console.log(`    MenuItems after open (${subMenus.length}):`);
        for (const mi of subMenus.slice(0, 20)) {
            console.log(`      "${mi.name}"`);
        }
        subMenus.forEach(m => m.release());

        // 메뉴 닫기
        sendKey(VK_ESCAPE);
        await sleep(500);
    } else {
        console.log('    File menu not found');
    }
    fileMenus.forEach(f => f.release());

    // 7. 프로세스 상태 확인
    console.log('\n[9] Process status...');
    try { process.kill(pid, 0); console.log(`    PID ${pid}: RUNNING`); }
    catch { console.log(`    PID ${pid}: NOT RUNNING`); }

    // 8. 한글 종료 (Alt+F4)
    console.log('\n[10] Closing Hwp...');
    forceSetForeground(hwnd);
    await sleep(300);
    sendKeyCombo(0x12, 0x73); // Alt+F4
    await sleep(2000);

    // "저장 안 함" 버튼 찾기
    hwpElem.release();
    hwpElem = findWindowByPid(uia, pid);
    if (hwpElem) {
        const saveBtns = hwpElem.findByName('저장 안 함 : ALT+N');
        if (saveBtns.length > 0) {
            console.log('    Found "저장 안 함" button, clicking...');
            saveBtns[0].invoke();
            await sleep(500);
            // invoke 실패 대비 키보드
            sendKey(0x4E); // N
            await sleep(1000);
        } else {
            console.log('    No save dialog (sending N key)...');
            sendKey(0x4E); // Alt+N (저장 안 함)
            await sleep(1000);
        }
        saveBtns.forEach(b => b.release());
        hwpElem.release();
    }

    await sleep(1000);
    try { process.kill(pid, 0); console.log(`    PID ${pid}: still running`); }
    catch { console.log(`    PID ${pid}: closed successfully`); }

    uia.destroy();
    console.log('\n=== Test Complete ===');
}

main().catch(e => { console.error('Test failed:', e); process.exit(1); });
