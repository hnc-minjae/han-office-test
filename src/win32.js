/**
 * Win32 API wrapper for HWP UI automation
 * Organizes Win32 functions needed for keyboard/window control.
 * Uses koffi for native DLL binding (x64 Windows only).
 */
'use strict';

const koffi = require('koffi');
const { execSync, spawnSync } = require('child_process');

// === Lazy DLL cache ===
let _user32 = null;
let _kernel32 = null;

/**
 * Load user32.dll once and cache it.
 * @returns {object} koffi library handle
 */
function loadUser32() {
    if (!_user32) _user32 = koffi.load('user32.dll');
    return _user32;
}

/**
 * Load kernel32.dll once and cache it.
 * @returns {object} koffi library handle
 */
function loadKernel32() {
    if (!_kernel32) _kernel32 = koffi.load('kernel32.dll');
    return _kernel32;
}

// === Virtual key constants ===
const VK_ESCAPE  = 0x1B;
const VK_RETURN  = 0x0D;
const VK_TAB     = 0x09;
const VK_CONTROL = 0x11;
const VK_ALT     = 0x12;
const VK_SHIFT   = 0x10;
const VK_F1      = 0x70;
const VK_F2      = 0x71;
const VK_F3      = 0x72;
const VK_F4      = 0x73;
const VK_F5      = 0x74;
const VK_F6      = 0x75;
const VK_F7      = 0x76;
const VK_F8      = 0x77;
const VK_F9      = 0x78;
const VK_F10     = 0x79;
const VK_F11     = 0x7A;
const VK_F12     = 0x7B;

// === Window message constants ===
const WM_KEYDOWN = 0x0100;
const WM_KEYUP   = 0x0101;
const WM_CLOSE   = 0x0010;

// === KEYEVENTF flags ===
const KEYEVENTF_KEYUP = 0x0002;

// === ShowWindow commands ===
const SW_RESTORE = 9;

// === MOUSEEVENTF flags ===
const MOUSEEVENTF_LEFTDOWN  = 0x0002;
const MOUSEEVENTF_LEFTUP    = 0x0004;
const MOUSEEVENTF_RIGHTDOWN = 0x0008;
const MOUSEEVENTF_RIGHTUP   = 0x0010;

// === SendInput INPUT struct (x64: 40 bytes) ===
// type(4) + _pad0(4) + wVk(2) + wScan(2) + dwFlags(4) + time(4) + _pad1(4) + dwExtraInfo(8) + _pad2(8) = 40
const INPUT_KEYBOARD = koffi.struct('INPUT_KEYBOARD', {
    type:        'uint32',
    _pad0:       'uint32',
    wVk:         'uint16',
    wScan:       'uint16',
    dwFlags:     'uint32',
    time:        'uint32',
    _pad1:       'uint32',
    dwExtraInfo: 'uint64',
    _pad2:       koffi.array('uint8', 8),
});

// === SendInput INPUT_MOUSE struct (x64: 40 bytes) ===
// type(4) + _pad0(4) + dx(4) + dy(4) + mouseData(4) + dwFlags(4) + time(4) + _pad1(4) + dwExtraInfo(8) = 40
const INPUT_MOUSE = koffi.struct('INPUT_MOUSE', {
    type:        'uint32',
    _pad0:       'uint32',
    dx:          'int32',
    dy:          'int32',
    mouseData:   'uint32',
    dwFlags:     'uint32',
    time:        'uint32',
    _pad1:       'uint32',
    dwExtraInfo: 'uint64',
});

// === Lazy-initialized API function cache ===
let _apis = null;

function getApis() {
    if (_apis) return _apis;
    const u32 = loadUser32();
    const k32 = loadKernel32();
    _apis = {
        SendInput:                u32.func('uint32 __stdcall SendInput(uint32, INPUT_KEYBOARD *, int32)'),
        SendInputMouse:           u32.func('uint32 __stdcall SendInput(uint32, INPUT_MOUSE *, int32)'),
        SetCursorPos:             u32.func('int32 __stdcall SetCursorPos(int32, int32)'),
        SetForegroundWindow:      u32.func('int32 __stdcall SetForegroundWindow(void *)'),
        BringWindowToTop:         u32.func('int32 __stdcall BringWindowToTop(void *)'),
        ShowWindow:               u32.func('int32 __stdcall ShowWindow(void *, int32)'),
        GetForegroundWindow:      u32.func('void * __stdcall GetForegroundWindow()'),
        AttachThreadInput:        u32.func('int32 __stdcall AttachThreadInput(uint32, uint32, int32)'),
        GetWindowThreadProcessId: u32.func('uint32 __stdcall GetWindowThreadProcessId(void *, void *)'),
        IsHungAppWindow:          u32.func('int32 __stdcall IsHungAppWindow(void *)'),
        PostMessageW:             u32.func('int32 __stdcall PostMessageW(void *, uint32, uint64, int64)'),
        GetCurrentThreadId:       k32.func('uint32 __stdcall GetCurrentThreadId()'),
    };
    return _apis;
}

// === Internal helper: build an INPUT_KEYBOARD entry ===
function makeInput(vk, flags) {
    return {
        type:        1,  // INPUT_KEYBOARD = 1
        _pad0:       0,
        wVk:         vk,
        wScan:       0,
        dwFlags:     flags,
        time:        0,
        _pad1:       0,
        dwExtraInfo: 0,
        _pad2:       new Array(8).fill(0),
    };
}

/**
 * Send a single virtual key press (keydown + keyup) via SendInput.
 * @param {number} vk - Virtual key code
 * @returns {boolean} true if both events were accepted
 */
function sendKey(vk) {
    const { SendInput } = getApis();
    const r1 = SendInput(1, [makeInput(vk, 0)], 40);
    const r2 = SendInput(1, [makeInput(vk, KEYEVENTF_KEYUP)], 40);
    return r1 > 0 && r2 > 0;
}

/**
 * Send a modifier + key combo (e.g. Ctrl+S, Alt+F4) via SendInput.
 * @param {number} modVk - Modifier virtual key (VK_CONTROL, VK_ALT, etc.)
 * @param {number} vk    - Main virtual key
 */
function sendKeyCombo(modVk, vk) {
    const { SendInput } = getApis();
    SendInput(1, [makeInput(modVk, 0)], 40);
    SendInput(1, [makeInput(vk, 0)], 40);
    SendInput(1, [makeInput(vk, KEYEVENTF_KEYUP)], 40);
    SendInput(1, [makeInput(modVk, KEYEVENTF_KEYUP)], 40);
}

/**
 * Type a string of ASCII characters via SendInput (one key per char).
 * Uses uppercase char code as VK (works for A-Z, 0-9).
 * @param {string} text - ASCII text to type
 */
function typeAsciiText(text) {
    for (const ch of text) {
        sendKey(ch.toUpperCase().charCodeAt(0));
    }
}

/**
 * Click at absolute screen coordinates via SetCursorPos + SendInput.
 * @param {number} x - Screen X coordinate
 * @param {number} y - Screen Y coordinate
 * @param {object} [options]
 * @param {boolean} [options.rightClick=false] - Use right click instead of left
 */
function mouseClick(x, y, options = {}) {
    const { rightClick = false } = options;
    const apis = getApis();

    apis.SetCursorPos(x, y);
    // Brief pause for cursor to settle
    const start = Date.now();
    while (Date.now() - start < 30) { /* spin */ }

    const downFlag = rightClick ? MOUSEEVENTF_RIGHTDOWN : MOUSEEVENTF_LEFTDOWN;
    const upFlag   = rightClick ? MOUSEEVENTF_RIGHTUP   : MOUSEEVENTF_LEFTUP;

    const makeMouseInput = (flags) => ({
        type: 0,  // INPUT_MOUSE = 0
        _pad0: 0, dx: 0, dy: 0, mouseData: 0,
        dwFlags: flags, time: 0, _pad1: 0, dwExtraInfo: 0,
    });

    apis.SendInputMouse(1, [makeMouseInput(downFlag)], 40);
    apis.SendInputMouse(1, [makeMouseInput(upFlag)], 40);
}

/**
 * Write text to the clipboard via PowerShell Set-Clipboard, then send Ctrl+V.
 * This approach handles full Unicode/Korean text reliably without raw Win32
 * GlobalAlloc memory management complexity.
 * @param {string} text - Text to paste (supports Unicode/Korean)
 */
function clipboardPaste(text) {
    let clipSuccess = false;

    // Primary: PowerShell -EncodedCommand (Base64 UTF-16LE, no escaping issues)
    try {
        const psCommand = "Set-Clipboard -Value '" + text.replace(/'/g, "''") + "'";
        const encoded = Buffer.from(psCommand, 'utf16le').toString('base64');
        execSync(
            `powershell -NoProfile -NonInteractive -EncodedCommand ${encoded}`,
            { timeout: 5000, windowsHide: true, stdio: 'ignore' }
        );
        clipSuccess = true;
    } catch (_e) {
        // Fallback: pipe UTF-16LE bytes to clip.exe
        try {
            const utf16 = Buffer.from(text, 'utf16le');
            spawnSync('clip', [], {
                input: utf16,
                timeout: 3000,
                windowsHide: true,
                stdio: ['pipe', 'ignore', 'ignore'],
            });
            clipSuccess = true;
        } catch (_e2) {
            clipSuccess = false;
        }
    }

    if (!clipSuccess) {
        throw new Error('clipboardPaste: failed to write text to clipboard');
    }

    // Small delay to ensure clipboard is populated before paste
    const start = Date.now();
    while (Date.now() - start < 100) { /* spin */ }

    // Send Ctrl+V to paste
    sendKeyCombo(VK_CONTROL, 0x56); // 0x56 = 'V'
}

/**
 * Send WM_KEYDOWN + WM_KEYUP to a specific window via PostMessageW.
 * Useful as a fallback when SendInput cannot reach background windows.
 * @param {*} hwnd - Native window handle (koffi void*)
 * @param {number} vk - Virtual key code
 */
function postKey(hwnd, vk) {
    const { PostMessageW } = getApis();
    PostMessageW(hwnd, WM_KEYDOWN, vk, 0);
    PostMessageW(hwnd, WM_KEYUP, vk, 0);
}

/**
 * Bring a window to the foreground by attaching thread input, restoring,
 * bringing to top, and calling SetForegroundWindow.
 * @param {*} hwnd - Native window handle (koffi void*)
 */
function forceSetForeground(hwnd) {
    const apis = getApis();
    const fgWnd = apis.GetForegroundWindow();
    const fgThread = apis.GetWindowThreadProcessId(fgWnd, null);
    const myThread = apis.GetCurrentThreadId();
    if (fgThread !== myThread) {
        apis.AttachThreadInput(fgThread, myThread, 1);
    }
    apis.ShowWindow(hwnd, SW_RESTORE);
    apis.BringWindowToTop(hwnd);
    apis.SetForegroundWindow(hwnd);
    if (fgThread !== myThread) {
        apis.AttachThreadInput(fgThread, myThread, 0);
    }
}

/**
 * Search the UIA desktop tree for the top-level window belonging to a PID.
 * Matches windows with className 'FrameWindowImpl' (HWP main window class).
 * @param {import('./uia').UIAutomation} uia - Initialized UIAutomation instance
 * @param {number} pid - Target process ID
 * @returns {import('./uia').UIElement|null}
 */
function findWindowByPid(uia, pid) {
    const root = uia.getRootElement();
    if (!root) return null;
    const windows = root.findAllChildren();
    let found = null;
    for (const win of windows) {
        try {
            if (win.processId === pid && win.className === 'FrameWindowImpl') {
                found = win;
                break;
            }
        } catch (_e) {
            // ignore COM errors on individual elements
        }
        if (!found) win.release();
    }
    root.release();
    return found;
}

/**
 * Extract the native HWND pointer from a UIElement's nativeWindowHandle.
 * @param {import('./uia').UIElement} uiElement
 * @returns {*|null} koffi void* HWND, or null if handle is zero
 */
function getHwnd(uiElement) {
    const hwndVal = uiElement.nativeWindowHandle;
    if (hwndVal === 0n) return null;
    const buf = Buffer.alloc(8);
    buf.writeBigUInt64LE(hwndVal);
    return koffi.decode(buf, 'void *');
}

/**
 * Check whether a process is still alive via signal 0.
 * @param {number} pid
 * @returns {boolean}
 */
function isProcessAlive(pid) {
    try {
        process.kill(pid, 0);
        return true;
    } catch {
        return false;
    }
}

/**
 * Check whether a window is hung (not responding) via IsHungAppWindow.
 * @param {*} hwnd - Native window handle (koffi void*)
 * @returns {boolean} true if the window is hung
 */
function isHungAppWindow(hwnd) {
    if (!hwnd) return false;
    return getApis().IsHungAppWindow(hwnd) !== 0;
}

// === Key name -> VK code map for parseKeyExpression ===
const KEY_NAME_MAP = {
    escape: VK_ESCAPE, esc: VK_ESCAPE,
    enter:  VK_RETURN, return: VK_RETURN,
    tab:    VK_TAB,
    control: VK_CONTROL, ctrl: VK_CONTROL,
    alt:    VK_ALT,
    shift:  VK_SHIFT,
    f1:  VK_F1,  f2:  VK_F2,  f3:  VK_F3,  f4:  VK_F4,
    f5:  VK_F5,  f6:  VK_F6,  f7:  VK_F7,  f8:  VK_F8,
    f9:  VK_F9,  f10: VK_F10, f11: VK_F11, f12: VK_F12,
    space:     0x20,
    backspace: 0x08,
    delete:    0x2E, del:    0x2E,
    insert:    0x2D, ins:    0x2D,
    home:      0x24,
    end:       0x23,
    pageup:    0x21, pgup:   0x21,
    pagedown:  0x22, pgdown: 0x22,
    up:        0x26,
    down:      0x28,
    left:      0x25,
    right:     0x27,
};

/**
 * Parse a key expression string into an array of VK codes.
 * Supports "Ctrl+S", "Alt+F4", "Shift+Enter", "Escape", "F5", "S".
 * Single letters map to their uppercase VK code (A-Z = 0x41-0x5A).
 * All tokens except the last are treated as modifiers.
 *
 * @param {string} expr - Key expression (e.g. "Ctrl+S", "Alt+F4", "Enter")
 * @returns {number[]} Array of VK codes; empty array if parsing fails
 */
function parseKeyExpression(expr) {
    if (!expr || typeof expr !== 'string') return [];
    const vks = [];
    for (const part of expr.split('+').map(p => p.trim())) {
        const lower = part.toLowerCase();
        if (KEY_NAME_MAP[lower] !== undefined) {
            vks.push(KEY_NAME_MAP[lower]);
        } else if (part.length === 1) {
            vks.push(part.toUpperCase().charCodeAt(0));
        } else {
            process.stderr.write(`[win32] parseKeyExpression: unknown token "${part}"\n`);
        }
    }
    return vks;
}

/**
 * Execute a parsed key expression (array of VK codes from parseKeyExpression).
 * All tokens except the last are treated as modifiers (held down during the main key press).
 * @param {number[]} vks - Array of VK codes
 */
function executeKeyExpression(vks) {
    if (!vks || vks.length === 0) return;

    if (vks.length === 1) {
        sendKey(vks[0]);
        return;
    }

    // Hold modifiers, press main key, release in reverse order
    const { SendInput } = getApis();
    const modifiers = vks.slice(0, -1);
    const mainKey = vks[vks.length - 1];

    for (const mod of modifiers) {
        SendInput(1, [makeInput(mod, 0)], 40);
    }
    SendInput(1, [makeInput(mainKey, 0)], 40);
    SendInput(1, [makeInput(mainKey, KEYEVENTF_KEYUP)], 40);
    for (const mod of modifiers.reverse()) {
        SendInput(1, [makeInput(mod, KEYEVENTF_KEYUP)], 40);
    }
}

module.exports = {
    // DLL loaders
    loadUser32,
    loadKernel32,
    // Keyboard input
    sendKey,
    sendKeyCombo,
    typeAsciiText,
    clipboardPaste,
    // Mouse input
    mouseClick,
    // Window messaging
    postKey,
    forceSetForeground,
    // UIA helpers
    findWindowByPid,
    getHwnd,
    // Process / window status
    isProcessAlive,
    isHungAppWindow,
    // Key expression parser & executor
    parseKeyExpression,
    executeKeyExpression,
    // VK constants
    VK_ESCAPE, VK_RETURN, VK_TAB,
    VK_CONTROL, VK_ALT, VK_SHIFT,
    VK_F1, VK_F2, VK_F3, VK_F4,
    VK_F5, VK_F6, VK_F7, VK_F8,
    VK_F9, VK_F10, VK_F11, VK_F12,
    // WM constants
    WM_KEYDOWN, WM_KEYUP, WM_CLOSE,
};
