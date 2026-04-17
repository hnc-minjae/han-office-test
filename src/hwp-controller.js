/**
 * High-level 한글 (Hwp.exe) automation controller
 * Wraps UIA interactions into clean MCP-friendly functions
 */
'use strict';

const { spawn } = require('child_process');
const { UIAutomation, ControlTypeId, TreeScope } = require('./uia');
const sessionModule = require('./session');
const win32 = require('./win32');

const { PRODUCTS } = sessionModule;

// Shared UIA instance (initialized once)
let _uia = null;

function getUia() {
    if (!_uia) {
        _uia = new UIAutomation();
        _uia.init();
    }
    return _uia;
}

function sleep(ms) {
    return new Promise(r => setTimeout(r, ms));
}

/**
 * Find a top-level FrameWindowImpl window by PID
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
        } catch (e) { /* ignore stale elements */ }
        if (win !== found) win.release();
    }
    root.release();
    return found;
}

/**
 * Find a FrameWindowImpl by partial or exact title
 */
function findWindowByTitle(uia, titleFragment) {
    const root = uia.getRootElement();
    if (!root) return null;
    const windows = root.findAllChildren();
    let found = null;
    for (const win of windows) {
        try {
            if (win.className === 'FrameWindowImpl' && win.name.includes(titleFragment)) {
                found = win;
                break;
            }
        } catch (e) { /* ignore */ }
        if (win !== found) win.release();
    }
    root.release();
    return found;
}

/**
 * Serialize a UIElement to a plain info object (does not release)
 */
function elementToInfo(el) {
    return {
        name: el.name,
        controlType: el.controlTypeName,
        className: el.className,
        automationId: el.automationId,
        isEnabled: el.isEnabled,
    };
}

/**
 * Recursively build a UI tree up to maxDepth
 */
function buildTree(el, depth, maxDepth, controlTypeFilter) {
    const info = elementToInfo(el);
    info.children = [];

    if (depth < maxDepth) {
        const children = el.findAllChildren();
        for (const child of children) {
            const childNode = buildTree(child, depth + 1, maxDepth, controlTypeFilter);
            child.release();
            if (!controlTypeFilter || controlTypeFilter.includes(childNode.controlType)) {
                info.children.push(childNode);
            }
        }
    }
    return info;
}

/**
 * Count all nodes in a tree
 */
function countNodes(node) {
    return 1 + (node.children || []).reduce((sum, c) => sum + countNodes(c), 0);
}

/**
 * Refresh hwpElement from session (release old, find new)
 */
function refreshHwpElement(session) {
    const uia = getUia();
    if (session.hwpElement) {
        try { session.hwpElement.release(); } catch (e) { /* ignore */ }
        session.hwpElement = null;
    }
    const el = findWindowByPid(uia, session.hwpProcess.pid);
    if (el) session.hwpElement = el;
    return session.hwpElement;
}

/**
 * Release all cached search results
 */
function clearCachedSearch(session) {
    if (session._cachedSearchResults) {
        for (const el of session._cachedSearchResults) {
            try { el.release(); } catch (e) { /* ignore */ }
        }
    }
    session._cachedSearchResults = [];
}

// =============================================================================
// Session management
// =============================================================================

/**
 * Launch a Hancom Office application and wait for its window to appear
 * @param {object} options
 * @param {string} [options.product='hwp']  'hwp' | 'hword' | 'hshow' | 'hcell'
 * @param {string} [options.hwpPath]        Override exe path (auto-resolved from product if omitted)
 * @param {boolean} [options.closeLauncher=true]
 * @param {number} [options.timeoutMs=20000]
 */
async function launch(options = {}) {
    const {
        product = 'hwp',
        hwpPath,
        closeLauncher = true,
        timeoutMs = 20000,
    } = options;

    const productInfo = PRODUCTS[product];
    if (!productInfo) {
        throw new Error(`Unknown product: "${product}". Use one of: ${Object.keys(PRODUCTS).join(', ')}`);
    }
    const exePath = hwpPath || productInfo.exe;

    const uia = getUia();
    const proc = spawn(exePath, [], { detached: true, stdio: 'ignore' });
    proc.unref();
    const pid = proc.pid;

    // Poll until window appears
    let hwpElement = null;
    const deadline = Date.now() + timeoutMs;
    while (Date.now() < deadline) {
        await sleep(1000);
        hwpElement = findWindowByPid(uia, pid);
        if (hwpElement) break;
    }

    if (!hwpElement) {
        throw new Error(`Timed out waiting for ${productInfo.name} window (PID ${pid})`);
    }

    const hwnd = win32.getHwnd(hwpElement);

    if (closeLauncher) {
        win32.forceSetForeground(hwnd);
        win32.sendKey(win32.VK_ESCAPE);
        await sleep(2000);

        // Refresh element and verify launcher is gone
        hwpElement.release();
        hwpElement = findWindowByPid(uia, pid);
        if (!hwpElement) {
            throw new Error(`${productInfo.name} window disappeared after closing launcher`);
        }

        const title = hwpElement.name;
        if (!title.includes('빈 문서')) {
            // Try PostMessage fallback
            win32.postKey(hwnd, win32.VK_ESCAPE);
            await sleep(1500);
            hwpElement.release();
            hwpElement = findWindowByPid(uia, pid);
        }
    }

    const windowTitle = hwpElement ? hwpElement.name : '';

    const session = sessionModule.initSession({
        hwpProcess: { pid, hwnd, exePath },
        hwpElement,
        product,
        _cachedSearchResults: [],
        crashHistory: [],
    });

    return {
        sessionId: session.sessionId,
        product,
        productName: productInfo.name,
        pid,
        windowTitle,
        status: 'ready',
    };
}

/**
 * Attach to an already-running Hancom Office instance
 * @param {object} options
 * @param {string} [options.product='hwp']  'hwp' | 'hword' | 'hshow' | 'hcell'
 * @param {number} [options.pid]
 * @param {string} [options.windowTitle]
 */
async function attach(options = {}) {
    const { product = 'hwp', pid, windowTitle } = options;
    const uia = getUia();

    const productInfo = PRODUCTS[product];
    if (!productInfo) {
        throw new Error(`Unknown product: "${product}". Use one of: ${Object.keys(PRODUCTS).join(', ')}`);
    }

    let hwpElement = null;

    if (pid) {
        hwpElement = findWindowByPid(uia, pid);
    } else if (windowTitle) {
        hwpElement = findWindowByTitle(uia, windowTitle);
    } else {
        // Find a FrameWindowImpl whose title matches the product name
        const root = uia.getRootElement();
        if (root) {
            const windows = root.findAllChildren();
            // First pass: match by product name
            for (const win of windows) {
                try {
                    if (win.className === 'FrameWindowImpl' && win.name.includes(productInfo.name)) {
                        hwpElement = win;
                        break;
                    }
                } catch (e) { /* ignore */ }
                if (win !== hwpElement) win.release();
            }
            // Fallback: any FrameWindowImpl if product-specific match failed
            if (!hwpElement) {
                const windows2 = root.findAllChildren();
                for (const win of windows2) {
                    try {
                        if (win.className === 'FrameWindowImpl') {
                            hwpElement = win;
                            break;
                        }
                    } catch (e) { /* ignore */ }
                    if (win !== hwpElement) win.release();
                }
            }
            root.release();
        }
    }

    if (!hwpElement) {
        throw new Error(`Could not find a running ${productInfo.name} window`);
    }

    const foundPid = hwpElement.processId;
    const hwnd = win32.getHwnd(hwpElement);
    const title = hwpElement.name;

    const session = sessionModule.initSession({
        hwpProcess: { pid: foundPid, hwnd, exePath: productInfo.exe },
        hwpElement,
        product,
        _cachedSearchResults: [],
        crashHistory: [],
    });

    return {
        sessionId: session.sessionId,
        product,
        productName: productInfo.name,
        pid: foundPid,
        windowTitle: title,
        status: 'ready',
    };
}

/**
 * Close Hwp.exe
 * @param {object} options
 * @param {string} [options.saveAction='discard']  'discard' | 'save'
 * @param {boolean} [options.forceKill=false]
 * @param {number} [options.timeoutMs=5000]
 */
async function close(options = {}) {
    const {
        saveAction = 'discard',
        forceKill = false,
        timeoutMs = 5000,
    } = options;

    const session = sessionModule.getSession();
    let totalActions = 0;

    const { hwnd, pid } = session.hwpProcess;

    win32.forceSetForeground(hwnd);
    win32.sendKeyCombo(win32.VK_ALT, win32.VK_F4);
    totalActions++;

    await sleep(2000);

    // Check for save dialog
    refreshHwpElement(session);
    if (session.hwpElement) {
        if (saveAction === 'discard') {
            const candidates = [
                '저장 안 함 : ALT+N',
                '저장 안 함',
                '아니요',
            ];
            for (const name of candidates) {
                const btns = session.hwpElement.findByName(name);
                if (btns.length > 0) {
                    btns[0].invoke();
                    totalActions++;
                    btns.forEach(b => b.release());
                    break;
                }
                btns.forEach(b => b.release());
            }
        } else if (saveAction === 'save') {
            const candidates = ['저장', '예'];
            for (const name of candidates) {
                const btns = session.hwpElement.findByName(name);
                if (btns.length > 0) {
                    btns[0].invoke();
                    totalActions++;
                    btns.forEach(b => b.release());
                    break;
                }
                btns.forEach(b => b.release());
            }
        }
    }

    // Wait for process to exit
    const deadline = Date.now() + timeoutMs;
    let exited = false;
    while (Date.now() < deadline) {
        await sleep(300);
        try { process.kill(pid, 0); } catch (e) { exited = true; break; }
    }

    if (!exited && forceKill) {
        try { process.kill(pid, 'SIGKILL'); } catch (e) { /* ignore */ }
    }

    clearCachedSearch(session);
    if (session.hwpElement) {
        try { session.hwpElement.release(); } catch (e) { /* ignore */ }
    }
    sessionModule.endSession();

    return { closed: true, totalActions };
}

/**
 * Get current session status
 */
async function getStatus() {
    const session = sessionModule.getSession();
    const { pid, hwnd } = session.hwpProcess;

    const alive = win32.isProcessAlive(pid);
    const responding = alive ? !win32.isHungAppWindow(hwnd) : false;

    let windowTitle = '';
    if (session.hwpElement) {
        try { windowTitle = session.hwpElement.name; } catch (e) { /* ignore */ }
    }

    return { alive, responding, windowTitle, pid };
}

// =============================================================================
// UI Exploration
// =============================================================================

/**
 * Get the UI element tree
 * @param {object} options
 * @param {number} [options.depth=2]
 * @param {string[]} [options.controlTypes]
 */
async function getUiTree(options = {}) {
    const { depth = 2, controlTypes } = options;
    const session = sessionModule.getSession();

    refreshHwpElement(session);
    if (!session.hwpElement) {
        throw new Error('No Hwp window found');
    }

    const tree = buildTree(session.hwpElement, 0, depth, controlTypes || null);
    const totalElements = countNodes(tree);

    return { tree: [tree], totalElements };
}

/**
 * Find elements matching criteria, cache results
 * @param {object} options
 * @param {string} [options.name]
 * @param {string} [options.nameContains]
 * @param {string} [options.controlType]
 * @param {string} [options.automationId]
 * @param {number} [options.maxResults=10]
 */
async function findElement(options = {}) {
    const { name, nameContains, controlType, automationId, maxResults = 10 } = options;
    const session = sessionModule.getSession();

    clearCachedSearch(session);
    refreshHwpElement(session);
    if (!session.hwpElement) {
        throw new Error('No Hwp window found');
    }

    const ctId = controlType ? ControlTypeId[controlType] : null;
    const descendants = session.hwpElement.findAll(TreeScope.Descendants);
    const matched = [];

    for (const el of descendants) {
        let match = true;
        try {
            if (name && el.name !== name) match = false;
            if (nameContains && !el.name.includes(nameContains)) match = false;
            if (ctId && el.controlType !== ctId) match = false;
            if (automationId && el.automationId !== automationId) match = false;
        } catch (e) {
            match = false;
        }

        if (match && matched.length < maxResults) {
            matched.push(el);
        } else {
            el.release();
        }
    }

    session._cachedSearchResults = matched;

    const elements = matched.map((el, index) => ({
        index,
        ...elementToInfo(el),
    }));

    return { elements, count: elements.length };
}

/**
 * Get basic window info
 */
async function getWindowInfo() {
    const session = sessionModule.getSession();
    refreshHwpElement(session);
    if (!session.hwpElement) {
        throw new Error('No Hwp window found');
    }
    const el = session.hwpElement;
    return {
        title: el.name,
        className: el.className,
        pid: el.processId,
        isEnabled: el.isEnabled,
    };
}

/**
 * Get the currently focused UIA element
 */
async function getFocusedElement() {
    const uia = getUia();
    const el = uia.getFocusedElement();
    if (!el) return { name: '', controlType: '', className: '', automationId: '', isEnabled: false };
    const info = elementToInfo(el);
    el.release();
    return info;
}

// =============================================================================
// Interaction
// =============================================================================

/**
 * Click through a menu path
 * @param {object} options
 * @param {string[]} options.path  e.g. ['파일 F', '새 문서']
 * @param {number} [options.waitMs=500]
 */
async function clickMenu(options = {}) {
    const { path = [], waitMs = 500 } = options;
    const session = sessionModule.getSession();

    refreshHwpElement(session);
    if (!session.hwpElement) {
        throw new Error('No Hwp window found');
    }

    for (const itemName of path) {
        const items = session.hwpElement.findByName(itemName);
        if (items.length === 0) {
            throw new Error(`Menu item not found: "${itemName}"`);
        }

        const item = items[0];
        const invoked = item.invoke();
        if (!invoked) {
            item.expand();
        }
        if (!invoked) {
            item.setFocus();
            await sleep(100);
            win32.sendKey(win32.VK_RETURN);
        }
        items.forEach(i => i.release());

        await sleep(waitMs);

        // Refresh element after menu interaction
        refreshHwpElement(session);
    }

    return { clicked: true, menuPath: path };
}

/**
 * Click a button by name or partial name
 * @param {object} options
 * @param {string} [options.name]
 * @param {string} [options.nameContains]
 */
async function clickButton(options = {}) {
    const { name, nameContains } = options;
    const session = sessionModule.getSession();

    refreshHwpElement(session);
    if (!session.hwpElement) {
        throw new Error('No Hwp window found');
    }

    const buttons = session.hwpElement.findByControlType('Button');
    let target = null;

    for (const btn of buttons) {
        const btnName = btn.name;
        if (name && btnName === name) { target = btn; }
        else if (nameContains && btnName.includes(nameContains)) { target = btn; }
        if (!target) btn.release();
        else break;
    }
    // release remaining
    buttons.forEach(b => { try { if (b !== target) b.release(); } catch (e) { /* ignore */ } });

    if (!target) {
        throw new Error(`Button not found: ${name || nameContains}`);
    }

    const buttonName = target.name;
    const clicked = target.invoke();
    if (!clicked) {
        target.setFocus();
        await sleep(100);
        win32.sendKey(win32.VK_RETURN);
    }
    target.release();

    return { clicked: true, buttonName };
}

/**
 * Click a cached search result by index
 * @param {object} options
 * @param {number} options.index
 */
async function clickElement(options = {}) {
    const { index } = options;
    const session = sessionModule.getSession();

    const cached = session._cachedSearchResults;
    if (!cached || index < 0 || index >= cached.length) {
        throw new Error(`No cached element at index ${index}. Run findElement first.`);
    }

    const el = cached[index];
    const elementName = el.name;
    const clicked = el.invoke();
    if (!clicked) {
        el.setFocus();
        await sleep(100);
        win32.sendKey(win32.VK_RETURN);
    }

    return { clicked: true, elementName };
}

/**
 * Type text into the focused editor
 * @param {object} options
 * @param {string} options.text
 * @param {boolean} [options.useClipboard=false]
 */
async function typeText(options = {}) {
    const { text, useClipboard = false } = options;
    const session = sessionModule.getSession();

    win32.forceSetForeground(session.hwpProcess.hwnd);

    if (useClipboard) {
        win32.clipboardPaste(text);
    } else {
        win32.typeAsciiText(text);
    }

    return { typed: true, length: text.length };
}

/**
 * Press keys (supports combos like 'Ctrl+S', repeated)
 * @param {object} options
 * @param {string} options.keys
 * @param {number} [options.repeat=1]
 * @param {number} [options.intervalMs=50]
 */
async function pressKeys(options = {}) {
    const { keys, repeat = 1, intervalMs = 50 } = options;

    const parsed = win32.parseKeyExpression(keys);

    for (let i = 0; i < repeat; i++) {
        if (i > 0) await sleep(intervalMs);
        win32.executeKeyExpression(parsed);
    }

    return { sent: true, keys, repeat };
}

/**
 * Detect or handle a dialog
 * @param {object} options
 * @param {string} [options.buttonName]
 * @param {number} [options.timeoutMs=3000]
 * @param {boolean} [options.detectOnly=false]
 */
async function handleDialog(options = {}) {
    const { buttonName, timeoutMs = 3000, detectOnly = false } = options;
    const session = sessionModule.getSession();

    const deadline = Date.now() + timeoutMs;
    let dialogFound = false;
    let buttons = [];
    let clicked = null;

    while (Date.now() < deadline) {
        refreshHwpElement(session);
        if (!session.hwpElement) break;

        const btns = session.hwpElement.findByControlType('Button');
        if (btns.length > 0) {
            dialogFound = true;
            buttons = btns.map(b => b.name);

            if (!detectOnly && buttonName) {
                for (const btn of btns) {
                    if (btn.name === buttonName || btn.name.includes(buttonName)) {
                        btn.invoke();
                        clicked = btn.name;
                        break;
                    }
                }
            }
            btns.forEach(b => b.release());
            break;
        }
        btns.forEach(b => b.release());
        await sleep(200);
    }

    const result = { dialogFound, buttons };
    if (clicked !== null) result.clicked = clicked;
    return result;
}

/**
 * Bring the Hwp window to foreground
 */
async function setForeground() {
    const session = sessionModule.getSession();
    win32.forceSetForeground(session.hwpProcess.hwnd);
    return { ok: true };
}

module.exports = {
    launch,
    attach,
    close,
    getStatus,
    getUiTree,
    findElement,
    getWindowInfo,
    getFocusedElement,
    clickMenu,
    clickButton,
    clickElement,
    typeText,
    pressKeys,
    handleDialog,
    setForeground,
};
