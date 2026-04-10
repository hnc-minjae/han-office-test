/**
 * Singleton session state manager for HWP UI automation.
 * Tracks the active HWP process, UIA instance, window element,
 * action log, crash history, and heartbeat monitor state.
 */
'use strict';

const crypto = require('crypto');

// === Default HWP executable path ===
const DEFAULT_HWP_PATH =
    'C:\\Program Files (x86)\\Hnc\\Office 2024\\HOffice130\\Bin\\Hwp.exe';

// === Singleton session state ===
const session = {
    sessionId:   null,          // UUID string
    hwpProcess: {
        pid:        null,       // number
        hwnd:       null,       // koffi void* (native HWND)
        exePath:    DEFAULT_HWP_PATH,
        launchedAt: null,       // Date ISO string
    },
    uia:                null,   // UIAutomation instance (lazy init)
    hwpElement:         null,   // UIElement for current HWP window
    actionLog:          [],     // ActionEntry[]
    crashHistory:       [],     // CrashRecord[]
    isMonitoring:       false,  // true while heartbeat is active
    _heartbeatInterval: null,   // setInterval handle
    _cachedSearchResults: [],   // UIElement[] from last findElement call
};

/**
 * Return the singleton session object.
 * All fields are mutable directly by callers.
 * @returns {typeof session}
 */
function getSession() {
    return session;
}

/**
 * Initialize a new session with process info.
 * Generates a fresh UUID, resets logs and state.
 * @param {number} pid      - HWP process ID
 * @param {*}      hwnd     - Native HWND (koffi void*), may be null initially
 * @param {string} [exePath] - HWP executable path (defaults to DEFAULT_HWP_PATH)
 */
function initSession(pid, hwnd, exePath) {
    // Clean up any previous session resources first
    _releaseResources();

    session.sessionId            = crypto.randomUUID();
    session.hwpProcess.pid       = pid;
    session.hwpProcess.hwnd      = hwnd || null;
    session.hwpProcess.exePath   = exePath || DEFAULT_HWP_PATH;
    session.hwpProcess.launchedAt = new Date().toISOString();
    session.actionLog            = [];
    session.crashHistory         = [];
    session.isMonitoring         = false;
    session._heartbeatInterval   = null;
    session._cachedSearchResults = [];
    // hwpElement and uia remain from previous session if present — release them
    session.hwpElement           = null;
    // uia is retained if already initialized (lazy singleton)
}

/**
 * Tear down the current session:
 * - Stops the heartbeat interval
 * - Releases cached UIElements
 * - Releases the HWP window element
 * - Clears process state (but retains the UIA instance for reuse)
 */
function endSession() {
    // Stop heartbeat
    if (session._heartbeatInterval !== null) {
        clearInterval(session._heartbeatInterval);
        session._heartbeatInterval = null;
    }
    session.isMonitoring = false;

    _releaseResources();

    // Reset process info but keep sessionId for post-mortem queries
    session.hwpProcess.pid       = null;
    session.hwpProcess.hwnd      = null;
    session.hwpProcess.launchedAt = null;
}

/**
 * Lazy-initialize and return the UIAutomation singleton.
 * The instance is created once and cached for the lifetime of the process.
 * @returns {import('./uia').UIAutomation}
 */
function getUia() {
    if (!session.uia) {
        const { UIAutomation } = require('./uia');
        const uia = new UIAutomation();
        uia.init();
        session.uia = uia;
    }
    return session.uia;
}

/**
 * Re-find the HWP window element by PID via UIA.
 * Releases any previously held hwpElement before searching.
 * Updates session.hwpElement and (if found) session.hwpProcess.hwnd.
 * @returns {import('./uia').UIElement|null} The found element, or null
 */
function refreshHwpElement() {
    // Release stale element
    if (session.hwpElement) {
        try { session.hwpElement.release(); } catch (_e) { /* ignore */ }
        session.hwpElement = null;
    }

    const pid = session.hwpProcess.pid;
    if (!pid) return null;

    const { findWindowByPid, getHwnd } = require('./win32');
    const uia = getUia();
    const elem = findWindowByPid(uia, pid);

    if (elem) {
        session.hwpElement = elem;
        // Update cached HWND
        try {
            const hwnd = getHwnd(elem);
            if (hwnd) session.hwpProcess.hwnd = hwnd;
        } catch (_e) { /* ignore */ }
    }

    return session.hwpElement;
}

/**
 * Release all cached UIElement[] from the last findElement call.
 * Calls .release() on each element and clears the array.
 */
function clearCachedSearch() {
    for (const el of session._cachedSearchResults) {
        try { el.release(); } catch (_e) { /* ignore */ }
    }
    session._cachedSearchResults = [];
}

// === Internal helper ===
function _releaseResources() {
    clearCachedSearch();

    if (session.hwpElement) {
        try { session.hwpElement.release(); } catch (_e) { /* ignore */ }
        session.hwpElement = null;
    }
}

module.exports = {
    DEFAULT_HWP_PATH,
    getSession,
    initSession,
    endSession,
    getUia,
    refreshHwpElement,
    clearCachedSearch,
};
