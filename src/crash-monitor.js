/**
 * Crash detection and monitoring module for 한글 (Hwp.exe)
 * Detects process exits, hung windows, and WER dialogs
 */
'use strict';

const sessionModule = require('./session');
const win32 = require('./win32');

// =============================================================================
// Internal helpers
// =============================================================================

/**
 * Build a synthetic "last action" entry for heartbeat-detected crashes
 */
function syntheticActionEntry() {
    return {
        seq: -1,
        tool: 'heartbeat',
        params: {},
        timestamp: new Date().toISOString(),
    };
}

// =============================================================================
// Exported API
// =============================================================================

/**
 * Check if the Hwp process is alive and responding
 * @returns {{ alive: boolean, responding: boolean, windowTitle: string }}
 */
function checkProcessStatus() {
    const session = sessionModule.getSession();
    const { pid, hwnd } = session.hwpProcess;

    const alive = win32.isProcessAlive(pid);
    const responding = alive ? !win32.isHungAppWindow(hwnd) : false;

    let windowTitle = '';
    if (session.hwpElement) {
        try { windowTitle = session.hwpElement.name; } catch (e) { /* stale element */ }
    }

    return { alive, responding, windowTitle };
}

/**
 * Determine the type of crash
 * @returns {string}  'process_exit' | 'not_responding' | 'wer_dialog' | 'unknown'
 */
function determineCrashType() {
    const session = sessionModule.getSession();
    const { pid, hwnd } = session.hwpProcess;

    // 1. Check if process is gone
    const alive = win32.isProcessAlive(pid);
    if (!alive) return 'process_exit';

    // 2. Check if window is hung
    if (win32.isHungAppWindow(hwnd)) return 'not_responding';

    // 3. Check for WER (Windows Error Reporting) dialog
    if (win32.isWerDialogPresent && win32.isWerDialogPresent(pid)) return 'wer_dialog';

    return 'unknown';
}

/**
 * Handle a detected crash: record it, optionally auto-report to Jira, clean up refs
 * @param {object} triggeringActionEntry  The action that triggered (or detected) the crash
 */
function handleCrashDetected(triggeringActionEntry) {
    const session = sessionModule.getSession();

    // Collect last 5 actions for context
    const actionHistory = session.actionHistory || [];
    const actionContext = actionHistory.slice(-5);

    const crashType = determineCrashType();

    /** @type {CrashRecord} */
    const record = {
        index: session.crashHistory.length,
        detectedAt: new Date().toISOString(),
        type: crashType,
        lastAction: triggeringActionEntry || actionContext[actionContext.length - 1] || null,
        actionContext,
        reported: false,
        jiraKey: null,
    };

    session.crashHistory.push(record);

    // Auto-report to Jira if configured
    if (process.env.JIRA_AUTO_REPORT) {
        try {
            // Lazy-require to avoid hard dependency
            const jiraReporter = require('./jira-reporter');
            jiraReporter.submitCrashReport(record).then(key => {
                record.jiraKey = key;
                record.reported = true;
            }).catch(err => {
                // Non-fatal: log but do not throw
                console.error('[crash-monitor] Jira auto-report failed:', err.message);
            });
        } catch (e) {
            console.error('[crash-monitor] jira-reporter not available:', e.message);
        }
    }

    // Clean up stale references — do NOT auto-relaunch
    if (session.hwpElement) {
        try { session.hwpElement.release(); } catch (e) { /* ignore */ }
        session.hwpElement = null;
    }
    session.hwpProcess = { pid: null, hwnd: null };
}

/**
 * Start a heartbeat interval that polls for process liveness
 * @param {number} [intervalMs=2000]
 */
function startHeartbeat(intervalMs = 2000) {
    const session = sessionModule.getSession();

    // Stop any existing heartbeat
    stopHeartbeat();

    session._heartbeatInterval = setInterval(() => {
        let currentSession;
        try { currentSession = sessionModule.getSession(); } catch (e) { return; }

        const { pid } = currentSession.hwpProcess || {};
        if (!pid) return;

        const alive = win32.isProcessAlive(pid);
        if (!alive) {
            clearInterval(currentSession._heartbeatInterval);
            currentSession._heartbeatInterval = null;
            handleCrashDetected(syntheticActionEntry());
        }
    }, intervalMs);
}

/**
 * Stop the heartbeat interval
 */
function stopHeartbeat() {
    let session;
    try { session = sessionModule.getSession(); } catch (e) { return; }

    if (session._heartbeatInterval) {
        clearInterval(session._heartbeatInterval);
        session._heartbeatInterval = null;
    }
}

/**
 * Return the full crash history for the current session
 * @returns {CrashRecord[]}
 */
function getCrashHistory() {
    const session = sessionModule.getSession();
    return session.crashHistory || [];
}

module.exports = {
    checkProcessStatus,
    determineCrashType,
    handleCrashDetected,
    startHeartbeat,
    stopHeartbeat,
    getCrashHistory,
};
