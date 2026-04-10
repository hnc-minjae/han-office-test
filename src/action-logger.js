/**
 * Action logging system with decorator pattern for HWP UI automation.
 * Wraps MCP tool handlers to record timing, process status, and crash detection.
 */
'use strict';

const fs = require('fs');
const { getSession } = require('./session');

// === Auto-incrementing sequence counter ===
let _seq = 0;

// === Tools that do not require a live process before execution ===
const NO_PRECHECK_TOOLS = new Set([
    'hwp_launch',
    'hwp_attach',
    'hwp_configure_jira',
    'hwp_status',
]);

// === Lazy crash-monitor loader (avoids circular dependency) ===
function getCrashMonitor() {
    return require('./crash-monitor');
}

/**
 * Wrap an async MCP tool handler with action logging.
 *
 * The wrapper:
 *   1. Creates an ActionEntry with auto-incremented seq, timestamp, tool, params
 *   2. Pre-checks process alive (skipped for launch/attach/configure/status tools)
 *   3. Executes the handler and measures durationMs
 *   4. Post-checks process alive + responding (IsHungAppWindow)
 *   5. If process died during action: marks result.status='crash_detected',
 *      calls crash-monitor handleCrashDetected
 *   6. Pushes the entry to session.actionLog
 *   7. Returns the handler result (or re-throws errors)
 *
 * @param {string}   toolName - MCP tool name (e.g. 'hwp_click_menu')
 * @param {Function} handler  - Async function (params) => result
 * @returns {Function} Wrapped async handler
 */
function withLogging(toolName, handler) {
    return async function loggedHandler(params) {
        const session = getSession();
        const seq = ++_seq;
        const timestamp = new Date().toISOString();

        /** @type {ActionEntry} */
        const entry = {
            seq,
            timestamp,
            tool: toolName,
            params: params || {},
            durationMs: 0,
            result: {
                status: 'success',
                data:   null,
                error:  null,
            },
            processSnapshot: {
                alive:       true,
                responding:  true,
                windowTitle: '',
            },
        };

        // --- Pre-check: verify process is alive before executing ---
        if (!NO_PRECHECK_TOOLS.has(toolName)) {
            const pid = session.hwpProcess && session.hwpProcess.pid;
            if (pid) {
                const win32 = require('./win32');
                const alive = win32.isProcessAlive(pid);
                if (!alive) {
                    entry.result.status = 'process_dead';
                    entry.result.error  = `Pre-check failed: process ${pid} is not alive`;
                    entry.durationMs    = 0;
                    session.actionLog.push(entry);
                    throw new Error(entry.result.error);
                }
            }
        }

        // --- Execute handler ---
        const startTime = Date.now();
        let handlerResult;
        try {
            handlerResult = await handler(params);
            entry.durationMs = Date.now() - startTime;
            entry.result.data = handlerResult !== undefined ? handlerResult : null;
        } catch (err) {
            entry.durationMs  = Date.now() - startTime;
            entry.result.status = 'error';
            entry.result.error  = err && err.message ? err.message : String(err);
            session.actionLog.push(entry);
            throw err;
        }

        // --- Post-check: process alive + responding ---
        const pid = session.hwpProcess && session.hwpProcess.pid;
        const hwnd = session.hwpProcess && session.hwpProcess.hwnd;
        if (pid) {
            try {
                const win32 = require('./win32');
                const alive = win32.isProcessAlive(pid);
                const responding = alive && hwnd
                    ? !win32.isHungAppWindow(hwnd)
                    : alive;

                // Capture window title for snapshot
                let windowTitle = '';
                if (session.hwpElement) {
                    try { windowTitle = session.hwpElement.name || ''; } catch (_e) { /* ignore */ }
                }

                entry.processSnapshot = { alive, responding, windowTitle };

                // Crash detection: process died during the action
                if (!alive) {
                    entry.result.status = 'crash_detected';
                    entry.result.error  = `Process ${pid} died during action "${toolName}"`;
                    try {
                        const crashMonitor = getCrashMonitor();
                        if (typeof crashMonitor.handleCrashDetected === 'function') {
                            await crashMonitor.handleCrashDetected({
                                toolName,
                                seq,
                                timestamp,
                                pid,
                            });
                        }
                    } catch (_e) {
                        // crash-monitor may not exist yet — log and continue
                        process.stderr.write(`[action-logger] crash-monitor unavailable: ${_e.message}\n`);
                    }
                }
            } catch (_e) {
                // Win32 post-check errors are non-fatal — snapshot stays at defaults
            }
        }

        session.actionLog.push(entry);
        return handlerResult;
    };
}

/**
 * Retrieve a filtered slice of the action log.
 * @param {object} [options]
 * @param {number} [options.lastN]      - Return only the last N entries
 * @param {string} [options.since]      - ISO timestamp; return entries at or after this time
 * @param {string} [options.toolFilter] - Return only entries matching this tool name
 * @returns {ActionEntry[]}
 */
function getActionLog(options) {
    const session = getSession();
    let log = session.actionLog.slice(); // shallow copy

    if (options && options.toolFilter) {
        log = log.filter(e => e.tool === options.toolFilter);
    }

    if (options && options.since) {
        const sinceMs = new Date(options.since).getTime();
        log = log.filter(e => new Date(e.timestamp).getTime() >= sinceMs);
    }

    if (options && typeof options.lastN === 'number' && options.lastN > 0) {
        log = log.slice(-options.lastN);
    }

    return log;
}

/**
 * Write a full JSON report of the action log to a file.
 * The report includes session metadata and the complete action log.
 * @param {string} outputPath - Absolute path to write the JSON report
 */
function exportReport(outputPath) {
    const session = getSession();
    const report = {
        exportedAt:  new Date().toISOString(),
        sessionId:   session.sessionId,
        hwpProcess:  {
            pid:        session.hwpProcess.pid,
            exePath:    session.hwpProcess.exePath,
            launchedAt: session.hwpProcess.launchedAt,
        },
        totalActions: session.actionLog.length,
        crashHistory: session.crashHistory,
        actionLog:    session.actionLog,
    };
    fs.writeFileSync(outputPath, JSON.stringify(report, null, 2), 'utf8');
}

/**
 * Clear the action log and reset the sequence counter.
 */
function clearLog() {
    const session = getSession();
    session.actionLog = [];
    _seq = 0;
}

module.exports = {
    withLogging,
    getActionLog,
    exportReport,
    clearLog,
};
