#!/usr/bin/env node
'use strict';

const { McpServer } = require('@modelcontextprotocol/sdk/server/mcp.js');
const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js');
const { z } = require('zod');

// Lazy-require modules so the server can start even if individual modules fail
function getController() { return require('./hwp-controller'); }
function getCrashMonitor() { return require('./crash-monitor'); }
function getActionLogger() { return require('./action-logger'); }
function getJiraReporter() { return require('./jira-reporter'); }
function getMonkeyTest() { return require('./monkey-test'); }

const server = new McpServer({
  name: 'hwp-ui-test',
  version: '1.0.0',
});

// ---------------------------------------------------------------------------
// Logging wrapper - wraps a handler with action logging.
// Logging tools themselves must NOT be wrapped to avoid recursion.
// ---------------------------------------------------------------------------
async function withLogging(toolName, args, fn) {
  const logger = getActionLogger();
  const startTime = Date.now();
  try {
    const result = await fn();
    if (logger && logger.logAction) {
      logger.logAction({ tool: toolName, args, result, durationMs: Date.now() - startTime, success: true });
    }
    return result;
  } catch (err) {
    if (logger && logger.logAction) {
      logger.logAction({ tool: toolName, args, error: err.message, durationMs: Date.now() - startTime, success: false });
    }
    throw err;
  }
}

// ---------------------------------------------------------------------------
// Helper: wrap result in MCP tool content format
// ---------------------------------------------------------------------------
function mcpResult(data) {
  return { content: [{ type: 'text', text: JSON.stringify(data) }] };
}

// ---------------------------------------------------------------------------
// SESSION TOOLS (4)
// ---------------------------------------------------------------------------

server.tool(
  'hwp_launch',
  'Launch a Hancom Office application (한글, 한워드, 한쇼, 한셀)',
  {
    product:       z.enum(['hwp', 'hword', 'hshow', 'hcell']).optional(),
    hwpPath:       z.string().optional(),
    closeLauncher: z.boolean().optional(),
    timeoutMs:     z.number().optional(),
  },
  async (args) => withLogging('hwp_launch', args, async () => {
    const result = await getController().launch(args);
    return mcpResult(result);
  })
);

server.tool(
  'hwp_attach',
  'Attach to a running Hancom Office process (한글, 한워드, 한쇼, 한셀)',
  {
    product:     z.enum(['hwp', 'hword', 'hshow', 'hcell']).optional(),
    pid:         z.number().optional(),
    windowTitle: z.string().optional(),
  },
  async (args) => withLogging('hwp_attach', args, async () => {
    const result = await getController().attach(args);
    return mcpResult(result);
  })
);

server.tool(
  'hwp_close',
  'Close the HWP application',
  {
    saveAction: z.enum(['save', 'discard', 'cancel']).optional(),
    forceKill:  z.boolean().optional(),
  },
  async (args) => withLogging('hwp_close', args, async () => {
    const result = await getController().close(args);
    return mcpResult(result);
  })
);

server.tool(
  'hwp_status',
  'Get current HWP session status',
  {},
  async (args) => withLogging('hwp_status', args, async () => {
    const result = await getController().getStatus();
    return mcpResult(result);
  })
);

// ---------------------------------------------------------------------------
// UI EXPLORATION TOOLS (4)
// ---------------------------------------------------------------------------

server.tool(
  'hwp_get_ui_tree',
  'Get the UI automation element tree',
  {
    depth:        z.number().optional(),
    controlTypes: z.array(z.string()).optional(),
  },
  async (args) => withLogging('hwp_get_ui_tree', args, async () => {
    const result = await getController().getUiTree(args);
    return mcpResult(result);
  })
);

server.tool(
  'hwp_find_element',
  'Find UI elements by name, control type, or automation ID',
  {
    name:         z.string().optional(),
    nameContains: z.string().optional(),
    controlType:  z.string().optional(),
    automationId: z.string().optional(),
    maxResults:   z.number().optional(),
  },
  async (args) => withLogging('hwp_find_element', args, async () => {
    const result = await getController().findElement(args);
    return mcpResult(result);
  })
);

server.tool(
  'hwp_get_window_info',
  'Get information about the HWP main window',
  {},
  async (args) => withLogging('hwp_get_window_info', args, async () => {
    const result = await getController().getWindowInfo();
    return mcpResult(result);
  })
);

server.tool(
  'hwp_get_focused_element',
  'Get the currently focused UI element',
  {},
  async (args) => withLogging('hwp_get_focused_element', args, async () => {
    const result = await getController().getFocusedElement();
    return mcpResult(result);
  })
);

// ---------------------------------------------------------------------------
// INTERACTION TOOLS (8)
// ---------------------------------------------------------------------------

server.tool(
  'hwp_click_menu',
  'Click a menu item by path (e.g. ["파일", "저장"])',
  {
    path:   z.array(z.string()),
    waitMs: z.number().optional(),
  },
  async (args) => withLogging('hwp_click_menu', args, async () => {
    const result = await getController().clickMenu(args);
    return mcpResult(result);
  })
);

server.tool(
  'hwp_click_button',
  'Click a button by name or partial name',
  {
    name:         z.string().optional(),
    nameContains: z.string().optional(),
  },
  async (args) => withLogging('hwp_click_button', args, async () => {
    const result = await getController().clickButton(args);
    return mcpResult(result);
  })
);

server.tool(
  'hwp_click_element',
  'Click a UI element by index from the last find result',
  {
    index: z.number(),
  },
  async (args) => withLogging('hwp_click_element', args, async () => {
    const result = await getController().clickElement(args);
    return mcpResult(result);
  })
);

server.tool(
  'hwp_type_text',
  'Type text into the focused element',
  {
    text:         z.string(),
    useClipboard: z.boolean().optional(),
  },
  async (args) => withLogging('hwp_type_text', args, async () => {
    const result = await getController().typeText(args);
    return mcpResult(result);
  })
);

server.tool(
  'hwp_press_keys',
  'Press keyboard keys or key combinations (e.g. "Ctrl+S", "F5")',
  {
    keys:       z.string(),
    repeat:     z.number().optional(),
    intervalMs: z.number().optional(),
  },
  async (args) => withLogging('hwp_press_keys', args, async () => {
    const result = await getController().pressKeys(args);
    return mcpResult(result);
  })
);

server.tool(
  'hwp_handle_dialog',
  'Handle a dialog box by clicking a button',
  {
    buttonName:  z.string().optional(),
    timeoutMs:   z.number().optional(),
    detectOnly:  z.boolean().optional(),
  },
  async (args) => withLogging('hwp_handle_dialog', args, async () => {
    const result = await getController().handleDialog(args);
    return mcpResult(result);
  })
);

server.tool(
  'hwp_set_foreground',
  'Bring the HWP window to the foreground',
  {},
  async (args) => withLogging('hwp_set_foreground', args, async () => {
    const result = await getController().setForeground();
    return mcpResult(result);
  })
);

server.tool(
  'hwp_take_screenshot',
  'Take a screenshot of the HWP window (placeholder)',
  {
    outputPath: z.string().optional(),
  },
  async (_args) => {
    // GDI screenshot is complex - placeholder for future implementation
    return mcpResult({ error: 'not yet implemented' });
  }
);

// ---------------------------------------------------------------------------
// LOGGING TOOLS (3) - NOT wrapped with withLogging to avoid recursion
// ---------------------------------------------------------------------------

server.tool(
  'hwp_get_action_log',
  'Get the recorded action log entries',
  {
    lastN:      z.number().optional(),
    toolFilter: z.string().optional(),
  },
  async (args) => {
    try {
      const result = await getActionLogger().getLog(args);
      return mcpResult(result);
    } catch (err) {
      return mcpResult({ error: err.message });
    }
  }
);

server.tool(
  'hwp_export_report',
  'Export the action log as a report file',
  {
    outputPath: z.string().optional(),
  },
  async (args) => {
    try {
      const result = await getActionLogger().exportReport(args);
      return mcpResult(result);
    } catch (err) {
      return mcpResult({ error: err.message });
    }
  }
);

server.tool(
  'hwp_clear_log',
  'Clear the action log',
  {},
  async (_args) => {
    try {
      const result = await getActionLogger().clearLog();
      return mcpResult(result);
    } catch (err) {
      return mcpResult({ error: err.message });
    }
  }
);

// ---------------------------------------------------------------------------
// BUG REPORTING TOOLS (3) - NOT wrapped with withLogging
// ---------------------------------------------------------------------------

server.tool(
  'hwp_report_bug',
  'Report a bug to Jira',
  {
    summary:               z.string(),
    crashIndex:            z.number().optional(),
    additionalDescription: z.string().optional(),
    priority:              z.string().optional(),
    labels:                z.array(z.string()).optional(),
  },
  async (args) => {
    try {
      const result = await getJiraReporter().reportBug(args);
      return mcpResult(result);
    } catch (err) {
      return mcpResult({ error: err.message });
    }
  }
);

server.tool(
  'hwp_get_crash_history',
  'Get the crash history recorded by the crash monitor',
  {},
  async (_args) => {
    try {
      const result = await getCrashMonitor().getCrashHistory();
      return mcpResult(result);
    } catch (err) {
      return mcpResult({ error: err.message });
    }
  }
);

server.tool(
  'hwp_configure_jira',
  'Configure Jira integration settings',
  {
    baseUrl:    z.string().optional(),
    projectKey: z.string().optional(),
    apiToken:   z.string().optional(),
    email:      z.string().optional(),
    autoReport: z.boolean().optional(),
  },
  async (args) => {
    try {
      const result = await getJiraReporter().configure(args);
      return mcpResult(result);
    } catch (err) {
      return mcpResult({ error: err.message });
    }
  }
);

// ---------------------------------------------------------------------------
// MONKEY TESTING TOOLS (3) - NOT wrapped with withLogging
// ---------------------------------------------------------------------------

server.tool(
  'hwp_monkey_start',
  'Start a monkey test: randomly exercise UI features for a set duration',
  {
    product:    z.enum(['hwp', 'hword', 'hshow', 'hcell']).optional(),
    durationMs: z.number().optional(),
    seed:       z.number().optional(),
    maxCrashes: z.number().optional(),
    intervalMin: z.number().optional(),
    intervalMax: z.number().optional(),
  },
  async (args) => {
    try {
      const options = {
        product: args.product,
        durationMs: args.durationMs,
        seed: args.seed,
        maxCrashes: args.maxCrashes,
      };
      if (args.intervalMin || args.intervalMax) {
        options.interval = {
          min: args.intervalMin || 800,
          max: args.intervalMax || 3000,
        };
      }
      const result = await getMonkeyTest().startMonkeyTest(options);
      const { _promise, ...rest } = result;
      return mcpResult(rest);
    } catch (err) {
      return mcpResult({ error: err.message });
    }
  }
);

server.tool(
  'hwp_monkey_stop',
  'Stop the running monkey test and get the report',
  {},
  async (_args) => {
    try {
      const result = getMonkeyTest().stopMonkeyTest();
      return mcpResult(result);
    } catch (err) {
      return mcpResult({ error: err.message });
    }
  }
);

server.tool(
  'hwp_monkey_status',
  'Get the current monkey test status',
  {},
  async (_args) => {
    try {
      const result = getMonkeyTest().getMonkeyStatus();
      return mcpResult(result);
    } catch (err) {
      return mcpResult({ error: err.message });
    }
  }
);

// ---------------------------------------------------------------------------
// Graceful shutdown
// ---------------------------------------------------------------------------
process.on('SIGINT', () => {
  try { require('./session').endSession(); } catch (_) {}
  process.exit(0);
});
process.on('SIGTERM', () => {
  try { require('./session').endSession(); } catch (_) {}
  process.exit(0);
});

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch(console.error);
