'use strict';

/**
 * Jira REST API integration for automated bug reporting.
 * Uses Node.js built-in fetch (Node 18+).
 * Falls back gracefully when Jira is not configured.
 */

// ---------------------------------------------------------------------------
// Configuration (from environment variables, overridable at runtime)
// ---------------------------------------------------------------------------
const config = {
  baseUrl:    process.env.JIRA_BASE_URL    || '',
  projectKey: process.env.JIRA_PROJECT_KEY || '',
  apiToken:   process.env.JIRA_API_TOKEN   || '',
  email:      process.env.JIRA_EMAIL       || '',
  autoReport: process.env.JIRA_AUTO_REPORT === 'true',
  issueType:  process.env.JIRA_ISSUE_TYPE  || 'Bug',
};

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

/**
 * Build the Basic auth header value from email + apiToken.
 * @returns {string}
 */
function buildAuthHeader() {
  return 'Basic ' + Buffer.from(config.email + ':' + config.apiToken).toString('base64');
}

/**
 * Validate that Jira is configured (baseUrl, projectKey, email, apiToken all set).
 * @throws {Error} if not configured
 */
function assertConfigured() {
  if (!config.baseUrl || !config.projectKey || !config.email || !config.apiToken) {
    throw new Error(
      'Jira is not configured. Set JIRA_BASE_URL, JIRA_PROJECT_KEY, JIRA_EMAIL, and JIRA_API_TOKEN ' +
      'environment variables, or call hwp_configure_jira first.'
    );
  }
}

/**
 * Perform a Jira REST API request.
 * @param {string} method  HTTP method
 * @param {string} path    API path (e.g. "/rest/api/3/issue")
 * @param {object} [body]  Request body (will be JSON-serialized)
 * @returns {Promise<object>} Parsed JSON response
 */
async function jiraRequest(method, path, body) {
  const url = config.baseUrl.replace(/\/$/, '') + path;
  const headers = {
    'Authorization': buildAuthHeader(),
    'Content-Type': 'application/json',
    'Accept': 'application/json',
  };

  const options = { method, headers };
  if (body) {
    options.body = JSON.stringify(body);
  }

  let response;
  try {
    response = await fetch(url, options);
  } catch (err) {
    throw new Error('Jira request failed (network error): ' + err.message);
  }

  const text = await response.text();
  let json;
  try {
    json = text ? JSON.parse(text) : {};
  } catch (_) {
    json = { raw: text };
  }

  if (!response.ok) {
    const messages = json.errorMessages || (json.errors ? Object.values(json.errors) : []);
    const detail = messages.length ? messages.join('; ') : text;
    throw new Error(`Jira API error ${response.status}: ${detail}`);
  }

  return json;
}

/**
 * Build an Atlassian Document Format (ADF) description for a crash report.
 * @param {object|null} crashRecord
 * @param {string} [additionalDescription]
 * @returns {object} ADF document
 */
function buildAdfDescription(crashRecord, additionalDescription) {
  const content = [];

  // --- 크래시 요약 ---
  content.push({
    type: 'heading',
    attrs: { level: 2 },
    content: [{ type: 'text', text: '크래시 요약' }],
  });

  if (crashRecord) {
    content.push({
      type: 'paragraph',
      content: [
        { type: 'text', text: '유형: ', marks: [{ type: 'strong' }] },
        { type: 'text', text: String(crashRecord.type || 'unknown') },
        { type: 'hardBreak' },
        { type: 'text', text: '발생 시각: ', marks: [{ type: 'strong' }] },
        { type: 'text', text: String(crashRecord.timestamp || '') },
        { type: 'hardBreak' },
        { type: 'text', text: '트리거 액션: ', marks: [{ type: 'strong' }] },
        { type: 'text', text: crashRecord.triggerAction ? String(crashRecord.triggerAction) : '알 수 없음' },
      ],
    });
  } else {
    content.push({
      type: 'paragraph',
      content: [{ type: 'text', text: '크래시 레코드 없음 (수동 보고)' }],
    });
  }

  if (additionalDescription) {
    content.push({
      type: 'paragraph',
      content: [{ type: 'text', text: additionalDescription }],
    });
  }

  // --- 재현 단계 ---
  content.push({
    type: 'heading',
    attrs: { level: 2 },
    content: [{ type: 'text', text: '재현 단계' }],
  });

  const actionContext = (crashRecord && crashRecord.actionContext) ? crashRecord.actionContext : [];
  if (actionContext.length > 0) {
    content.push({
      type: 'orderedList',
      content: actionContext.map((step, i) => ({
        type: 'listItem',
        content: [{
          type: 'paragraph',
          content: [{ type: 'text', text: `[${i + 1}] ${step.tool || 'action'}: ${JSON.stringify(step.args || {})}` }],
        }],
      })),
    });
  } else {
    content.push({
      type: 'paragraph',
      content: [{ type: 'text', text: '재현 단계 정보 없음' }],
    });
  }

  // --- 환경 정보 ---
  content.push({
    type: 'heading',
    attrs: { level: 2 },
    content: [{ type: 'text', text: '환경 정보' }],
  });

  const envInfo = {
    os: process.platform + ' ' + (process.env.OS || ''),
    nodeVersion: process.version,
    hwpPath: (crashRecord && crashRecord.hwpPath) ? crashRecord.hwpPath : process.env.HWP_PATH || 'unknown',
    arch: process.arch,
  };

  content.push({
    type: 'codeBlock',
    attrs: { language: 'json' },
    content: [{ type: 'text', text: JSON.stringify(envInfo, null, 2) }],
  });

  // --- 전체 액션 로그 ---
  content.push({
    type: 'heading',
    attrs: { level: 2 },
    content: [{ type: 'text', text: '전체 액션 로그' }],
  });

  const fullLog = (crashRecord && crashRecord.fullLog) ? crashRecord.fullLog : [];
  content.push({
    type: 'codeBlock',
    attrs: { language: 'json' },
    content: [{ type: 'text', text: JSON.stringify(fullLog, null, 2) }],
  });

  return { version: 1, type: 'doc', content };
}

/**
 * Get a crash record from the session by index.
 * @param {number|undefined} crashIndex
 * @returns {object|null}
 */
function getCrashRecord(crashIndex) {
  if (crashIndex === undefined || crashIndex === null) return null;
  try {
    const session = require('./session');
    const history = session.crashHistory || [];
    return history[crashIndex] || null;
  } catch (_) {
    return null;
  }
}

// ---------------------------------------------------------------------------
// Exported API
// ---------------------------------------------------------------------------

/**
 * Update Jira configuration at runtime.
 * Returns current config with apiToken masked.
 * @param {object} options
 * @returns {object}
 */
function configure(options) {
  if (!options || typeof options !== 'object') {
    return getConfig();
  }
  if (options.baseUrl    !== undefined) config.baseUrl    = options.baseUrl;
  if (options.projectKey !== undefined) config.projectKey = options.projectKey;
  if (options.apiToken   !== undefined) config.apiToken   = options.apiToken;
  if (options.email      !== undefined) config.email      = options.email;
  if (options.autoReport !== undefined) config.autoReport = Boolean(options.autoReport);
  if (options.issueType  !== undefined) config.issueType  = options.issueType;
  return getConfig();
}

/**
 * Return current config with apiToken masked.
 * @returns {object}
 */
function getConfig() {
  return {
    baseUrl:    config.baseUrl,
    projectKey: config.projectKey,
    apiToken:   config.apiToken ? '***' + config.apiToken.slice(-4) : '',
    email:      config.email,
    autoReport: config.autoReport,
    issueType:  config.issueType,
  };
}

/**
 * Check for a duplicate Jira issue using JQL.
 * Searches for issues with the same tool and crash type within 7 days.
 * @param {object} crashRecord
 * @returns {Promise<{isDuplicate: boolean, existingKey?: string}>}
 */
async function checkDuplicate(crashRecord) {
  if (!crashRecord || !crashRecord.type) {
    return { isDuplicate: false };
  }

  const triggerTool = (crashRecord.triggerAction || crashRecord.lastAction || {}).tool || '';
  const crashType   = crashRecord.type || '';

  const jql =
    `project = "${config.projectKey}" ` +
    `AND issuetype = "${config.issueType}" ` +
    `AND summary ~ "${crashType}" ` +
    (triggerTool ? `AND summary ~ "${triggerTool}" ` : '') +
    `AND created >= -7d ` +
    `ORDER BY created DESC`;

  try {
    const result = await jiraRequest('GET', `/rest/api/3/search?jql=${encodeURIComponent(jql)}&maxResults=1`);
    if (result.issues && result.issues.length > 0) {
      return { isDuplicate: true, existingKey: result.issues[0].key };
    }
    return { isDuplicate: false };
  } catch (err) {
    // If the search fails, treat as no duplicate to avoid blocking bug reports
    process.stderr.write('[jira-reporter] checkDuplicate failed: ' + err.message + '\n');
    return { isDuplicate: false };
  }
}

/**
 * Add a comment to an existing Jira issue.
 * @param {string} issueKey
 * @param {object} crashRecord
 * @param {string} [additionalDescription]
 * @returns {Promise<void>}
 */
async function addCommentToIssue(issueKey, crashRecord, additionalDescription) {
  const commentText = [
    '자동 감지된 중복 크래시 보고.',
    additionalDescription ? '\n' + additionalDescription : '',
    '\n\n발생 시각: ' + (crashRecord ? crashRecord.timestamp : new Date().toISOString()),
  ].join('');

  await jiraRequest('POST', `/rest/api/3/issue/${issueKey}/comment`, {
    body: {
      version: 1,
      type: 'doc',
      content: [{
        type: 'paragraph',
        content: [{ type: 'text', text: commentText }],
      }],
    },
  });
}

/**
 * Report a bug to Jira.
 * If a duplicate exists, adds a comment to the existing issue.
 * Otherwise creates a new issue.
 *
 * @param {object} options
 * @param {string} options.summary
 * @param {number} [options.crashIndex]
 * @param {string} [options.additionalDescription]
 * @param {string} [options.priority='High']
 * @param {string[]} [options.labels]
 * @returns {Promise<object>}
 */
async function reportBug(options) {
  assertConfigured();

  const {
    summary,
    crashIndex,
    additionalDescription,
    priority = 'High',
    labels = ['auto-test', 'crash'],
  } = options || {};

  if (!summary) {
    throw new Error('summary is required for reportBug');
  }

  const crashRecord = getCrashRecord(crashIndex);

  // Check for duplicate
  const { isDuplicate, existingKey } = await checkDuplicate(crashRecord);
  if (isDuplicate && existingKey) {
    await addCommentToIssue(existingKey, crashRecord, additionalDescription);
    return {
      duplicate: true,
      existingKey,
      existingUrl: config.baseUrl.replace(/\/$/, '') + '/browse/' + existingKey,
    };
  }

  // Build ADF description
  const description = buildAdfDescription(crashRecord, additionalDescription);

  // Create new issue
  const issuePayload = {
    fields: {
      project:     { key: config.projectKey },
      summary,
      description,
      issuetype:   { name: config.issueType },
      priority:    { name: priority },
      labels,
    },
  };

  const created = await jiraRequest('POST', '/rest/api/3/issue', issuePayload);
  const issueKey = created.key;
  const issueUrl = config.baseUrl.replace(/\/$/, '') + '/browse/' + issueKey;

  return { created: true, issueKey, issueUrl };
}

/**
 * Automatically submit a crash record to Jira.
 * Called by crash-monitor when autoReport is enabled.
 * @param {object} crashRecord
 * @returns {Promise<void>}
 */
async function autoSubmitCrash(crashRecord) {
  if (!config.autoReport) return;

  try {
    assertConfigured();
  } catch (_) {
    return; // silently skip if not configured
  }

  const lastAction = (crashRecord.actionContext && crashRecord.actionContext.length > 0)
    ? crashRecord.actionContext[crashRecord.actionContext.length - 1]
    : (crashRecord.lastAction || {});

  const toolName   = lastAction.tool || 'unknown';
  const crashType  = crashRecord.type || 'crash';
  const summary    = `[자동감지] 한글 크래시: ${toolName} 실행 중 ${crashType}`;

  try {
    const result = await reportBug({
      summary,
      additionalDescription: '이 이슈는 크래시 모니터에 의해 자동으로 생성되었습니다.',
      priority: 'High',
      labels: ['auto-test', 'crash', 'auto-reported'],
    });

    crashRecord.reported = true;
    if (result.issueKey) {
      crashRecord.jiraKey = result.issueKey;
    } else if (result.existingKey) {
      crashRecord.jiraKey = result.existingKey;
    }
  } catch (err) {
    process.stderr.write('[jira-reporter] autoSubmitCrash failed: ' + err.message + '\n');
  }
}

module.exports = {
  configure,
  getConfig,
  reportBug,
  autoSubmitCrash,
  checkDuplicate,
};
