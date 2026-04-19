/**
 * Stress Scenarios — 영역 간(region-crossing) copy/paste 같은 고가치 시퀀스.
 *
 * 각 시나리오는 step 배열로 정의된다. 실행기는 step을 순차 실행하고,
 * 중간에 실패해도 다음 iteration에 영향이 가지 않도록 항상 forceReturnToBody로
 * 복귀한다.
 *
 * step 타입:
 *   { kind: 'keys',   keys: 'Ctrl+A' }
 *   { kind: 'text',   text: '붙여넣을 텍스트' }
 *   { kind: 'click',  x, y }
 *   { kind: 'ribbon', tab, name }      — 맵에서 찾아 클릭
 *   { kind: 'goto',   region: 'header'|'footer'|'body' }
 *   { kind: 'wait',   ms: 300 }
 */
'use strict';

const controller = require('./hwp-controller');
const win32 = require('./win32');
const { STATES } = require('./stress-context');

const sleep = (ms) => new Promise(r => setTimeout(r, ms));
const log = (icon, msg) => process.stderr.write(`${icon} [scenario] ${msg}\n`);

// =============================================================================
// Region 진입/복귀
// =============================================================================

async function gotoRegion(region, map, ctx) {
    switch (region) {
        case 'body':
            await ctx.forceReturnToBody();
            break;
        case 'header':
            await enterHeaderOrFooter(map, /머리말/, ctx, STATES.IN_HEADER);
            break;
        case 'footer':
            await enterHeaderOrFooter(map, /꼬리말/, ctx, STATES.IN_FOOTER);
            break;
        default:
            throw new Error(`unknown region: ${region}`);
    }
}

async function enterHeaderOrFooter(map, pattern, ctx, targetState) {
    const tab = map.tabs['쪽'];
    if (!tab) throw new Error('쪽 탭 없음');
    const item = (tab.ribbonItems || []).find(i => i.name && pattern.test(i.name));
    if (!item) throw new Error(`${targetState} 진입 버튼 없음`);
    win32.mouseClick(item.clickX, item.clickY);
    await sleep(500);
    // 드롭다운에서 기본 항목(양쪽) Enter
    await controller.pressKeys({ keys: 'Enter' });
    await sleep(500);
    ctx.transition(targetState, `enterViaRibbon:${pattern}`);
}

// =============================================================================
// Step 실행기
// =============================================================================

async function execStep(step, map, ctx) {
    switch (step.kind) {
        case 'keys':
            await controller.pressKeys({ keys: step.keys });
            await sleep(150);
            break;
        case 'text':
            await controller.typeText({ text: step.text, useClipboard: true });
            await sleep(150);
            break;
        case 'click':
            win32.mouseClick(step.x, step.y);
            await sleep(200);
            break;
        case 'ribbon': {
            const tab = map.tabs[step.tab];
            if (!tab) throw new Error(`탭 없음: ${step.tab}`);
            const pattern = step.name instanceof RegExp ? step.name : new RegExp(step.name);
            const item = (tab.ribbonItems || []).find(i => i.name && pattern.test(i.name));
            if (!item) throw new Error(`항목 없음: ${step.tab} > ${step.name}`);
            win32.mouseClick(item.clickX, item.clickY);
            await sleep(250);
            break;
        }
        case 'goto':
            await gotoRegion(step.region, map, ctx);
            break;
        case 'wait':
            await sleep(step.ms || 300);
            break;
        default:
            throw new Error(`unknown step kind: ${step.kind}`);
    }
}

/**
 * 시나리오 1개 실행. 실패해도 body 복귀 후 반환.
 * @returns {{ id: string, status: 'ok'|'error', error?: string, durationMs: number }}
 */
async function runScenario(scenario, map, ctx) {
    const start = Date.now();
    try {
        for (const step of scenario.steps) {
            await execStep(step, map, ctx);
        }
        await ctx.forceReturnToBody();
        return { id: scenario.id, status: 'ok', durationMs: Date.now() - start };
    } catch (e) {
        try { await ctx.forceReturnToBody(); } catch (_) {}
        return { id: scenario.id, status: 'error', error: e.message, durationMs: Date.now() - start };
    }
}

// =============================================================================
// 시나리오 정의
// =============================================================================

const SCENARIOS = [
    {
        id: 'copy-header-to-body',
        description: '머리말 텍스트를 복사하여 본문에 붙여넣기',
        steps: [
            { kind: 'goto', region: 'header' },
            { kind: 'keys', keys: 'Ctrl+A' },
            { kind: 'keys', keys: 'Ctrl+C' },
            { kind: 'goto', region: 'body' },
            { kind: 'keys', keys: 'Ctrl+End' },
            { kind: 'keys', keys: 'Enter' },
            { kind: 'keys', keys: 'Ctrl+V' },
        ],
    },
    {
        id: 'copy-body-to-footer',
        description: '본문 단락을 복사해 꼬리말에 붙여넣기',
        steps: [
            { kind: 'goto', region: 'body' },
            { kind: 'keys', keys: 'Ctrl+Home' },
            { kind: 'keys', keys: 'Shift+End' },
            { kind: 'keys', keys: 'Ctrl+C' },
            { kind: 'goto', region: 'footer' },
            { kind: 'keys', keys: 'Ctrl+End' },
            { kind: 'keys', keys: 'Space' },
            { kind: 'keys', keys: 'Ctrl+V' },
        ],
    },
    {
        id: 'copy-body-line-and-paste-multiple',
        description: '본문 한 줄 복사 후 본문 끝에 3번 붙여넣기',
        steps: [
            { kind: 'keys', keys: 'Ctrl+Home' },
            { kind: 'keys', keys: 'Shift+End' },
            { kind: 'keys', keys: 'Ctrl+C' },
            { kind: 'keys', keys: 'Ctrl+End' },
            { kind: 'keys', keys: 'Enter' },
            { kind: 'keys', keys: 'Ctrl+V' },
            { kind: 'keys', keys: 'Enter' },
            { kind: 'keys', keys: 'Ctrl+V' },
            { kind: 'keys', keys: 'Enter' },
            { kind: 'keys', keys: 'Ctrl+V' },
        ],
    },
    {
        id: 'cut-line-paste-elsewhere',
        description: '본문 중간 한 줄 잘라내서 끝에 붙여넣기',
        steps: [
            { kind: 'keys', keys: 'Ctrl+Home' },
            { kind: 'keys', keys: 'Down' },
            { kind: 'keys', keys: 'Home' },
            { kind: 'keys', keys: 'Shift+End' },
            { kind: 'keys', keys: 'Ctrl+X' },
            { kind: 'keys', keys: 'Ctrl+End' },
            { kind: 'keys', keys: 'Enter' },
            { kind: 'keys', keys: 'Ctrl+V' },
        ],
    },
    {
        id: 'select-all-body-duplicate',
        description: '본문 전체 선택 → 복사 → 끝에 붙여넣기 (문서 확장)',
        steps: [
            { kind: 'keys', keys: 'Ctrl+A' },
            { kind: 'keys', keys: 'Ctrl+C' },
            { kind: 'keys', keys: 'Ctrl+End' },
            { kind: 'keys', keys: 'Enter' },
            { kind: 'keys', keys: 'Ctrl+V' },
        ],
    },
    {
        id: 'undo-redo-chain',
        description: 'Undo 5회 → Redo 5회 — undo/redo 스택 stress',
        steps: [
            { kind: 'keys', keys: 'Ctrl+Z' },
            { kind: 'keys', keys: 'Ctrl+Z' },
            { kind: 'keys', keys: 'Ctrl+Z' },
            { kind: 'keys', keys: 'Ctrl+Z' },
            { kind: 'keys', keys: 'Ctrl+Z' },
            { kind: 'keys', keys: 'Ctrl+Y' },
            { kind: 'keys', keys: 'Ctrl+Y' },
            { kind: 'keys', keys: 'Ctrl+Y' },
            { kind: 'keys', keys: 'Ctrl+Y' },
            { kind: 'keys', keys: 'Ctrl+Y' },
        ],
    },
    {
        id: 'paste-inside-table',
        description: '본문 텍스트 복사 → 표 셀 안에 붙여넣기',
        steps: [
            { kind: 'keys', keys: 'Ctrl+Home' },
            { kind: 'keys', keys: 'Shift+End' },
            { kind: 'keys', keys: 'Ctrl+C' },
            { kind: 'keys', keys: 'Ctrl+PageDown' },  // 다음 표로
            { kind: 'wait', ms: 300 },
            { kind: 'keys', keys: 'Ctrl+V' },
        ],
    },
    {
        id: 'find-and-replace-dialog',
        description: '찾기 다이얼로그 열기 → 취소',
        steps: [
            { kind: 'keys', keys: 'Ctrl+F' },
            { kind: 'wait', ms: 500 },
            { kind: 'keys', keys: 'Escape' },
        ],
    },
    {
        id: 'page-navigation',
        description: 'Ctrl+PageDown/PageUp 반복 — 긴 문서 탐색 stress',
        steps: [
            { kind: 'keys', keys: 'Ctrl+Home' },
            { kind: 'keys', keys: 'PageDown' },
            { kind: 'keys', keys: 'PageDown' },
            { kind: 'keys', keys: 'PageDown' },
            { kind: 'keys', keys: 'PageUp' },
            { kind: 'keys', keys: 'PageUp' },
            { kind: 'keys', keys: 'Ctrl+End' },
            { kind: 'keys', keys: 'Ctrl+Home' },
        ],
    },
    {
        id: 'bold-italic-underline-toggle',
        description: '선택한 단락에 굵게/기울임/밑줄 토글 반복',
        steps: [
            { kind: 'keys', keys: 'Ctrl+Home' },
            { kind: 'keys', keys: 'Shift+End' },
            { kind: 'keys', keys: 'Ctrl+B' },
            { kind: 'keys', keys: 'Ctrl+I' },
            { kind: 'keys', keys: 'Ctrl+U' },
            { kind: 'keys', keys: 'Ctrl+B' },
            { kind: 'keys', keys: 'Ctrl+I' },
            { kind: 'keys', keys: 'Ctrl+U' },
        ],
    },
];

module.exports = { SCENARIOS, runScenario, execStep, gotoRegion };
