'use strict';
/**
 * Phase 2 Discovery — 4종 샘플 드롭다운의 UIA 노출 방식을 관찰한다.
 *
 * 대상: 붙이기, 쪽 여백, 글머리표, 표
 *
 * 각 드롭다운에 대해:
 *   1. 탭 전환 후 리본에서 항목을 찾아 boundingRect 수집
 *   2. 클릭 전 FrameWindowImpl / Desktop root 자식 스냅샷
 *   3. 항목 중앙 클릭 → 팝업 대기
 *   4. 클릭 후 스냅샷 diff로 팝업 요소 식별
 *   5. 팝업 하위 트리 순회 (depth 3, max 100 nodes, visited set)
 *   6. Escape로 닫기
 *
 * 결과: docs/discovery-dropdown-results.md
 */

const fs = require('fs');
const path = require('path');
const controller = require('./src/hwp-controller');
const sessionModule = require('./src/session');
const win32 = require('./src/win32');
const { MenuMapper } = require('./src/menu-mapper');
const { ExpandCollapseState } = require('./src/uia');

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// 대상 샘플 (plan §6 Phase 2)
// 리본 항목의 실제 UIA name은 "항목명 : 액세스키" 형식
const TARGETS = [
    '붙이기 : ALT+P',     // 편집 탭 — 단순 메뉴 (가설)
    '쪽 여백 : ALT+J',    // 편집 탭 — 갤러리(텍스트) (가설)
    '글머리표 : ALT+L',   // 서식 탭 — 복합 (가설)
    '표 : ALT+B',         // 편집 탭 — 그리드 UI (과거 실패 지점)
];

const LIMITS = {
    popupWaitMs: 1200,
    popupMaxDepth: 3,
    popupMaxNodes: 100,
    perProbeTimeoutMs: 12000,
    closeEscapeMax: 4,
    betweenProbeDelayMs: 600,
    tabSwitchWaitMs: 500,
};

// ---------------------------------------------------------------------------
// Snapshot helpers
// ---------------------------------------------------------------------------

function snapshotChildren(parentEl) {
    // 자식 포인터를 살려둔 채로 식별 정보도 수집.
    // 여러 Popup이 동시에 존재할 수 있으므로 rect까지 포함해 유일성 확보.
    const children = parentEl.findAllChildren();
    const infos = children.map((c) => {
        const info = { className: '', controlType: '', name: '', rect: null };
        try { info.className = c.className || ''; } catch (_) {}
        try { info.controlType = c.controlTypeName || ''; } catch (_) {}
        try { info.name = c.name || ''; } catch (_) {}
        try { info.rect = c.boundingRect; } catch (_) {}
        return info;
    });
    return { children, infos };
}

function releaseAll(list) {
    for (const c of list) {
        try { c.release(); } catch (_) { /* ignore */ }
    }
}

function fingerprint(info) {
    const r = info.rect;
    const rectStr = r ? `${r.left},${r.top},${r.right},${r.bottom}` : '';
    return `${info.className}|${info.controlType}|${info.name}|${rectStr}`;
}

function findAddedChildren(before, after) {
    const beforeSet = new Set(before.infos.map(fingerprint));
    const added = [];
    for (let i = 0; i < after.infos.length; i++) {
        if (!beforeSet.has(fingerprint(after.infos[i]))) {
            added.push({ info: after.infos[i], element: after.children[i], index: i });
        }
    }
    return added;
}

// ---------------------------------------------------------------------------
// Popup tree walk — visited set + depth + node limit
// ---------------------------------------------------------------------------

function walkPopupTree(el, depth, visited, nodes) {
    if (depth > LIMITS.popupMaxDepth) return;
    if (nodes.length >= LIMITS.popupMaxNodes) return;

    // 속성 먼저 수집 — visited 키를 안정적으로 만들기 위해
    const node = { depth, name: '', controlType: '', className: '', rect: null, expandCollapseState: null, isDialogOpener: false };
    try { node.name = el.name || ''; } catch (_) {}
    try { node.controlType = el.controlTypeName || ''; } catch (_) {}
    try { node.className = el.className || ''; } catch (_) {}
    try { node.rect = el.boundingRect; } catch (_) {}
    try { node.expandCollapseState = el.expandCollapseState; } catch (_) { node.expandCollapseState = 'error'; }
    node.isDialogOpener = typeof node.name === 'string' && node.name.endsWith('...');

    // (className, name, rect) 기반 visited 키 — koffi pointer 직접 String 변환 회피
    const r = node.rect;
    const visitedKey = `${node.className}|${node.name}|${r ? `${r.left},${r.top},${r.right},${r.bottom}` : ''}`;
    if (visited.has(visitedKey)) return;
    visited.add(visitedKey);

    nodes.push(node);

    if (depth >= LIMITS.popupMaxDepth) return;

    const children = el.findAllChildren();
    try {
        for (const child of children) {
            if (nodes.length >= LIMITS.popupMaxNodes) break;
            walkPopupTree(child, depth + 1, visited, nodes);
        }
    } finally {
        releaseAll(children);
    }
}

/**
 * 현재 활성 탭에서 이름으로 리본 항목을 다시 찾아 element(살아있는 포인터) + rect 반환.
 * 호출자는 사용 후 반드시 element.release() 호출 필요.
 */
function refreshItem(itemName) {
    const session = sessionModule.getSession();
    sessionModule.refreshHwpElement();
    if (!session.hwpElement) return null;

    const topChildren = session.hwpElement.findAllChildren();
    const toolbox = topChildren.find((c) => {
        try { return c.className === 'ToolBoxImpl'; } catch (_) { return false; }
    });
    topChildren.filter((c) => c !== toolbox).forEach((c) => { try { c.release(); } catch (_) {} });
    if (!toolbox) return null;

    let hit = null;
    const tbChildren = toolbox.findAllChildren();
    for (const c of tbChildren) {
        try {
            if (c.name === itemName) {
                hit = { element: c, rect: c.boundingRect };
                break;
            }
        } catch (_) {}
    }
    // 매치되지 않은 자식만 해제
    tbChildren.filter((c) => !hit || c !== hit.element).forEach((c) => { try { c.release(); } catch (_) {} });
    toolbox.release();
    return hit;
}

// ---------------------------------------------------------------------------
// Pre-sweep: 한 번의 탭 순회로 대상 전체의 좌표를 수집
// ---------------------------------------------------------------------------

async function prebuildLocations(mapper, targets) {
    const locations = {};
    const menuTabs = await mapper._collectMenuTabs();
    const remaining = new Set(targets);

    for (const tab of menuTabs) {
        if (remaining.size === 0) break;
        if (tab.name === '파일' || !tab.isEnabled) continue;

        await mapper._switchTab(tab);
        await sleep(LIMITS.tabSwitchWaitMs);

        const session = sessionModule.getSession();
        sessionModule.refreshHwpElement();
        if (!session.hwpElement) continue;

        const topChildren = session.hwpElement.findAllChildren();
        const toolbox = topChildren.find((c) => {
            try { return c.className === 'ToolBoxImpl'; } catch (_) { return false; }
        });
        topChildren.filter((c) => c !== toolbox).forEach((c) => { try { c.release(); } catch (_) {} });
        if (!toolbox) continue;

        const tbChildren = toolbox.findAllChildren();
        for (const c of tbChildren) {
            try {
                const name = c.name;
                if (name && remaining.has(name)) {
                    const loc = {
                        tab: tab.name,
                        tabClickX: tab.clickX,
                        tabClickY: tab.clickY,
                        rect: c.boundingRect,
                        controlType: c.controlTypeName,
                        className: c.className,
                    };
                    try { loc.preExpandState = c.expandCollapseState; } catch (_) { loc.preExpandState = null; }
                    locations[name] = loc;
                    remaining.delete(name);
                }
            } catch (_) {}
        }
        tbChildren.forEach((c) => { try { c.release(); } catch (_) {} });
        toolbox.release();
    }
    return { locations, notFound: Array.from(remaining) };
}

// ---------------------------------------------------------------------------
// Single probe
// ---------------------------------------------------------------------------

async function probeDropdown(target, loc) {
    const start = Date.now();
    const out = {
        target,
        tab: loc ? loc.tab : null,
        itemControlType: loc ? loc.controlType : null,
        itemClassName: loc ? loc.className : null,
        preExpandState: loc ? loc.preExpandState : null,
        clickPoint: null,
        popupFound: false,
        popupLocation: null,
        popupInfo: null,
        items: [],
        notes: [],
        durationMs: 0,
    };

    if (!loc) {
        out.notes.push('item-not-found-in-any-tab');
        out.durationMs = Date.now() - start;
        return out;
    }

    // HWP가 foreground인지 확실히 보장
    try { await controller.setForeground(); } catch (_) {}
    await sleep(150);
    // Split-button 대응: plan §6.4 — 오른쪽 하단 1/4 영역 클릭
    const rW = loc.rect.right - loc.rect.left;
    const rH = loc.rect.bottom - loc.rect.top;
    out.clickPoint = {
        x: Math.round(loc.rect.left + rW * 0.8),
        y: Math.round(loc.rect.top + rH * 0.8),
    };
    out.rectInfo = { width: rW, height: rH, ...loc.rect };

    // 대상 탭으로 재전환 (이전 프로브가 다른 탭으로 이동했을 수 있음)
    try {
        win32.mouseClick(Number(loc.tabClickX), Number(loc.tabClickY));
    } catch (e) {
        out.notes.push(`tab-click-error:${e.message}`);
        out.durationMs = Date.now() - start;
        return out;
    }
    await sleep(LIMITS.tabSwitchWaitMs);

    // 탭 전환 후 rect 재수집 — pre-sweep 시점의 좌표가 stale 할 수 있음
    const freshItem = refreshItem(target);
    let itemEl = null;
    if (freshItem) {
        itemEl = freshItem.element;
        const fR = freshItem.rect;
        const fW = fR.right - fR.left;
        const fH = fR.bottom - fR.top;
        out.clickPoint = {
            x: Math.round(fR.left + fW * 0.8),
            y: Math.round(fR.top + fH * 0.8),
        };
        out.rectInfo = { width: fW, height: fH, ...fR, refreshed: true };
    } else {
        out.notes.push('rect-refresh-failed-using-cached');
    }

    // Before snapshots
    sessionModule.refreshHwpElement();
    const session = sessionModule.getSession();
    const frame = session.hwpElement;
    const uia = sessionModule.getUia();
    const rootEl = uia.getRootElement();

    const beforeFrame = snapshotChildren(frame);
    const beforeRoot = snapshotChildren(rootEl);
    // Before 자식들은 식별 정보만 필요하므로 즉시 해제
    releaseAll(beforeFrame.children);
    releaseAll(beforeRoot.children);

    // Expand() 패턴은 한컴 MenuItem에서 no-op 확인됨 (plan §1.4) — 마우스 클릭만 사용
    if (itemEl) { try { itemEl.release(); } catch (_) {} }
    win32.mouseClick(out.clickPoint.x, out.clickPoint.y);
    await sleep(LIMITS.popupWaitMs);

    // After snapshots — 자식들은 팝업 후보이므로 살려둠
    sessionModule.refreshHwpElement();
    const frameAfter = sessionModule.getSession().hwpElement;
    const rootElAfter = uia.getRootElement();
    const afterFrame = snapshotChildren(frameAfter);
    const afterRoot = snapshotChildren(rootElAfter);

    // Diff to find new children
    const addedFrame = findAddedChildren(beforeFrame, afterFrame);
    const addedRoot = findAddedChildren(beforeRoot, afterRoot);

    let popupElement = null;
    if (addedFrame.length > 0) {
        popupElement = addedFrame[0].element;
        out.popupLocation = 'frame';
        out.popupInfo = addedFrame[0].info;
        if (addedFrame.length > 1) out.notes.push(`multiple-frame-additions:${addedFrame.length}`);
    } else if (addedRoot.length > 0) {
        popupElement = addedRoot[0].element;
        out.popupLocation = 'desktop';
        out.popupInfo = addedRoot[0].info;
        if (addedRoot.length > 1) out.notes.push(`multiple-root-additions:${addedRoot.length}`);
    }
    out.popupFound = !!popupElement;

    if (popupElement) {
        const visited = new Set();
        walkPopupTree(popupElement, 0, visited, out.items);
        if (out.items.length >= LIMITS.popupMaxNodes) out.notes.push('tree-truncated-max-nodes');
    } else {
        out.notes.push('no-popup-detected');
        // 진단: 클릭 후 frame/desktop 상태 스냅샷 기록
        out.diagnostic = {
            frameBeforeCount: beforeFrame.infos.length,
            frameAfterCount: afterFrame.infos.length,
            rootBeforeCount: beforeRoot.infos.length,
            rootAfterCount: afterRoot.infos.length,
            frameBefore: beforeFrame.infos.map(fingerprint),
            frameAfter: afterFrame.infos.map(fingerprint),
            rootAfterPopupLike: afterRoot.infos.filter((i) =>
                /Popup|Menu|Dropdown|Impl/i.test(i.className || '')
                && !/FrameWindowImpl|TaskbarFrame|Shell_/i.test(i.className || ''),
            ),
        };
    }

    // 모든 after 자식 해제 (popupElement 포함)
    releaseAll(afterFrame.children);
    releaseAll(afterRoot.children);
    try { rootEl.release(); } catch (_) {}
    try { rootElAfter.release(); } catch (_) {}

    // Close popup
    for (let i = 0; i < LIMITS.closeEscapeMax; i++) {
        await controller.pressKeys({ keys: 'Escape' });
        await sleep(200);
    }

    out.durationMs = Date.now() - start;
    return out;
}

async function probeWithTimeout(target, loc) {
    return Promise.race([
        probeDropdown(target, loc),
        new Promise((_, rej) =>
            setTimeout(() => rej(new Error(`probe-timeout:${LIMITS.perProbeTimeoutMs}ms`)),
            LIMITS.perProbeTimeoutMs),
        ),
    ]);
}

// ---------------------------------------------------------------------------
// Markdown rendering
// ---------------------------------------------------------------------------

const ExpStateLabel = {
    0: 'Collapsed',
    1: 'Expanded',
    2: 'PartiallyExpanded',
    3: 'LeafNode',
    null: '-',
};

function classifyHint(r) {
    if (!r.popupFound) return 'no-popup';
    const items = r.items || [];
    if (items.length <= 1) return 'empty-or-visual-grid';
    const inner = items.filter((i) => i.depth > 0);
    if (inner.length === 0) return 'empty-popup';
    const named = inner.filter((i) => i.name && i.name.trim());
    const menuItems = inner.filter((i) => i.controlType === 'MenuItem');
    const buttons = inner.filter((i) => i.controlType === 'Button');
    const lists = inner.filter((i) => i.controlType === 'List');
    const dialogOpeners = inner.filter((i) => i.isDialogOpener);
    if (named.length === 0) return 'gallery-visual (unnamed children)';
    if (menuItems.length > 0 && dialogOpeners.length > 0) return 'mixed';
    if (menuItems.length > 0) return 'menu';
    if (lists.length > 0 || buttons.length > 0) return 'gallery-text';
    return 'other';
}

function renderMarkdown(results) {
    const lines = [];
    lines.push('# 드롭다운 Discovery 결과');
    lines.push('');
    lines.push(`- 생성: ${new Date().toISOString()}`);
    lines.push(`- 대상: ${TARGETS.join(', ')}`);
    lines.push(`- 제한: depth=${LIMITS.popupMaxDepth}, maxNodes=${LIMITS.popupMaxNodes}, popupWait=${LIMITS.popupWaitMs}ms, probeTimeout=${LIMITS.perProbeTimeoutMs}ms`);
    lines.push('');
    lines.push('## 요약');
    lines.push('');
    lines.push('| 대상 | 탭 | 리본 controlType | 팝업 | 위치 | 자식수 | 분류 힌트 | ms |');
    lines.push('|------|----|----|------|------|------|-----------|-----|');
    for (const r of results) {
        lines.push(`| ${r.target} | ${r.tab || '-'} | ${r.itemControlType || '-'} | ${r.popupFound ? '✅' : '❌'} | ${r.popupLocation || '-'} | ${r.items ? r.items.length : 0} | ${classifyHint(r)} | ${r.durationMs || 0} |`);
    }
    lines.push('');

    for (const r of results) {
        lines.push(`## ${r.target} (${r.tab || '?'})`);
        lines.push('');
        if (r.error) {
            lines.push(`**에러:** \`${r.error}\``);
            if (r.errorStack) {
                lines.push('');
                lines.push('```');
                lines.push(r.errorStack);
                lines.push('```');
            }
            lines.push('');
            continue;
        }
        lines.push(`- 리본 항목: \`${r.itemClassName}\` / ${r.itemControlType}`);
        lines.push(`- 사전 ExpandCollapseState: ${r.preExpandState === null ? '(pattern unsupported)' : ExpStateLabel[r.preExpandState]}`);
        lines.push(`- 클릭 좌표: (${r.clickPoint?.x}, ${r.clickPoint?.y})`);
        if (r.rectInfo) {
            lines.push(`- 버튼 rect: ${r.rectInfo.left},${r.rectInfo.top}-${r.rectInfo.right},${r.rectInfo.bottom} (${r.rectInfo.width}×${r.rectInfo.height}, ${r.rectInfo.refreshed ? 'refreshed' : 'cached'})`);
        }
        lines.push(`- 팝업 컨테이너: \`${r.popupInfo?.className || 'N/A'}\` (${r.popupInfo?.controlType || 'N/A'})${r.popupInfo?.name ? ` "${r.popupInfo.name}"` : ''}`);
        lines.push(`- 팝업 위치: ${r.popupLocation || 'N/A'}`);
        lines.push(`- notes: ${r.notes && r.notes.length ? r.notes.join(', ') : '-'}`);
        if (r.diagnostic) {
            lines.push('');
            lines.push('### 진단 (팝업 미감지 시)');
            lines.push(`- frame 자식: ${r.diagnostic.frameBeforeCount} → ${r.diagnostic.frameAfterCount}`);
            lines.push(`- desktop 자식: ${r.diagnostic.rootBeforeCount} → ${r.diagnostic.rootAfterCount}`);
            if (r.diagnostic.rootAfterPopupLike && r.diagnostic.rootAfterPopupLike.length > 0) {
                lines.push('- desktop에서 popup-like 요소:');
                for (const it of r.diagnostic.rootAfterPopupLike) {
                    lines.push(`  - \`${it.className}\` (${it.controlType}) "${it.name}"`);
                }
            }
            lines.push('- frame 자식 전체 (after):');
            for (const fp of r.diagnostic.frameAfter) lines.push(`  - ${fp}`);
        }
        lines.push('');
        if (r.items && r.items.length > 0) {
            lines.push('| depth | name | controlType | className | ExpState | dialogOpener |');
            lines.push('|-------|------|-------------|-----------|----------|--------------|');
            const show = r.items.slice(0, 60);
            for (const it of show) {
                const nm = (it.name || '').replace(/\|/g, '¦');
                const expLabel = it.expandCollapseState === null ? '-' :
                    (it.expandCollapseState === 'error' ? 'err' : ExpStateLabel[it.expandCollapseState]);
                lines.push(`| ${it.depth} | ${nm} | ${it.controlType} | ${it.className} | ${expLabel} | ${it.isDialogOpener ? '✓' : ''} |`);
            }
            if (r.items.length > 60) lines.push(`| … | (${r.items.length - 60} more rows truncated) | | | | |`);
            lines.push('');
        }
    }
    return lines.join('\n');
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function run() {
    process.stderr.write('▶ Hwp attach\n');
    await controller.attach({ product: 'hwp' });
    await controller.setForeground();
    await sleep(400);
    for (let i = 0; i < 3; i++) {
        await controller.pressKeys({ keys: 'Escape' });
        await sleep(200);
    }

    const mapper = new MenuMapper({ product: 'hwp', probeDialogs: false });

    process.stderr.write('▶ Pre-sweep: 대상 항목 위치 수집\n');
    const { locations, notFound } = await prebuildLocations(mapper, TARGETS);
    for (const t of TARGETS) {
        const loc = locations[t];
        if (loc) {
            process.stderr.write(`  "${t}" → tab="${loc.tab}" (${loc.controlType})\n`);
        } else {
            process.stderr.write(`  "${t}" → NOT FOUND\n`);
        }
    }
    if (notFound.length > 0) {
        process.stderr.write(`⚠ 못 찾은 대상: ${notFound.join(', ')}\n`);
    }

    const results = [];
    for (const target of TARGETS) {
        process.stderr.write(`\n▶ 프로브: "${target}"\n`);
        let result;
        try {
            result = await probeWithTimeout(target, locations[target] || null);
        } catch (e) {
            result = {
                target,
                error: e.message,
                errorStack: (e.stack || '').split('\n').slice(0, 5).join('\n'),
                notes: ['timeout-or-fatal'],
                items: [],
            };
            // Hard recovery
            for (let i = 0; i < 5; i++) {
                try { await controller.pressKeys({ keys: 'Escape' }); } catch (_) {}
                await sleep(150);
            }
            try { await controller.setForeground(); } catch (_) {}
            await sleep(300);
        }
        results.push(result);
        process.stderr.write(
            `  tab=${result.tab || '-'} popupFound=${result.popupFound} ` +
            `items=${result.items?.length || 0} ms=${result.durationMs || 0}\n`,
        );
        await sleep(LIMITS.betweenProbeDelayMs);
    }

    // Write md
    const mdPath = path.join('docs', 'discovery-dropdown-results.md');
    try { fs.mkdirSync(path.dirname(mdPath), { recursive: true }); } catch (_) {}
    fs.writeFileSync(mdPath, renderMarkdown(results), 'utf8');
    process.stderr.write(`\n✅ ${mdPath}\n`);

    // JSON dump to stdout for further processing
    console.log(JSON.stringify(results, null, 2));
}

run().catch((e) => {
    console.error('FATAL:', e && e.message ? e.message : e);
    process.exit(1);
});
