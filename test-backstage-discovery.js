'use strict';
/**
 * Phase H-1 Discovery: 파일 Backstage 컨테이너/카테고리 식별.
 *
 * 전략:
 *   1. 파일 탭 클릭 전 FrameWindowImpl descendants 스냅샷 (depth 3)
 *   2. 파일 탭 클릭 + 대기
 *   3. 파일 탭 클릭 후 스냅샷 (depth 4 — Backstage 내부까지 탐색)
 *   4. Diff: 새 노드, 특히 depth=1의 새 자식(Backstage 후보) + 대형 rect 노드
 *   5. 카테고리 후보 수집 (Button/MenuItem/ListItem + 이름)
 *   6. Escape × 2로 복귀 + 자식 개수 정상화 확인
 *
 * 안전:
 *   - 파일 탭 1회만 클릭. 내부 카테고리는 절대 클릭 안 함.
 *   - Backstage 복귀 실패 시에도 Escape가 여러 번 호출되므로 최소한 리본 탭으로는 돌아감.
 *
 * 출력:
 *   - stderr: 사람이 읽을 수 있는 요약
 *   - maps/backstage-discovery.json: 전체 after 트리 + diff (Phase H-2 입력)
 */

const fs = require('fs');
const controller = require('./src/hwp-controller');
const sessionModule = require('./src/session');
const { MenuMapper } = require('./src/menu-mapper');

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function rectArea(r) {
    if (!r) return 0;
    return Math.max(0, r.right - r.left) * Math.max(0, r.bottom - r.top);
}

function rectStr(r) {
    if (!r) return '';
    return `${r.left},${r.top}-${r.right},${r.bottom}(${r.right - r.left}×${r.bottom - r.top})`;
}

/**
 * Frame root의 descendants를 재귀적으로 depth까지 수집.
 * COM 포인터 누수를 막기 위해 순회 중 release pool에 모아 마지막에 일괄 해제.
 */
function snapshotTree(root, maxDepth) {
    const out = [];
    const pool = [];
    function walk(el, depth, parentPath) {
        let info;
        try {
            info = {
                depth,
                parentPath,
                className: el.className || '',
                controlType: el.controlTypeName || '',
                name: ((el.name || '') + '').slice(0, 100),
                isEnabled: (() => { try { return el.isEnabled; } catch (_) { return null; } })(),
                rect: (() => { try { return el.boundingRect; } catch (_) { return null; } })(),
            };
        } catch (_) { return; }
        out.push(info);
        if (depth >= maxDepth) return;
        let children = [];
        try { children = el.findAllChildren(); } catch (_) { children = []; }
        const myPath = `${parentPath}/${info.className || info.controlType}:${info.name || ''}`;
        for (const c of children) {
            pool.push(c);
            walk(c, depth + 1, myPath);
        }
    }
    walk(root, 0, '');
    for (const p of pool) { try { p.release(); } catch (_) {} }
    return out;
}

function nodeKey(n) {
    return `${n.depth}|${n.className}|${n.controlType}|${n.name}|${rectStr(n.rect)}`;
}

function diffSnapshots(before, after) {
    const bSet = new Set(before.map(nodeKey));
    const aSet = new Set(after.map(nodeKey));
    return {
        added: after.filter((n) => !bSet.has(nodeKey(n))),
        removed: before.filter((n) => !aSet.has(nodeKey(n))),
    };
}

async function run() {
    process.stderr.write('▶ HWP attach + 정리\n');
    await controller.attach({ product: 'hwp' });
    await controller.setForeground();
    await sleep(500);

    const mapper = new MenuMapper({ product: 'hwp', probeDialogs: false, probeDropdowns: false });

    // 런처(시작 화면) 해제 — Escape로는 닫히지 않고 "새 문서" ListItem 더블클릭 필요
    process.stderr.write('▶ 런처 해제 시도\n');
    await mapper._dismissLauncher();
    await sleep(500);
    for (let i = 0; i < 2; i++) { await controller.pressKeys({ keys: 'Escape' }); await sleep(200); }

    // 파일 탭 좌표
    process.stderr.write('▶ 메뉴 탭 수집\n');
    const tabs = await mapper._collectMenuTabs();
    const fileTab = tabs.find((t) => t.name === '파일');
    if (!fileTab) throw new Error('파일 탭을 찾을 수 없음');
    process.stderr.write(`  파일 탭: name="${fileTab.uiaName}" click=(${fileTab.clickX},${fileTab.clickY}) enabled=${fileTab.isEnabled}\n`);

    // Before snapshot (depth 3)
    process.stderr.write('▶ Before snapshot (depth 3)\n');
    sessionModule.refreshHwpElement();
    const frameBefore = sessionModule.getSession().hwpElement;
    if (!frameBefore) throw new Error('Frame before refresh failed');
    const frameRectBefore = (() => { try { return frameBefore.boundingRect; } catch (_) { return null; } })();
    const before = snapshotTree(frameBefore, 3);
    const beforeTop = before.filter((n) => n.depth === 1);
    process.stderr.write(`  노드 총 ${before.length}, depth=1 자식 ${beforeTop.length}\n`);
    for (const c of beforeTop) {
        process.stderr.write(`    · ${c.className}|${c.controlType}|${c.name} ${rectStr(c.rect)}\n`);
    }

    // 파일 탭 클릭
    process.stderr.write('▶ 파일 탭 클릭\n');
    await mapper._switchTab(fileTab);
    await sleep(1200); // Backstage 전환 여유

    // After snapshot (depth 4 — Backstage 내부까지)
    process.stderr.write('▶ After snapshot (depth 4)\n');
    sessionModule.refreshHwpElement();
    const frameAfter = sessionModule.getSession().hwpElement;
    if (!frameAfter) throw new Error('Frame after refresh failed');
    const frameRectAfter = (() => { try { return frameAfter.boundingRect; } catch (_) { return null; } })();
    const after = snapshotTree(frameAfter, 4);
    const afterTop = after.filter((n) => n.depth === 1);
    process.stderr.write(`  노드 총 ${after.length}, depth=1 자식 ${afterTop.length}\n`);
    for (const c of afterTop) {
        process.stderr.write(`    · ${c.className}|${c.controlType}|${c.name} ${rectStr(c.rect)}\n`);
    }

    // Diff
    const d = diffSnapshots(before, after);
    process.stderr.write(`\n▶ Diff: +${d.added.length}  -${d.removed.length}\n`);

    // 새 top-level
    const newTop = d.added.filter((n) => n.depth === 1);
    if (newTop.length > 0) {
        process.stderr.write(`\n★ 새 top-level (depth=1) 자식: ${newTop.length}개\n`);
        for (const n of newTop) {
            process.stderr.write(`  ${n.className}|${n.controlType}|${n.name} ${rectStr(n.rect)}\n`);
        }
    } else {
        process.stderr.write(`\n⚠ 새 top-level 없음 — Backstage가 기존 컨테이너 내부를 교체했을 가능성\n`);
    }

    // 대형 새 노드 (frame 30% 이상)
    const frameArea = rectArea(frameRectAfter);
    const bigNew = d.added
        .filter((n) => n.rect && frameArea > 0 && rectArea(n.rect) > frameArea * 0.3)
        .sort((a, b) => rectArea(b.rect) - rectArea(a.rect));
    if (bigNew.length > 0) {
        process.stderr.write(`\n★ 대형 새 노드 (frame의 30% 이상): ${bigNew.length}개\n`);
        for (const n of bigNew.slice(0, 15)) {
            process.stderr.write(`  d=${n.depth} ${n.className}|${n.controlType}|${n.name} ${rectStr(n.rect)}\n`);
        }
    }

    // 카테고리 후보 (Button/MenuItem/ListItem/Text + 이름 있음, 작은 버튼 크기 상정)
    const catCandidates = d.added
        .filter((n) => n.name && n.name.length > 0 && n.name.length < 30)
        .filter((n) => ['Button', 'MenuItem', 'ListItem', 'Text', 'TreeItem', 'TabItem'].includes(n.controlType))
        .filter((n) => n.rect && rectArea(n.rect) > 0 && rectArea(n.rect) < frameArea * 0.2);

    if (catCandidates.length > 0) {
        process.stderr.write(`\n★ 카테고리 후보: ${catCandidates.length}개\n`);
        for (const n of catCandidates.slice(0, 50)) {
            process.stderr.write(`  d=${n.depth} ${n.controlType}|"${n.name}" ${rectStr(n.rect)}\n`);
        }
    }

    // 전체 덤프 저장
    if (!fs.existsSync('maps')) fs.mkdirSync('maps');
    fs.writeFileSync('maps/backstage-discovery.json', JSON.stringify({
        generatedAt: new Date().toISOString(),
        fileTab,
        frameRect: { before: frameRectBefore, after: frameRectAfter },
        before: { nodes: before.length, topChildren: beforeTop },
        after: { nodes: after.length, topChildren: afterTop, fullTree: after },
        diff: { added: d.added, removed: d.removed },
        categoryCandidates: catCandidates,
        bigNewNodes: bigNew,
    }, null, 2));
    process.stderr.write(`\n▶ maps/backstage-discovery.json 저장\n`);

    // 정리: Escape × 2 (계획서에 "Backstage 종료는 Escape 1회로 가능" 기술, 안전 margin 2회)
    process.stderr.write('\n▶ 복귀 (Escape × 2)\n');
    await controller.setForeground();
    await sleep(300);
    for (let i = 0; i < 2; i++) { await controller.pressKeys({ keys: 'Escape' }); await sleep(500); }

    // 복구 검증
    sessionModule.refreshHwpElement();
    const frameFinal = sessionModule.getSession().hwpElement;
    const finalTopSnap = snapshotTree(frameFinal, 1).filter((n) => n.depth === 1);
    process.stderr.write(`▶ 복귀 후 depth=1 자식 ${finalTopSnap.length} (Before ${beforeTop.length})\n`);
    process.stderr.write(finalTopSnap.length === beforeTop.length ? '  ✅ 자식 개수 복원\n' : '  ⚠ 자식 개수 불일치\n');
}

run().catch((e) => {
    process.stderr.write(`FATAL: ${e.message}\n${e.stack || ''}\n`);
    process.exit(1);
});
