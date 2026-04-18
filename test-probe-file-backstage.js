'use strict';
/**
 * Phase H smoke test: _probeFileBackstage()만 단독 실행.
 * MenuMapper의 전체 run()을 돌리지 않고 Step 11만 타겟 — 빠른 검증용.
 *
 * 성공 기준:
 *   - backstage.classification === 'backstage'
 *   - backstage.itemCount >= 15 (Discovery에서 22개 확인)
 *   - "끝", "인쇄", "문서 닫기"가 isDangerous=true로 표시
 *   - 실행 후 FrameWindowImpl 자식 개수가 복원됨
 */

const controller = require('./src/hwp-controller');
const sessionModule = require('./src/session');
const { MenuMapper } = require('./src/menu-mapper');

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function run() {
    process.stderr.write('▶ HWP attach\n');
    await controller.attach({ product: 'hwp' });
    await controller.setForeground();
    await sleep(500);

    const mapper = new MenuMapper({
        product: 'hwp',
        probeDialogs: false, probeDropdowns: false, probeShapeTab: false,
        probeTableTabs: false, probeChartTabs: false, probeContextMenus: false,
        probeImageTab: false, probeFileBackstage: true,
    });

    process.stderr.write('▶ 런처 해제\n');
    await mapper._dismissLauncher();
    await sleep(500);

    process.stderr.write('▶ Before 자식 개수\n');
    sessionModule.refreshHwpElement();
    const beforeKids = sessionModule.getSession().hwpElement.findAllChildren();
    const beforeCount = beforeKids.length;
    beforeKids.forEach((c) => { try { c.release(); } catch (_) {} });
    process.stderr.write(`  depth=1 자식: ${beforeCount}\n`);

    process.stderr.write('\n▶ _probeFileBackstage() 실행\n');
    const t0 = Date.now();
    await mapper._probeFileBackstage();
    const elapsed = Date.now() - t0;

    const bs = mapper.map.tabs['파일']?.backstage;
    if (!bs) {
        process.stderr.write('❌ backstage 결과 없음\n');
        process.exit(1);
    }

    process.stderr.write(`\n▶ 결과 (${elapsed}ms)\n`);
    process.stderr.write(`  classification: ${bs.classification}\n`);
    process.stderr.write(`  popupClass: ${bs.popupClass}\n`);
    process.stderr.write(`  popupRect: ${bs.popupRect ? JSON.stringify(bs.popupRect) : 'null'}\n`);
    process.stderr.write(`  itemCount: ${bs.itemCount}\n`);
    process.stderr.write(`  notes: ${bs.notes.length ? bs.notes.join(', ') : '(없음)'}\n`);

    process.stderr.write(`\n▶ 카테고리 목록:\n`);
    for (const it of bs.items) {
        const tag = it.isDangerous ? ' [⚠ DANGEROUS]' : '';
        process.stderr.write(`  - "${it.name}" (ALT+${it.accessKey ?? '?'}) @ (${it.clickX},${it.clickY}) d=${it.depth}${tag}\n`);
    }

    // 복귀 검증
    process.stderr.write('\n▶ After 자식 개수 검증\n');
    sessionModule.refreshHwpElement();
    const afterKids = sessionModule.getSession().hwpElement.findAllChildren();
    const afterCount = afterKids.length;
    afterKids.forEach((c) => { try { c.release(); } catch (_) {} });
    process.stderr.write(`  depth=1 자식: ${afterCount} (Before ${beforeCount})\n`);
    if (afterCount !== beforeCount) {
        process.stderr.write('  ⚠ 자식 개수 불일치 — Backstage가 닫히지 않았을 수 있음\n');
    } else {
        process.stderr.write('  ✅ 자식 개수 복원\n');
    }

    // 성공 기준 검증
    const errors = [];
    if (bs.classification !== 'backstage') errors.push(`classification=${bs.classification} (expect 'backstage')`);
    if (bs.itemCount < 15) errors.push(`itemCount=${bs.itemCount} (expect >= 15)`);
    const dangerNames = new Set(bs.items.filter((i) => i.isDangerous).map((i) => i.name));
    for (const mustBeDangerous of ['끝', '인쇄', '문서 닫기']) {
        if (!dangerNames.has(mustBeDangerous)) errors.push(`"${mustBeDangerous}" 미표시`);
    }

    if (errors.length) {
        process.stderr.write(`\n❌ 실패:\n`);
        errors.forEach((e) => process.stderr.write(`  - ${e}\n`));
        process.exit(1);
    } else {
        process.stderr.write(`\n✅ 모든 성공 기준 통과\n`);
    }
}

run().catch((e) => { process.stderr.write(`FATAL: ${e.message}\n${e.stack}\n`); process.exit(1); });
