'use strict';
/**
 * 편집 탭 드롭다운만 프로빙하는 단일 탭 테스트.
 * Phase 4 menu-mapper 통합 후 Phase 5 전체 재실행 전 smoke test.
 */
const { MenuMapper } = require('./src/menu-mapper');
const controller = require('./src/hwp-controller');

async function run() {
    await controller.attach({ product: 'hwp' });
    await controller.setForeground();
    for (let i = 0; i < 3; i++) {
        await controller.pressKeys({ keys: 'Escape' });
        await new Promise(r => setTimeout(r, 200));
    }

    const mapper = new MenuMapper({
        product: 'hwp',
        probeDialogs: false,   // 다이얼로그 프로빙 생략
        probeDropdowns: true,
    });

    // Step 1,2,3 수행하되 다이얼로그 생략. Step 5를 편집 탭으로 제한.
    mapper.map.mappedAt = new Date().toISOString();
    await controller.attach({ product: 'hwp' });
    await controller.setForeground();
    await new Promise(r => setTimeout(r, 400));
    await controller.pressKeys({ keys: 'Escape' });
    await new Promise(r => setTimeout(r, 300));

    // 메뉴 탭 + 리본 항목 수집 (편집 탭 필요)
    const menuTabs = await mapper._collectMenuTabs();
    for (const tab of menuTabs) {
        if (tab.name !== '편집') {
            mapper.map.tabs[tab.name] = { accessKey: tab.accessKey, uiaName: tab.uiaName, ribbonItems: [], note: 'skipped-for-test' };
            continue;
        }
        const ribbonItems = await mapper._collectRibbonItems(tab);
        for (const it of ribbonItems) it.type = it.hasDropdown ? 'dropdown' : 'action';
        mapper.map.tabs[tab.name] = {
            accessKey: tab.accessKey,
            uiaName: tab.uiaName,
            ribbonItems,
        };
        mapper.map.stats.totalTabs++;
        mapper.map.stats.totalRibbonItems += ribbonItems.length;
        process.stderr.write(`▶ 편집 탭: ${ribbonItems.length}개 항목 (드롭다운 ${ribbonItems.filter(i => i.hasDropdown).length}개)\n`);
    }

    // Step 5
    process.stderr.write('\n▶ Step 5: 드롭다운 프로빙 시작\n');
    await mapper._probeAllDropdowns();

    // 결과 요약
    process.stderr.write('\n=== 편집 탭 드롭다운 프로빙 결과 ===\n');
    const editItems = mapper.map.tabs['편집'].ribbonItems.filter(i => i.hasDropdown);
    const byClass = {};
    for (const item of editItems) {
        const c = item.dropdown?.classification || 'n/a';
        byClass[c] = (byClass[c] || 0) + 1;
    }
    process.stderr.write('분류별 집계: ' + JSON.stringify(byClass) + '\n');
    process.stderr.write('stats: ' + JSON.stringify(mapper.map.stats) + '\n');

    for (const item of editItems) {
        const dd = item.dropdown;
        if (!dd) continue;
        process.stderr.write(`  - "${item.name}" → ${dd.classification} (${dd.itemCount}개, ${dd.probeDurationMs}ms)${dd.notes?.length ? ' notes=' + dd.notes.join(',') : ''}\n`);
    }

    // JSON dump
    console.log(JSON.stringify(mapper.map.tabs['편집'], null, 2));
}

run().catch(e => { console.error('FATAL:', e.message, '\n', e.stack); process.exit(1); });
