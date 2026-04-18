'use strict';
/**
 * Phase A-2 단독 테스트: 도형 탭만 프로빙.
 * 전체 run 없이 Step 6만 실행하여 빠른 검증.
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
        probeDialogs: false,
        probeDropdowns: false,
        probeShapeTab: true,
    });

    // 최소 Step 1,2만 수행 (Step 6이 기대하는 map.tabs['편집'] 준비)
    mapper.map.mappedAt = new Date().toISOString();
    await controller.pressKeys({ keys: 'Escape' });
    await new Promise(r => setTimeout(r, 300));

    const menuTabs = await mapper._collectMenuTabs();
    for (const tab of menuTabs) {
        if (tab.name === '파일' || !tab.isEnabled) {
            mapper.map.tabs[tab.name] = { accessKey: tab.accessKey, uiaName: tab.uiaName, ribbonItems: [], note: tab.isEnabled ? 'backstage' : 'disabled' };
            continue;
        }
        if (tab.name !== '편집') {
            // 다른 탭은 리본 수집 건너뜀 (Step 6은 편집 탭만 필요)
            mapper.map.tabs[tab.name] = { accessKey: tab.accessKey, uiaName: tab.uiaName, ribbonItems: [], note: 'skipped-for-test' };
            continue;
        }
        const items = await mapper._collectRibbonItems(tab);
        for (const it of items) it.type = it.hasDropdown ? 'dropdown' : 'action';
        mapper.map.tabs[tab.name] = { accessKey: tab.accessKey, uiaName: tab.uiaName, ribbonItems: items };
        mapper.map.stats.totalTabs++;
        mapper.map.stats.totalRibbonItems += items.length;
    }

    // Step 6 단독 실행
    process.stderr.write('\n▶ Step 6: 도형 탭 프로빙 시작\n');
    await mapper._probeShapeTab();

    // 결과 요약
    process.stderr.write('\n=== 도형 탭 프로빙 결과 ===\n');
    const shapeData = mapper.map.tabs['도형'];
    if (!shapeData || !shapeData.ribbonItems || shapeData.ribbonItems.length === 0) {
        process.stderr.write('❌ 도형 탭 데이터 없음\n');
        process.exit(1);
    }
    const byClass = {};
    const dropdowns = shapeData.ribbonItems.filter(i => i.dropdown);
    for (const it of dropdowns) {
        const c = it.dropdown.classification;
        byClass[c] = (byClass[c] || 0) + 1;
    }
    process.stderr.write(`리본 항목: ${shapeData.ribbonItems.length}개\n`);
    process.stderr.write(`드롭다운: ${dropdowns.length}개\n`);
    process.stderr.write(`분류별: ${JSON.stringify(byClass)}\n`);
    process.stderr.write(`stats: ${JSON.stringify(mapper.map.stats)}\n`);

    process.stderr.write('\n== 드롭다운 항목 ==\n');
    for (const it of dropdowns) {
        const dd = it.dropdown;
        process.stderr.write(`  - "${it.name}" → ${dd.classification} (${dd.itemCount}개, ${dd.probeDurationMs || 0}ms)${dd.notes?.length ? ' notes=' + dd.notes.join(',') : ''}\n`);
    }

    console.log(JSON.stringify(shapeData, null, 2));
}

run().catch(e => {
    process.stderr.write(`FATAL: ${e.message}\n${e.stack}\n`);
    process.exit(1);
});
