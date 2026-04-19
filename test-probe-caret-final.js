'use strict';
/**
 * 최종 검증: 프로덕션 `_probeMemoTab / _probeAnnotationTab / _probeHeaderFooterTab`을
 * 순차 실행하여 map.tabs에 contextual tab이 정상 기록되는지 확인.
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
        probeShapeTab: false,
        probeTableTabs: false,
        probeChartTabs: false,
        probeImageTab: false,
        probeContextMenus: false,
        probeFileBackstage: false,
    });
    mapper.map.mappedAt = new Date().toISOString();

    await controller.pressKeys({ keys: 'Escape' });
    await new Promise(r => setTimeout(r, 300));

    // 입력 / 쪽 탭 리본만 사전 수집 (각 probe가 map.tabs[...]에서 item을 찾음)
    const menuTabs = await mapper._collectMenuTabs();
    for (const tab of menuTabs) {
        if (tab.name === '파일' || !tab.isEnabled) {
            mapper.map.tabs[tab.name] = { accessKey: tab.accessKey, uiaName: tab.uiaName, ribbonItems: [], note: tab.isEnabled ? 'backstage' : 'disabled' };
            continue;
        }
        if (tab.name !== '입력' && tab.name !== '쪽') {
            mapper.map.tabs[tab.name] = { accessKey: tab.accessKey, uiaName: tab.uiaName, ribbonItems: [], note: 'skipped' };
            continue;
        }
        const items = await mapper._collectRibbonItems(tab);
        for (const it of items) it.type = it.hasDropdown ? 'dropdown' : 'action';
        mapper.map.tabs[tab.name] = { accessKey: tab.accessKey, uiaName: tab.uiaName, ribbonItems: items };
    }

    // 프로덕션 probe 3개 실행
    for (const [name, fn] of [
        ['memo', () => mapper._probeMemoTab()],
        ['annotation', () => mapper._probeAnnotationTab()],
        ['headerFooter', () => mapper._probeHeaderFooterTab()],
    ]) {
        process.stderr.write(`\n▶ ${name}\n`);
        try { await fn(); }
        catch (e) { process.stderr.write(`  ✗ ${name} 예외: ${e.message}\n`); }
        try { await mapper._recover(); } catch (_) {}
        await new Promise(r => setTimeout(r, 500));
    }

    // 결과 요약
    process.stderr.write('\n=== 결과 ===\n');
    const contextTabs = Object.entries(mapper.map.tabs)
        .filter(([, v]) => v.contextState)
        .map(([k, v]) => ({ name: k, contextState: v.contextState, items: v.ribbonItems?.length || 0 }));
    process.stderr.write(`Contextual tabs 생성: ${contextTabs.length}개\n`);
    for (const t of contextTabs) {
        process.stderr.write(`  "${t.name}" contextState=${t.contextState} ribbonItems=${t.items}\n`);
    }
    process.stderr.write(`stats: ${JSON.stringify(mapper.map.stats)}\n`);

    console.log(JSON.stringify({
        contextualTabs: contextTabs,
        memoTab: mapper.map.tabs['메모'] || null,
        annotationTab: mapper.map.tabs['주석'] || null,
        headerFooterTab: mapper.map.tabs['머리말/꼬리말'] || null,
    }, null, 2));
}

run().catch(e => {
    process.stderr.write(`FATAL: ${e.message}\n${e.stack}\n`);
    process.exit(1);
});
