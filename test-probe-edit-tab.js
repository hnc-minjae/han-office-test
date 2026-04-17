'use strict';
/** 편집 탭만 프로빙하는 단일 탭 테스트 */
const { MenuMapper } = require('./src/menu-mapper');
const controller = require('./src/hwp-controller');

async function run() {
    // 안전하게 열린 다이얼로그 정리
    await controller.attach({ product: 'hwp' });
    await controller.setForeground();
    for (let i = 0; i < 3; i++) {
        await controller.pressKeys({ keys: 'Escape' });
        await new Promise(r => setTimeout(r, 200));
    }

    const mapper = new MenuMapper({ product: 'hwp', probeDialogs: true });
    // 편집 탭만 프로빙하도록 _probeAllDialogs를 재정의
    const origProbe = mapper._probeAllDialogs.bind(mapper);
    mapper._probeAllDialogs = async function () {
        const menuTabs = await this._collectMenuTabs();
        const tabLookup = {};
        for (const t of menuTabs) tabLookup[t.name] = t;

        const tabName = '편집';
        const tabData = this.map.tabs[tabName];
        const tabInfo = tabLookup[tabName];
        if (!tabData || !tabInfo) throw new Error('편집 탭 정보 없음');

        process.stderr.write(`\n▶ 편집 탭 프로빙 시작 (항목 ${tabData.ribbonItems.length}개)\n`);

        for (const item of tabData.ribbonItems) {
            if (item.hasDropdown) { item.type = 'dropdown'; continue; }
            try {
                await this._switchTab(tabInfo);
                const dialog = await this._probeItem(item);
                if (dialog) {
                    item.type = 'dialog';
                    item.dialog = dialog;
                    this.map.stats.totalDialogs++;
                    const ctrlCount = Object.values(dialog.controls).flat().length;
                    this.map.stats.totalControls += ctrlCount;
                    process.stderr.write(`  💬 "${item.name}" → ${dialog.title} (${dialog.tabs.length}탭, ${ctrlCount}컨트롤)\n`);
                } else {
                    item.type = 'action';
                    process.stderr.write(`  ⚡ "${item.name}" → action\n`);
                }
            } catch (e) {
                process.stderr.write(`  ⚠ "${item.name}" 실패: ${e.message}\n`);
                item.type = 'error';
                this.map.stats.errors++;
                await this._recover();
            }
            await this._ensureNoDialog();
        }
    };

    await mapper.run();

    // 결과 출력
    const editData = mapper.map.tabs['편집'];
    console.log('\n=== 편집 탭 프로빙 결과 ===');
    const byType = { action: 0, dialog: 0, dropdown: 0, error: 0, unknown: 0 };
    for (const item of editData.ribbonItems) {
        byType[item.type] = (byType[item.type] || 0) + 1;
    }
    console.log('항목 유형별 집계:', JSON.stringify(byType));

    console.log('\n다이얼로그 감지된 항목:');
    for (const item of editData.ribbonItems.filter(i => i.type === 'dialog')) {
        console.log(`  - "${item.name}" → "${item.dialog.title}"`);
        console.log(`    탭: ${item.dialog.tabs.join(', ')}`);
        for (const [tabName, ctrls] of Object.entries(item.dialog.controls)) {
            console.log(`    [${tabName}] ${ctrls.length}개 컨트롤`);
        }
    }
}

run().catch(e => { console.error('ERROR:', e.message); process.exit(1); });
