'use strict';
/** 서식 탭만 프로빙 (편집 탭 프로빙 없이) - 오염 원인 탐지 */
const { MenuMapper } = require('./src/menu-mapper');
const controller = require('./src/hwp-controller');

async function run() {
    await controller.attach({ product: 'hwp' });
    await controller.setForeground();
    for (let i = 0; i < 3; i++) {
        await controller.pressKeys({ keys: 'Escape' });
        await new Promise(r => setTimeout(r, 200));
    }

    const mapper = new MenuMapper({ product: 'hwp', probeDialogs: true });
    mapper._probeAllDialogs = async function () {
        const menuTabs = await this._collectMenuTabs();
        const tabLookup = {};
        for (const t of menuTabs) tabLookup[t.name] = t;

        const tabName = '서식';
        const tabData = this.map.tabs[tabName];
        const tabInfo = tabLookup[tabName];
        if (!tabData || !tabInfo) throw new Error('서식 탭 정보 없음');

        process.stderr.write(`\n▶ 서식 탭 프로빙 시작 (항목 ${tabData.ribbonItems.length}개)\n`);

        let cnt = 0;
        for (const item of tabData.ribbonItems) {
            cnt++;
            if (cnt > 5) break; // 처음 5개만 테스트
            if (item.hasDropdown) { item.type = 'dropdown'; process.stderr.write(`  ▼ "${item.name}" → dropdown\n`); continue; }
            try {
                await this._switchTab(tabInfo);
                const dialog = await this._probeItem(item);
                if (dialog) {
                    item.type = 'dialog';
                    item.dialog = dialog;
                    process.stderr.write(`  💬 "${item.name}" → ${dialog.title}\n`);
                } else {
                    item.type = 'action';
                    process.stderr.write(`  ⚡ "${item.name}" → action (다이얼로그 미감지)\n`);
                }
            } catch (e) {
                process.stderr.write(`  ⚠ "${item.name}" 실패: ${e.message}\n`);
                item.type = 'error';
            }
            await this._ensureNoDialog();
        }
    };

    await mapper.run();
}

run().catch(e => { console.error('ERROR:', e.message); process.exit(1); });
