'use strict';
/**
 * Phase F smoke test: 이미지를 클립보드로 붙여넣기 → 새 컨텍스트 탭 감지.
 */
const { execSync } = require('child_process');
const controller = require('./src/hwp-controller');
const sessionModule = require('./src/session');
const win32 = require('./src/win32');
const { MenuMapper } = require('./src/menu-mapper');

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/**
 * PowerShell로 작은 PNG 이미지를 클립보드에 올림.
 * STA apartment가 필요하므로 -Sta 플래그 필수.
 */
function putImageOnClipboard() {
    const script = `
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$bmp = New-Object System.Drawing.Bitmap 80,60
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.FillRectangle([System.Drawing.Brushes]::CornflowerBlue, 0, 0, 80, 60)
$g.Dispose()
[System.Windows.Forms.Clipboard]::SetImage($bmp)
Start-Sleep -Milliseconds 300
`;
    const encoded = Buffer.from(script, 'utf16le').toString('base64');
    execSync(`powershell -Sta -NoProfile -EncodedCommand ${encoded}`, {
        timeout: 10000,
        windowsHide: true,
        stdio: 'ignore',
    });
}

async function run() {
    await controller.attach({ product: 'hwp' });
    await controller.setForeground();
    await sleep(500);
    for (let i = 0; i < 3; i++) { await controller.pressKeys({ keys: 'Escape' }); await sleep(200); }

    await controller.pressKeys({ keys: 'Ctrl+End' });
    await sleep(300);

    const mapper = new MenuMapper({ product: 'hwp', probeDialogs: false, probeDropdowns: false, probeShapeTab: false, probeTableTabs: false, probeChartTabs: false, probeContextMenus: false });
    const initialTabs = await mapper._collectMenuTabs();
    const initialEnabled = new Set(initialTabs.filter(t => t.isEnabled).map(t => t.name));
    process.stderr.write(`초기 활성 탭: ${Array.from(initialEnabled).join(', ')}\n`);

    // Image clipboard
    process.stderr.write('▶ PowerShell로 이미지 클립보드 세팅\n');
    try { putImageOnClipboard(); } catch (e) { process.stderr.write(`⚠ clipboard 실패: ${e.message}\n`); throw e; }
    process.stderr.write('▶ Ctrl+V 붙여넣기\n');
    await controller.pressKeys({ keys: 'Ctrl+V' });
    await sleep(2000);

    // 새 탭 감지
    const afterTabs = await mapper._collectMenuTabs();
    const afterEnabled = new Set(afterTabs.filter(t => t.isEnabled).map(t => t.name));
    const newTabs = Array.from(afterEnabled).filter(n => !initialEnabled.has(n));
    process.stderr.write(`붙여넣기 후 활성 탭: ${Array.from(afterEnabled).join(', ')}\n`);
    process.stderr.write(`새 탭: ${newTabs.length > 0 ? newTabs.join(', ') : '없음'}\n`);

    if (newTabs.length > 0) {
        for (const tabName of newTabs) {
            const tabInfo = afterTabs.find(t => t.name === tabName);
            await mapper._switchTab(tabInfo);
            await sleep(500);
            const items = await mapper._collectRibbonItems(tabInfo);
            process.stderr.write(`  "${tabName}" 리본: ${items.length}개\n`);
            for (const it of items.slice(0, 10)) {
                process.stderr.write(`    - ${it.name} (${it.controlType}${it.hasDropdown ? ', dropdown' : ''})\n`);
            }
        }
    }

    // 정리
    process.stderr.write('▶ 정리 (Escape + Ctrl+Z × 5)\n');
    await controller.setForeground();
    for (let i = 0; i < 3; i++) { await controller.pressKeys({ keys: 'Escape' }); await sleep(200); }
    for (let i = 0; i < 5; i++) { await controller.pressKeys({ keys: 'Ctrl+Z' }); await sleep(300); }

    const finalTabs = await mapper._collectMenuTabs();
    const finalEnabled = new Set(finalTabs.filter(t => t.isEnabled).map(t => t.name));
    const remaining = Array.from(finalEnabled).filter(n => !initialEnabled.has(n));
    process.stderr.write(`정리 후 남은 새 탭: ${remaining.length > 0 ? remaining.join(', ') : '없음 ✅'}\n`);
}

run().catch((e) => { process.stderr.write(`FATAL: ${e.message}\n${e.stack}\n`); process.exit(1); });
