#!/usr/bin/env node
/**
 * verify-edit-modes — DEPRECATED.
 * 편집 모드 리본 프로빙(_probeMemoEditMode 등)은 제거되었고, 대신 contextual tab
 * 프로버(_probeMemoTab/_probeAnnotationTab/_probeHeaderFooterTab)로 대체됨.
 * 이 스크립트는 현재 참조하는 메서드들이 존재하지 않아 실행 불가.
 * 신규 구조 검증은 `node src/menu-mapper.js`를 실행해 Step 9.5~9.7 로그를 확인할 것.
 */
'use strict';

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const controller = require('./hwp-controller');
const sessionModule = require('./session');
const win32 = require('./win32');
const { MenuMapper } = require('./menu-mapper');

const sleep = (ms) => new Promise(r => setTimeout(r, ms));
const log = (msg) => process.stderr.write(`${msg}\n`);

function listHwpPids() {
    try {
        const out = execSync('tasklist /FI "IMAGENAME eq Hwp.exe" /FO CSV /NH',
            { encoding: 'utf8', windowsHide: true });
        const pids = [];
        for (const line of out.split('\n')) {
            const m = /^"Hwp\.exe","(\d+)"/.exec(line);
            if (m) pids.push(parseInt(m[1], 10));
        }
        return pids;
    } catch (_) { return []; }
}

async function main() {
    const mapPath = path.resolve('maps/hwp-menu-map.json');
    const existing = JSON.parse(fs.readFileSync(mapPath, 'utf8'));
    log(`▶ 기존 map 로드 (${Object.keys(existing.tabs).length}개 탭)`);

    // HWP launch via controller (winax 없어도 됨)
    const beforePids = listHwpPids();
    await controller.launch({ product: 'hwp', timeoutMs: 30000 });
    const afterPids = listHwpPids();
    const pid = afterPids.find(p => !beforePids.includes(p)) || afterPids[0];
    log(`✓ HWP launched, PID=${pid}`);
    await controller.setForeground();
    await sleep(400);
    const hwnd = sessionModule.getSession().hwpProcess.hwnd;
    if (hwnd) win32.maximizeWindow(hwnd);
    await sleep(600);

    // MenuMapper 인스턴스에 기존 map 주입 (probe 메서드가 this.map을 참조하므로)
    const mapper = new MenuMapper({ product: 'hwp' });
    mapper.map = existing;

    const results = {};

    log('\n=== Phase I: 메모 편집 모드 ===');
    try {
        results.memo = await mapper._probeMemoEditMode();
        log(`  entered=${results.memo.entered}, items=${results.memo.ribbonItems?.length || 0}, exit="${results.memo.exitButton?.name || '-'}"`);
    } catch (e) { log(`  ERR: ${e.message}`); }

    log('\n=== Phase J1: 머리말 편집 모드 ===');
    try {
        results.header = await mapper._probeHeaderFooterEditMode('header');
        log(`  entered=${results.header.entered}, items=${results.header.ribbonItems?.length || 0}, exit="${results.header.exitButton?.name || '-'}"`);
    } catch (e) { log(`  ERR: ${e.message}`); }

    log('\n=== Phase J2: 꼬리말 편집 모드 ===');
    try {
        results.footer = await mapper._probeHeaderFooterEditMode('footer');
        log(`  entered=${results.footer.entered}, items=${results.footer.ribbonItems?.length || 0}, exit="${results.footer.exitButton?.name || '-'}"`);
    } catch (e) { log(`  ERR: ${e.message}`); }

    log('\n=== Phase K1: 각주 편집 모드 ===');
    try {
        results.footnote = await mapper._probeFootnoteEditMode();
        log(`  entered=${results.footnote.entered}, items=${results.footnote.ribbonItems?.length || 0}, exit="${results.footnote.exitButton?.name || '-'}"`);
    } catch (e) { log(`  ERR: ${e.message}`); }

    log('\n=== Phase K2: 미주 편집 모드 ===');
    try {
        results.endnote = await mapper._probeEndnoteEditMode();
        log(`  entered=${results.endnote.entered}, items=${results.endnote.ribbonItems?.length || 0}, exit="${results.endnote.exitButton?.name || '-'}"`);
    } catch (e) { log(`  ERR: ${e.message}`); }

    const outPath = 'test-results/edit-modes-probe.json';
    fs.mkdirSync(path.dirname(outPath), { recursive: true });
    fs.writeFileSync(outPath, JSON.stringify(results, null, 2), 'utf8');
    log(`\n✓ 결과 저장: ${outPath}`);
}

main().catch(e => { log('✗ ' + e.message + '\n' + (e.stack || '')); process.exit(1); });
