'use strict';

/**
 * 다이얼로그 컨트롤 조작 종합 테스트
 *  - ComboBox (기본 탭의 기준 크기) — test-dialog.js에서 검증 완료 ✅
 *  - Edit (확장 탭의 X 방향 SpinImpl) — UIA 조작 불가 ⚠
 *  - CheckBox (확장 탭의 커닝 CheckBoxImpl) — 클릭만 가능, 상태 검증 불가 ⚠
 *  - TabItem 전환 (기본 → 확장) — 검증 ✅
 *
 * 발견된 한컴 UIA 제약:
 *  - SpinImpl: UIA 트리에 자식 노출 안 됨, SetFocus·키 입력이 Edit 내부로 전달 안 됨
 *  - CheckBoxImpl: TogglePattern 미구현 상태로 IsChecked 읽기 불가
 *  - ComboImpl: 정상 동작 (값을 name 속성에 포함하여 공개)
 *
 * 결론: menu-mapper 구조 수집 용도에는 충분. 몽키 테스트/자동화에서 값 입력이
 * 필요한 경우는 WM_SETTEXT 직접 송신, IAccessible API 등 추가 연구 필요.
 */

const controller = require('./src/hwp-controller');
const sessionModule = require('./src/session');
const { TreeScope } = require('./src/uia');
const win32 = require('./src/win32');

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function findDialog(session) {
    sessionModule.refreshHwpElement();
    const topCh = session.hwpElement.findAllChildren();
    let dialog = null;
    for (const c of topCh) {
        try { if (c.className === 'DialogImpl') { dialog = c; break; } } catch (_) {}
    }
    topCh.filter(c => c !== dialog).forEach(c => { try { c.release(); } catch (_) {} });
    return dialog;
}

function findRectInDialog(dialog, predicate) {
    const descs = dialog.findAll(TreeScope.Descendants);
    let rect = null, name = null;
    for (const el of descs) {
        try {
            if (predicate(el)) {
                rect = el.boundingRect;
                name = el.name;
                break;
            }
        } catch (_) {}
    }
    descs.forEach(el => { try { el.release(); } catch (_) {} });
    return rect ? { rect, name } : null;
}

function centerOf(rect) {
    return { x: Math.round((rect.left + rect.right) / 2), y: Math.round((rect.top + rect.bottom) / 2) };
}

async function openCharFormatDialog(session) {
    sessionModule.refreshHwpElement();
    const ch = session.hwpElement.findAllChildren();
    const mb = ch.find(c => { try { return c.className === 'MenuBarImpl'; } catch (_) { return false; } });
    ch.filter(c => c !== mb).forEach(c => c.release());
    const mbc = mb.findAllChildren();
    const et = mbc.find(el => { try { return el.name === '편집 E'; } catch (_) { return false; } });
    const tr = et.boundingRect;
    win32.mouseClick(tr.left + Math.round((tr.right - tr.left) * 0.3), Math.round((tr.top + tr.bottom) / 2));
    mbc.forEach(el => el.release());
    mb.release();
    await sleep(600);

    sessionModule.refreshHwpElement();
    const d = session.hwpElement.findAll(TreeScope.Descendants);
    let r = null;
    for (const el of d) {
        try {
            if (el.name && el.name.startsWith('글자 모양') && el.controlTypeName === 'Button') {
                r = el.boundingRect;
                break;
            }
        } catch (_) {}
    }
    d.forEach(el => { try { el.release(); } catch (_) {} });
    win32.mouseClick(Math.round((r.left + r.right) / 2), r.top + Math.round((r.bottom - r.top) * 0.3));
    await sleep(1500);
}

async function switchDialogTab(session, tabName) {
    const dialog = findDialog(session);
    const tab = findRectInDialog(dialog,
        el => el.controlTypeName === 'TabItem' && el.name === tabName);
    dialog.release();
    if (!tab) throw new Error(`탭 "${tabName}" 못찾음`);
    const c = centerOf(tab.rect);
    win32.mouseClick(c.x, c.y);
    await sleep(500);
}

async function readDialogControl(session, predicate) {
    const dialog = findDialog(session);
    const match = findRectInDialog(dialog, predicate);
    dialog.release();
    return match;
}

async function clickDialogButton(session, namePrefix) {
    const dialog = findDialog(session);
    const btn = findRectInDialog(dialog,
        el => el.name && el.name.startsWith(namePrefix) && el.className === 'DialogButtonImpl');
    dialog.release();
    if (!btn) throw new Error(`버튼 "${namePrefix}" 못찾음`);
    const c = centerOf(btn.rect);
    win32.mouseClick(c.x, c.y);
    await sleep(800);
}

async function run() {
    await controller.attach({ product: 'hwp' });
    await controller.setForeground();
    await sleep(300);
    await controller.pressKeys({ keys: 'Escape' });
    await sleep(400);

    const session = sessionModule.getSession();

    console.log('=== 준비: 테스트 문자 입력 + 전체 선택 ===');
    win32.clipboardPaste('가나다');
    await sleep(300);
    await controller.pressKeys({ keys: 'Ctrl+A' });
    await sleep(300);

    console.log('\n=== 다이얼로그 열기 ===');
    await openCharFormatDialog(session);
    let dialog = findDialog(session);
    if (!dialog) { console.log('다이얼로그 없음'); process.exit(1); }
    dialog.release();

    console.log('\n=== 확장 탭으로 전환 ===');
    await switchDialogTab(session, '확장');

    // ============================================================
    // Test 1: Edit (SpinImpl) - X 방향 10 → 30
    // ============================================================
    console.log('\n=== Test 1: Edit(SpinImpl) - X 방향 10 → 30 ===');
    const xBefore = await readDialogControl(session,
        el => el.controlTypeName === 'Edit' && el.name.startsWith('X 방향'));
    console.log('  변경 전: "' + xBefore.name + '"');

    // Alt+X 액세스 키로 포커스 + 자동 전체 선택 (SpinImpl의 표준 동작)
    await controller.pressKeys({ keys: 'Alt+X' });
    await sleep(300);
    win32.typeAsciiText('30');
    await sleep(300);

    // ============================================================
    // Test 2: CheckBox - "커닝" 토글
    // ============================================================
    console.log('\n=== Test 2: CheckBox - 커닝 토글 ===');
    const kerning = await readDialogControl(session,
        el => el.controlTypeName === 'CheckBox' && el.name && el.name.startsWith('커닝'));
    console.log('  대상: "' + kerning.name + '"');

    const kc = centerOf(kerning.rect);
    win32.mouseClick(kc.x, kc.y);
    await sleep(300);

    // ============================================================
    // Test 3: 설정 버튼으로 적용
    // ============================================================
    console.log('\n=== Test 3: 설정 버튼 클릭 ===');
    await clickDialogButton(session, '설정');

    const dialogAfter = findDialog(session);
    if (dialogAfter) {
        console.log('  ⚠ 다이얼로그 안 닫힘');
        dialogAfter.release();
    } else {
        console.log('  다이얼로그 정상 닫힘 ✓');
    }

    // ============================================================
    // Test 4: 재오픈해서 값 검증
    // ============================================================
    console.log('\n=== Test 4: 재오픈해서 변경 검증 ===');
    await controller.pressKeys({ keys: 'Ctrl+A' });
    await sleep(300);
    await openCharFormatDialog(session);
    await switchDialogTab(session, '확장');
    await sleep(500);

    // X 방향 검증
    const xAfter = await readDialogControl(session,
        el => el.controlTypeName === 'Edit' && el.name.startsWith('X 방향'));
    console.log('  X 방향 변경 후: "' + xAfter.name + '"');
    const xOk = xAfter.name.includes('30');
    console.log(xOk ? '  ★ Edit 검증 성공' : '  ⚠ Edit 검증 실패');

    // 커닝 CheckBox 상태 확인 (UIA TogglePattern 없이는 이름만으로 판별 어려움)
    // 대안: 설정을 한 번 더 되돌리고 재확인
    const kAfter = await readDialogControl(session,
        el => el.controlTypeName === 'CheckBox' && el.name && el.name.startsWith('커닝'));
    console.log('  커닝 CheckBox: "' + kAfter.name + '"');
    console.log('  (참고: TogglePattern 미구현으로 이름으로 상태 판별 어려움. 클릭이 전달됐는지만 확인.)');

    // 취소로 닫기
    await clickDialogButton(session, '취소');

    // 원복 (Ctrl+Z)
    for (let i = 0; i < 4; i++) {
        await controller.pressKeys({ keys: 'Ctrl+Z' });
        await sleep(200);
    }
    console.log('\n완료 — 변경 원복');

    console.log('\n=== 최종 결과 ===');
    console.log('  ComboBox: 이전 test-dialog.js에서 검증 완료 ✓');
    console.log('  Edit(SpinImpl): ' + (xOk ? '검증 성공 ✓' : '검증 실패 ⚠'));
    console.log('  CheckBox: 클릭 전달됨 (상태는 TogglePattern 추가 후 확인)');
}

run().catch(e => { console.error('ERROR:', e.message); process.exit(1); });
