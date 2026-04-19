/**
 * Stress Seeder — 장시간 stress 테스트 시작 시 1페이지 분량의 풍부한 개체
 * (단락·표·도형·각주·미주·메모·머리말·꼬리말)를 문서에 삽입한다.
 *
 * 모든 단계는 best-effort. 실패해도 log만 남기고 다음 단계 진행 —
 * seed가 완전치 않아도 runner는 가용한 상태에서 시작할 수 있어야 한다.
 *
 * HWP 단축키 규칙 (chord):
 *   Ctrl+N, T  — 표 만들기    (Ctrl+N 후 T)
 *   Ctrl+N, N  — 각주
 *   Ctrl+N, E  — 미주
 * parseKeyExpression은 "Ctrl+N+T"를 Ctrl 홀드로 해석하므로 chord는 두 번의
 * pressKeys 호출로 분리해야 한다.
 */
'use strict';

const controller = require('./hwp-controller');
const win32 = require('./win32');
const { STATES } = require('./stress-context');

const DELAY = { tick: 100, short: 200, medium: 450, long: 800, dialog: 700 };

const log = (icon, msg) => process.stderr.write(`${icon} [seeder] ${msg}\n`);
const sleep = (ms) => new Promise(r => setTimeout(r, ms));

// =============================================================================
// 코퍼스 — 한국어/영문/숫자/특수문자 혼합
// =============================================================================

const CORPUS = {
    title: 'STRESS 테스트 문서 — 한컴오피스 2027 안정성 검증',
    paragraphs: [
        '본 문서는 자동화된 장시간 stress 테스트의 기반이 되는 시드 문서입니다. ' +
        '한글·영문·숫자·특수문자를 혼합한 다양한 문단과 표, 각주, 미주, 메모 등의 개체를 포함합니다.',

        'The quick brown fox jumps over the lazy dog. 1234567890 !@#$%^&*() ' +
        '가나다라마바사아자차카타파하 — 다국어 혼합 문자열은 한글 2027의 유니코드 처리를 검증합니다.',

        '긴 문서에서도 안정적으로 동작해야 합니다. 표의 셀 이동, 머리말/꼬리말 전환, ' +
        '복사·붙여넣기 같은 자주 쓰이는 시나리오가 반복되어도 강제 종료 없이 ' +
        '예측 가능한 응답을 주어야 합니다.',
    ],
    tableCell: ['항목', '수량', '단가', '합계'],
    tableRows: [
        ['연필', '12', '500', '6,000'],
        ['공책', '5', '2,000', '10,000'],
        ['지우개', '8', '300', '2,400'],
    ],
    footnoteText: '각주 시드 — STRESS 테스트용 각주 콘텐츠입니다.',
    endnoteText:  '미주 시드 — 미주는 문서 끝에 누적됩니다.',
    memoText:     '메모 시드 — 사이드에 붙는 메모 개체 테스트.',
};

// =============================================================================
// 유틸
// =============================================================================

async function typeLine(text) {
    await controller.typeText({ text, useClipboard: true });
    await sleep(DELAY.tick);
}

async function pressKey(keys) {
    await controller.pressKeys({ keys });
    await sleep(DELAY.tick);
}

/** chord: Ctrl+N 후 T (두 번의 pressKeys) */
async function chord(first, second) {
    await controller.pressKeys({ keys: first });
    await sleep(80);
    await controller.pressKeys({ keys: second });
    await sleep(DELAY.short);
}

function findRibbonItem(map, tabName, pattern) {
    const tab = map && map.tabs && map.tabs[tabName];
    if (!tab || !tab.ribbonItems) return null;
    const re = pattern instanceof RegExp ? pattern : new RegExp(pattern);
    return tab.ribbonItems.find(i => i.name && re.test(i.name));
}

async function clickRibbonByMap(map, tabName, pattern, label) {
    const item = findRibbonItem(map, tabName, pattern);
    if (!item) {
        log('⚠', `${label}: 맵에서 "${tabName} > ${pattern}" 찾을 수 없음 — 스킵`);
        return false;
    }
    win32.mouseClick(item.clickX, item.clickY);
    await sleep(DELAY.medium);
    return true;
}

// =============================================================================
// 개별 시드 스텝 — 실패 시 경고만
// =============================================================================

async function seedTitleAndParagraphs() {
    try {
        await pressKey('Ctrl+Home');
        await typeLine(CORPUS.title);
        await pressKey('Enter');
        for (const p of CORPUS.paragraphs) {
            await typeLine(p);
            await pressKey('Enter');
        }
        log('✓', '제목·단락 3개 삽입');
        return true;
    } catch (e) {
        log('⚠', `제목·단락 실패: ${e.message}`);
        return false;
    }
}

async function seedTable() {
    try {
        // Ctrl+N, T → 표 만들기 다이얼로그 → 기본값으로 Enter → 3행 4열 생성 위해
        // dialog에서 행/열 조정이 필요하지만 여기선 기본값 사용하고 바로 Enter로 진행.
        // 이후 셀을 순회하며 텍스트 입력.
        await chord('Ctrl+N', 'T');
        await sleep(DELAY.dialog);
        // 다이얼로그에서 기본 2x2로 Enter (안정성 우선 — 복잡 설정은 드리프트 위험)
        await pressKey('Enter');
        await sleep(DELAY.long);

        // 표 안에 커서가 있다고 가정. 셀마다 텍스트 + Tab
        const cells = [...CORPUS.tableCell, ...CORPUS.tableRows.flat()];
        for (let i = 0; i < Math.min(4, cells.length); i++) {
            try {
                await typeLine(cells[i]);
                await pressKey('Tab');
            } catch (_) {}
        }
        // 표 밖으로 나가기: Ctrl+Pgdown(다음 표) 대신 Ctrl+End로 문서 끝
        await pressKey('Ctrl+End');
        log('✓', '표 삽입 + 셀 텍스트');
        return true;
    } catch (e) {
        log('⚠', `표 실패: ${e.message}`);
        try { await pressKey('Escape'); } catch (_) {}
        return false;
    }
}

async function seedFootnote() {
    try {
        await pressKey('Enter');
        await chord('Ctrl+N', 'N');
        await sleep(DELAY.long);
        await typeLine(CORPUS.footnoteText);
        // 각주 편집 모드 나가기 — Shift+Esc 또는 본문 클릭이 표준.
        // 가장 안전한 건 Esc 두 번.
        await pressKey('Escape');
        await pressKey('Escape');
        log('✓', '각주 삽입');
        return true;
    } catch (e) {
        log('⚠', `각주 실패: ${e.message}`);
        return false;
    }
}

async function seedEndnote() {
    try {
        await pressKey('Enter');
        await chord('Ctrl+N', 'E');
        await sleep(DELAY.long);
        await typeLine(CORPUS.endnoteText);
        await pressKey('Escape');
        await pressKey('Escape');
        log('✓', '미주 삽입');
        return true;
    } catch (e) {
        log('⚠', `미주 실패: ${e.message}`);
        return false;
    }
}

async function seedMemo(map) {
    try {
        // 입력 탭 > 메모(드롭다운 or 직접 버튼). 맵에서 "메모" 이름 찾기.
        const clicked = await clickRibbonByMap(map, '입력', /메모/, '메모');
        if (!clicked) return false;
        await sleep(DELAY.long);
        // 드롭다운일 수 있음 → "새 메모" 같은 첫 항목 Enter
        await pressKey('Enter');
        await sleep(DELAY.long);
        await typeLine(CORPUS.memoText);
        await pressKey('Escape');
        await pressKey('Escape');
        log('✓', '메모 삽입');
        return true;
    } catch (e) {
        log('⚠', `메모 실패: ${e.message}`);
        return false;
    }
}

async function seedHeader(map, runId) {
    try {
        // 쪽 탭 > 머리말/꼬리말 → 머리말 드롭다운 → 첫 항목(양쪽) Enter
        const clicked = await clickRibbonByMap(map, '쪽', /머리말\/꼬리말|머리말/, '머리말');
        if (!clicked) return false;
        await sleep(DELAY.long);
        await pressKey('Enter'); // 기본 항목
        await sleep(DELAY.long);
        await typeLine(`STRESS-HEADER-${runId}`);
        // 본문 복귀
        await pressKey('Escape');
        await pressKey('Escape');
        await pressKey('Ctrl+Home');
        log('✓', '머리말 삽입');
        return true;
    } catch (e) {
        log('⚠', `머리말 실패: ${e.message}`);
        try { await pressKey('Escape'); } catch (_) {}
        return false;
    }
}

async function seedFooter(map, runId) {
    try {
        const clicked = await clickRibbonByMap(map, '쪽', /꼬리말|머리말\/꼬리말/, '꼬리말');
        if (!clicked) return false;
        await sleep(DELAY.long);
        await pressKey('Enter');
        await sleep(DELAY.long);
        await typeLine(`STRESS-FOOTER-${runId}`);
        await pressKey('Escape');
        await pressKey('Escape');
        await pressKey('Ctrl+End');
        log('✓', '꼬리말 삽입');
        return true;
    } catch (e) {
        log('⚠', `꼬리말 실패: ${e.message}`);
        try { await pressKey('Escape'); } catch (_) {}
        return false;
    }
}

// =============================================================================
// 오케스트레이션
// =============================================================================

/**
 * 초기 시드 — runner 시작 시 1회 호출.
 * @param {object} map - hwp-menu-map.json 파싱 결과
 * @param {StressContext} ctx - 상태 추적기
 * @param {string} runId - 이번 run의 식별자(머리말/꼬리말 마커용)
 */
async function seed(map, ctx, runId) {
    log('▶', '문서 시딩 시작 (1페이지 분량)');
    const results = {};
    results.titleParagraphs = await seedTitleAndParagraphs();
    results.table           = await seedTable();
    results.footnote        = await seedFootnote();
    results.endnote         = await seedEndnote();
    results.memo            = await seedMemo(map);
    results.header          = await seedHeader(map, runId);
    results.footer          = await seedFooter(map, runId);
    try { await ctx.forceReturnToBody(); } catch (_) {}
    ctx.transition(STATES.BODY, 'seed-complete');
    const ok = Object.values(results).filter(Boolean).length;
    log('ℹ', `시딩 결과: ${ok}/${Object.keys(results).length} 성공`);
    return results;
}

/**
 * 경량 보강 — 30분마다 1회 호출해서 문서를 점점 키움.
 */
async function reseed(map, ctx) {
    log('▶', '경량 보강 시딩 (표/각주 추가)');
    try { await ctx.forceReturnToBody(); } catch (_) {}
    try { await pressKey('Ctrl+End'); } catch (_) {}
    await seedTable();
    await seedFootnote();
    try { await ctx.forceReturnToBody(); } catch (_) {}
}

module.exports = { seed, reseed, CORPUS };
