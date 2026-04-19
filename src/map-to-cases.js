/**
 * Map → Test Cases — maps/<product>-menu-map.json 을 읽어 stress runner가
 * 무작위 샘플링할 수 있는 평탄화된 case 목록을 만든다.
 *
 * 각 case는 다음 형태:
 *   {
 *     id:       'hwp.편집.붙이기',
 *     tab:      '편집',
 *     name:     '붙이기 : ALT+P',
 *     kind:     'ribbon' | 'contextMenu',
 *     category: 'action' | 'dialog' | 'dropdown',
 *     action:   { kind: 'click', x, y }  또는  { kind: 'rightClick', x, y, state }
 *     expect:   { kind: 'dialogAppears' | 'popupAppears' | 'noModal' },
 *     teardownKeys: ['Escape', 'Escape'],
 *     tags:     Set<string>        — body-friendly, table-friendly 등
 *     weight:   number              — 기본 가중치(stress에서 재조정)
 *     risky:    boolean             — map의 isDangerous/blacklisted 반영
 *   }
 *
 * stress runner는 이 case 목록을 StressContext.state에 따라 필터·가중해 샘플링한다.
 */
'use strict';

const fs = require('fs');
const { STATES, STATE_TAB_HINTS } = require('./stress-context');

// =============================================================================
// 위험 분류 — stress 루프에서 기본 제외
// =============================================================================

// 위험 패턴 — stress 실행 시 반드시 제외.
// 창 상태를 영구 변경하는 액션(최대화 풀림·리본 접힘·창 분할)은 후속 iter의
// 맵 좌표 클릭이 엉뚱한 곳을 때리게 만드는 연쇄 실패의 원인.
const DANGEROUS_NAME_PATTERNS = [
    // 파일 I/O
    /^인쇄/, /프린트/, /^저장/, /다른 이름/, /보내기/, /메일/,
    /^닫기/, /^종료/, /나가기/,
    /제거/, /삭제\s*-\s*파일/, /내보내기/, /가져오기/,
    /스크립트/, /매크로.*실행/,
    // 창 상태 변경 — 포커스·최대화·좌표 계산을 깨뜨림
    /^새 창/, /편집 화면 나누기/, /창 분할/, /창 배열/, /창 전환/,
    /리본 최소화/, /리본 숨기/, /^확대/, /^축소/, /화면 크기/,
    /^한 화면에/, /전체 화면/, /^100\s*%/, /^200\s*%/, /폭\s*맞춤/,
    // 모드 전환
    /읽기 전용/, /편집 화면\/미리\s*보기/, /미리\s*보기/,
];

// 보기 탭 전체는 stress에서 제외 — 설정 다이얼로그가 적고 창 상태 변경 액션이 대부분.
const SKIP_TABS = new Set(['보기']);

function isDangerousName(name) {
    if (!name) return false;
    return DANGEROUS_NAME_PATTERNS.some(p => p.test(name));
}

// =============================================================================
// 카테고리 추론
// =============================================================================

function categorizeRibbonItem(item) {
    if (item.type === 'dialog' || item.dialog) return 'dialog';
    if (item.type === 'dropdown' || item.dropdown || item.hasDropdown) return 'dropdown';
    return 'action';
}

function tagsForTab(tabName) {
    const tags = new Set([`tab:${tabName}`]);
    // state별 자연스러운 탭 매핑
    for (const [state, allowedTabs] of Object.entries(STATE_TAB_HINTS)) {
        if (allowedTabs.has(tabName)) tags.add(`state:${state}`);
    }
    return tags;
}

function cleanShortName(fullName) {
    if (!fullName) return '';
    // "붙이기 : ALT+P" → "붙이기"
    // "붙이기 ALT+P"  → "붙이기"
    return fullName.replace(/\s*[:：]?\s*ALT\+[A-Z0-9]+\s*$/i, '').trim();
}

// =============================================================================
// Ribbon 항목 → case
// =============================================================================

function ribbonCase(tabName, item) {
    const short = cleanShortName(item.name);
    const category = categorizeRibbonItem(item);
    const risky = isDangerousName(item.name) || item.type === 'skipped' ||
                  (item.dropdown && item.dropdown.classification === 'blacklisted');

    let expect;
    if (category === 'dialog')   expect = { kind: 'dialogAppears' };
    else if (category === 'dropdown') expect = { kind: 'popupAppears' };
    else                         expect = { kind: 'noModal' };

    // dropdown 케이스는 실제 항목을 선택하기 위해 itemCount를 보존.
    // item의 dropdown.items 개수를 signal로 사용 (맵퍼가 이미 프로빙해둔 값).
    const dropdownItemCount = (item.dropdown && Array.isArray(item.dropdown.items))
        ? item.dropdown.items.length : 0;

    // dialog 케이스는 내부 컨트롤·탭 개수를 보존 — Tab/Space 순회 강도에 사용.
    let dialogControlCount = 0;
    let dialogTabCount = 0;
    if (item.dialog) {
        dialogTabCount = Array.isArray(item.dialog.tabs) ? item.dialog.tabs.length : 0;
        const controlsByTab = item.dialog.controls || {};
        for (const arr of Object.values(controlsByTab)) {
            if (Array.isArray(arr)) dialogControlCount += arr.length;
        }
    }

    return {
        id: `hwp.${tabName}.${short}`,
        tab: tabName,
        name: item.name,
        shortName: short,
        kind: 'ribbon',
        category,
        action: { kind: 'click', x: item.clickX, y: item.clickY },
        expect,
        teardownKeys: ['Escape', 'Escape'],
        dropdownItemCount,
        dialogControlCount,
        dialogTabCount,
        tags: tagsForTab(tabName),
        weight: risky ? 0 : 1,
        risky,
    };
}

// =============================================================================
// Context menu → case
// =============================================================================

function contextCase(stateLabel, item, stateStateTag) {
    const short = cleanShortName(item.name);
    const risky = isDangerousName(item.name);
    // context menu 항목은 clickX/Y가 팝업 내부 좌표라 재현 불가 — 우리는 "우클릭 후
    // 팝업 순회"가 아니라 "항목 이름을 보고 조건 만족 여부 확인"이 아닌 이상
    // 직접 클릭은 하지 않는다. stress에선 ribbon만 랜덤 실행하고 context 메뉴는
    // scenario 쪽에서 우클릭으로 트리거.
    return null; // 제외 (ribbon만 stress에서 실행)
}

// =============================================================================
// 메인 변환
// =============================================================================

function mapToCases(mapJson) {
    const cases = [];

    for (const [tabName, tabData] of Object.entries(mapJson.tabs || {})) {
        if (!tabData || tabData.note === 'backstage') continue;    // 파일 탭은 stress 제외
        if (tabData.note === 'disabled') continue;
        if (SKIP_TABS.has(tabName)) continue;                       // 보기 탭 등 스킵
        if (!Array.isArray(tabData.ribbonItems)) continue;

        for (const item of tabData.ribbonItems) {
            if (!item || !item.name) continue;
            if (item.clickX == null || item.clickY == null) continue;
            cases.push(ribbonCase(tabName, item));
        }
    }

    return cases;
}

// =============================================================================
// state 기반 필터·가중치
// =============================================================================

/**
 * 현재 state에 어울리는 case 목록을 가중치와 함께 반환.
 * @returns {{ case: Case, weight: number }[]}
 */
function weightedForState(cases, state) {
    const allowedTabs = STATE_TAB_HINTS[state] || STATE_TAB_HINTS[STATES.BODY];
    const weighted = [];
    for (const c of cases) {
        if (c.risky) continue;
        const base = c.weight;
        if (base <= 0) continue;
        const stateBonus = allowedTabs.has(c.tab) ? 3 : 1;
        weighted.push({ case: c, weight: base * stateBonus });
    }
    return weighted;
}

function pickWeighted(weighted, prng) {
    if (weighted.length === 0) return null;
    const total = weighted.reduce((s, w) => s + w.weight, 0);
    let r = prng() * total;
    for (const w of weighted) {
        r -= w.weight;
        if (r <= 0) return w.case;
    }
    return weighted[weighted.length - 1].case;
}

// =============================================================================
// 파일 로더
// =============================================================================

function loadMap(path) {
    const raw = fs.readFileSync(path, 'utf8');
    return JSON.parse(raw);
}

module.exports = {
    mapToCases,
    weightedForState,
    pickWeighted,
    loadMap,
    isDangerousName,
};
