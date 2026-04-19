/**
 * Stress Verifier — 매 iteration의 "의도 수행 여부"를 측정하기 위한 검증 유틸.
 *
 * 각 iteration에서:
 *   1) 액션 전 snapshot() — COM state + isModified + fieldList + (선택적으로 UIA)
 *   2) 액션 수행
 *   3) 액션 후 snapshot()
 *   4) diff(before, after) — 구조화된 변화 목록
 *   5) classify(case, interaction, diff) — { verdict: pass|fail|unknown, reason }
 *
 * 설계 원칙:
 *   - best-effort: COM 미가용 시 null snapshot으로 graceful degrade
 *   - 저비용: 기본 snapshot은 ~30ms. UIA 확장 옵션은 필요 시만.
 *   - V1 범위: ribbon 리본 케이스의 dialog/popup 분류 위주. action은 unknown.
 *     추후 case 이름별 expected pattern 테이블을 확장해 unknown 비율을 줄여간다.
 */
'use strict';

const hwpCom = require('./hwp-com');
const sessionModule = require('./session');

// =============================================================================
// Snapshot
// =============================================================================

/**
 * 현재 HWP 상태의 경량 스냅샷.
 * @param {object} [options]
 * @param {boolean} [options.includeUia]  UIA 다이얼로그/탭 체크 포함. 기본 false.
 * @returns {object|null}
 */
function snapshot(options = {}) {
    const s = {
        t: Date.now(),
        com: null,
        isModified: null,
        fieldList: '',
        hasDialog: null,
    };

    if (hwpCom.isAvailable()) {
        s.com = hwpCom.getStateSnapshot();
        const hwp = hwpCom.getHwpObject();
        if (hwp) {
            try { s.isModified = Boolean(hwp.IsModified); } catch (_) {}
            try { s.fieldList = String(hwp.GetFieldList(0, 0) || ''); } catch (_) {}
        }
    }

    if (options.includeUia) {
        try {
            sessionModule.refreshHwpElement();
            const el = sessionModule.getSession().hwpElement;
            if (el) {
                const children = el.findAllChildren();
                s.hasDialog = false;
                for (const c of children) {
                    try { if (c.className === 'DialogImpl') s.hasDialog = true; } catch (_) {}
                    try { c.release(); } catch (_) {}
                }
            }
        } catch (_) {}
    }

    return s;
}

/**
 * 로그용 간결 요약 — 전체 snapshot JSON은 ndjson을 비대하게 만드므로 핵심만 뽑아낸다.
 */
function summarize(snap) {
    if (!snap) return null;
    return {
        parentCtrlId: snap.com ? snap.com.parentCtrlId : null,
        selectedCtrlId: snap.com ? snap.com.selectedCtrlId : null,
        pageCount: snap.com ? snap.com.pageCount : null,
        isEmpty: snap.com ? snap.com.isEmpty : null,
        isModified: snap.isModified,
        fieldCount: _parseFields(snap.fieldList).size,
        hasDialog: snap.hasDialog,
    };
}

// =============================================================================
// Diff
// =============================================================================

function _parseFields(fieldListStr) {
    if (!fieldListStr) return new Set();
    // HWP GetFieldList는 일반적으로 ","/";"/"\n"/"\r\n"로 구분된 필드명 문자열을 반환
    return new Set(String(fieldListStr).split(/[,;\r\n]+/).map(s => s.trim()).filter(Boolean));
}

/**
 * 두 snapshot의 구조화된 diff. 변화 없는 필드는 출력하지 않는다.
 */
function diff(before, after) {
    const d = { changed: false };
    if (!before || !after) return d;

    const b = before.com || {};
    const a = after.com || {};
    if (b.parentCtrlId !== a.parentCtrlId) {
        d.parentCtrlId = { from: b.parentCtrlId, to: a.parentCtrlId };
        d.changed = true;
    }
    if (b.selectedCtrlId !== a.selectedCtrlId) {
        d.selectedCtrlId = { from: b.selectedCtrlId, to: a.selectedCtrlId };
        d.changed = true;
    }
    if (b.pageCount != null && a.pageCount != null && b.pageCount !== a.pageCount) {
        d.pageCountDelta = a.pageCount - b.pageCount;
        d.changed = true;
    }
    if (b.isEmpty !== a.isEmpty && b.isEmpty != null && a.isEmpty != null) {
        d.isEmpty = { from: b.isEmpty, to: a.isEmpty };
        d.changed = true;
    }
    if (before.isModified !== after.isModified && before.isModified != null && after.isModified != null) {
        d.isModified = { from: before.isModified, to: after.isModified };
        d.changed = true;
    }

    const bFields = _parseFields(before.fieldList);
    const aFields = _parseFields(after.fieldList);
    const fieldsAdded = [...aFields].filter(f => !bFields.has(f));
    const fieldsRemoved = [...bFields].filter(f => !aFields.has(f));
    if (fieldsAdded.length) { d.fieldsAdded = fieldsAdded; d.changed = true; }
    if (fieldsRemoved.length) { d.fieldsRemoved = fieldsRemoved; d.changed = true; }

    return d;
}

// =============================================================================
// Classify
// =============================================================================

/**
 * 케이스 + 상호작용 + diff → verdict.
 * V1은 dialog/popup만 분류. action은 unknown.
 *
 * verdict 종류:
 *   'pass'    — 기대에 부합하는 결과
 *   'fail'    — 명백히 기대에 어긋남 (예: dialog 취소했는데 문서가 수정됨)
 *   'unknown' — V1이 분류하지 못하는 케이스
 *
 * reason은 사람이 읽어서 V2 분류 규칙을 추가할 때의 힌트.
 */
function classify(caseObj, interaction, d, before, after) {
    if (!caseObj) return { verdict: 'unknown', reason: 'no-case' };
    const category = caseObj.category;
    const r = { verdict: 'unknown', reason: '', signals: [] };
    // before/after는 선택적 — 주어지면 '이미 modified' 같은 상태-의존 판정에 사용.
    before = before || {};
    after = after || {};

    // -------- dialog 카테고리 --------
    if (category === 'dialog') {
        if (interaction === 'cancel' || interaction === 'navigate-cancel') {
            const modified = d.isModified && d.isModified.to === true;
            const pageChanged = d.pageCountDelta && d.pageCountDelta !== 0;
            const fieldAdded = d.fieldsAdded && d.fieldsAdded.length > 0;
            const fieldRemoved = d.fieldsRemoved && d.fieldsRemoved.length > 0;

            if (modified || pageChanged || fieldAdded || fieldRemoved) {
                r.verdict = 'fail';
                r.reason = 'dialog cancel but persistent state changed';
                if (modified) r.signals.push('modified');
                if (pageChanged) r.signals.push(`pageCount${d.pageCountDelta > 0 ? '+' : ''}${d.pageCountDelta}`);
                if (fieldAdded) r.signals.push(`fields+${d.fieldsAdded.length}`);
                if (fieldRemoved) r.signals.push(`fields-${d.fieldsRemoved.length}`);
            } else {
                r.verdict = 'pass';
                r.reason = 'dialog cancelled cleanly';
            }
            return r;
        }
        // 다른 interaction (none 등) — 다이얼로그가 실제로 떴는지 모름
        r.reason = `dialog expected but interaction=${interaction}`;
        return r;
    }

    // -------- dropdown 카테고리 (map-to-cases는 'dropdown'으로 분류) --------
    if (category === 'dropdown') {
        if (interaction === 'open-only') {
            if (d.changed) {
                r.verdict = 'fail';
                r.reason = 'dropdown open-only but persistent state changed';
                if (d.isModified) r.signals.push('modified');
                if (d.fieldsAdded?.length) r.signals.push(`fields+${d.fieldsAdded.length}`);
                if (d.fieldsRemoved?.length) r.signals.push(`fields-${d.fieldsRemoved.length}`);
            } else {
                r.verdict = 'pass';
                r.reason = 'dropdown opened and closed cleanly';
            }
            return r;
        }
        if (interaction === 'select') {
            // dropdown select는 효과가 다양 — V1에서는 변화 있으면 pass, 없으면 unknown
            // (no-op인 항목도 있으므로 변화 없음을 바로 fail로 보지 않음)
            if (d.changed) {
                r.verdict = 'pass';
                r.reason = 'dropdown select produced state change';
                if (d.isModified) r.signals.push('modified');
                if (d.fieldsAdded?.length) r.signals.push(`fields+${d.fieldsAdded.length}`);
                if (d.parentCtrlId) r.signals.push(`parent:${d.parentCtrlId.from}→${d.parentCtrlId.to}`);
            } else {
                r.reason = 'dropdown select but no observable state change (could be no-op or toggle)';
            }
            return r;
        }
        r.reason = `dropdown expected but interaction=${interaction}`;
        return r;
    }

    // -------- action 카테고리 (V2) --------
    if (category === 'action') {
        const name = caseObj.shortName || caseObj.name || '';
        const rule = _matchActionRule(name, caseObj.tab);
        r.rule = rule ? rule.key : null;
        if (!rule) {
            r.reason = 'action: no rule matched — V2 extension needed';
            return r;
        }

        switch (rule.expect) {
            case 'modified':
                // isModified 전환 (false→true)가 가장 명확한 시그널
                if (d.isModified && d.isModified.to === true) {
                    r.verdict = 'pass';
                    r.reason = `${rule.key}: became modified`;
                    r.signals.push('modified-transition');
                    return r;
                }
                // 이미 modified였다면 — 다른 시그널로 확인 불가할 때 fall-through
                if (d.changed) {
                    r.verdict = 'pass';
                    r.reason = `${rule.key}: observable change (other signal)`;
                    r.signals.push('state-changed');
                    return r;
                }
                // 변화 없음 + 이미 modified였으면 판정 불가 (toggle이 no-op일 수도)
                // 변화 없음 + modified=false 유지면 실패 가능성 높음
                if (after.isModified === false || after.isModified === null) {
                    r.verdict = 'fail';
                    r.reason = `${rule.key}: expected modification, no change`;
                    return r;
                }
                r.reason = `${rule.key}: already modified, no new signal (ambiguous)`;
                return r;

            case 'insert':
                // HWP의 삽입 시그널은 snapshot에서 관찰이 제한적 (GetFieldList는 폼
                // 필드만 반환, memo/footnote는 컨트롤 트리에 있음). ParentCtrlId/
                // SelectedCtrlId 전환이 가장 신뢰 가능한 신호.
                if (d.fieldsAdded?.length || (d.pageCountDelta && d.pageCountDelta > 0)) {
                    r.verdict = 'pass';
                    r.reason = `${rule.key}: insertion detected (fields/pages)`;
                    if (d.fieldsAdded?.length) r.signals.push(`fields+${d.fieldsAdded.length}`);
                    if (d.pageCountDelta) r.signals.push(`pageCount+${d.pageCountDelta}`);
                    return r;
                }
                if (d.parentCtrlId || d.selectedCtrlId) {
                    r.verdict = 'pass';
                    r.reason = `${rule.key}: ctrl transition (caret moved into insert target)`;
                    if (d.parentCtrlId) r.signals.push(`parent:${d.parentCtrlId.from}→${d.parentCtrlId.to}`);
                    if (d.selectedCtrlId) r.signals.push(`sel:${d.selectedCtrlId.from}→${d.selectedCtrlId.to}`);
                    return r;
                }
                if (d.isModified && d.isModified.to === true) {
                    r.verdict = 'pass';
                    r.reason = `${rule.key}: document became modified`;
                    r.signals.push('modified-transition');
                    return r;
                }
                // V2 한계: snapshot에 신호가 안 잡혔지만 insert가 실제로 일어났을 수도
                // (예: 빈 문서에서 쪽 나누기 → pageCount 불변, memo 삽입 → field list 불변).
                // 엄격한 fail 대신 unknown으로 두고, mid-check 도입 이후 V3에서 정교화.
                r.reason = `${rule.key}: no observable signal (insertion may still have occurred; observation-limited)`;
                return r;

            case 'no-change':
                if (d.changed) {
                    // 필드나 페이지 추가는 의도치 않은 부작용
                    if (d.fieldsAdded?.length || d.pageCountDelta) {
                        r.verdict = 'fail';
                        r.reason = `${rule.key}: expected no-change but inserted content`;
                        if (d.fieldsAdded?.length) r.signals.push(`fields+${d.fieldsAdded.length}`);
                        if (d.pageCountDelta) r.signals.push(`pageCount${d.pageCountDelta > 0 ? '+' : ''}${d.pageCountDelta}`);
                        return r;
                    }
                    // isModified 전환만으로는 toggle 종류 동작이 있었을 수 있음 — 엄격 fail 하지 않음
                    r.reason = `${rule.key}: minor state change (isModified/ctrlId), not a failure`;
                    r.signals.push('state-touched');
                    return r;
                }
                r.verdict = 'pass';
                r.reason = `${rule.key}: no state change as expected`;
                return r;

            case 'select-change':
                // caret/selection 변화는 COM snapshot diff에서 parentCtrlId/selectedCtrlId로 관찰.
                if (d.parentCtrlId || d.selectedCtrlId) {
                    r.verdict = 'pass';
                    r.reason = `${rule.key}: selection/caret moved`;
                    if (d.parentCtrlId) r.signals.push(`parent:${d.parentCtrlId.from}→${d.parentCtrlId.to}`);
                    if (d.selectedCtrlId) r.signals.push(`sel:${d.selectedCtrlId.from}→${d.selectedCtrlId.to}`);
                    return r;
                }
                // 같은 컨텍스트 내 이동은 COM으로 관찰 불가 (예: 본문 → 본문)
                r.reason = `${rule.key}: same-context nav (not observable)`;
                return r;

            case 'unknown':
                r.reason = `${rule.key}: rule marks as unclassifiable`;
                return r;
        }
        return r;
    }

    return r;
}

// =============================================================================
// Action name → rule 매핑 (V2)
// =============================================================================

/**
 * 각 rule은 `pattern` (정규식) 또는 `names` (set)로 케이스 이름과 매칭.
 * `expect` 값에 따라 classify가 기대 시그널을 판단.
 *
 * 새 규칙 추가 가이드:
 *   1) 리포트 "Unknown Top 20"에서 빈번한 이름 패턴 추출
 *   2) 해당 동작이 문서를 바꾸는지(modified/insert) vs 아닌지(no-change/select-change)
 *   3) rule에 추가
 */
const ACTION_RULES = [
    // === INSERT 류 — 필드/페이지 추가 기대 ===
    { key: 'insert-memo',      pattern: /^메모$|^새 메모|메모 삽입/,                  expect: 'insert' },
    { key: 'insert-footnote',  pattern: /^각주$|각주 삽입/,                          expect: 'insert' },
    { key: 'insert-endnote',   pattern: /^미주$|미주 삽입/,                          expect: 'insert' },
    { key: 'insert-comment',   pattern: /^주석 추가|댓글 추가/,                      expect: 'insert' },
    { key: 'insert-break',     pattern: /^(쪽|단|줄|구역|페이지) 나누기/,             expect: 'insert' },

    // === FORMAT 적용 (선택 영역) — modified 기대 ===
    { key: 'fmt-toggle',       pattern: /^(진하게|기울임|밑줄|취소선|위 첨자|아래 첨자|강조점|큰따옴표)/, expect: 'modified' },
    { key: 'fmt-case',         pattern: /영문 (대소문자|대문자|소문자)/,             expect: 'modified' },
    { key: 'fmt-char-spacing', pattern: /(자간|장평|글자 (확대|축소|배율))/,         expect: 'modified' },
    { key: 'fmt-align',        pattern: /^(가운데|왼쪽|오른쪽|양쪽|배분) 정렬/,       expect: 'modified' },
    { key: 'fmt-valign',       pattern: /(위쪽|아래쪽|가운데)\s*맞춤/,              expect: 'modified' },
    { key: 'fmt-indent',       pattern: /(첫 줄|왼쪽|오른쪽)\s*(들여쓰기|내어쓰기)/, expect: 'modified' },
    { key: 'fmt-outline',      pattern: /한 수준\s*(감소|증가)|들여쓰기 적용|내어쓰기 적용/, expect: 'modified' },
    { key: 'fmt-numbering',    pattern: /문단 번호|글머리표|번호 매기기/,            expect: 'modified' },
    { key: 'fmt-dropcap',      pattern: /문단 첫 글자 장식/,                        expect: 'modified' },
    { key: 'fmt-linespacing',  pattern: /줄 간격|행간/,                             expect: 'modified' },

    // === 표/셀 편집 — modified 기대 ===
    { key: 'table-cell-edit',  pattern: /셀\s*(나누기|합치기|지우기|삭제)/,          expect: 'modified' },
    { key: 'table-cell-size',  pattern: /셀\s*(높이|너비).*같게|셀\s*크기/,         expect: 'modified' },
    { key: 'table-row-col',    pattern: /(줄|칸|행|열)\s*(추가|삭제|나누기|합치기)/,  expect: 'modified' },
    { key: 'table-border',     pattern: /(셀|표)\s*(테두리|배경)/,                  expect: 'modified' },
    { key: 'table-sort',       pattern: /^정렬$|표 정렬/,                           expect: 'modified' },
    { key: 'table-formula',    pattern: /계산식|쉬운 계산식/,                       expect: 'modified' },
    { key: 'table-layout-z',   pattern: /글 (앞|뒤)으로|글자처럼 취급/,              expect: 'modified' },

    // === 객체(도형/그림/차트) 편집 — 선택 상태 의존. V2는 대체로 unknown ===
    { key: 'obj-outline',      pattern: /윤곽선|바깥선/,                            expect: 'unknown' },
    { key: 'obj-fill',         pattern: /(채우기|배경)$/,                           expect: 'unknown' },
    { key: 'obj-effect',       pattern: /효과|그림자|네온|반사/,                    expect: 'unknown' },
    { key: 'obj-property',     pattern: /^(표|도형|그림|차트) 속성/,                expect: 'modified' },

    // === 네비게이션 — 변화 없거나 selection 이동 ===
    { key: 'nav-prev-next',    pattern: /^(이전|다음|처음|마지막|앞|뒤)\s+/,        expect: 'select-change' },
    { key: 'nav-goto',         pattern: /^찾기|바꾸기|이동$/,                       expect: 'unknown' },

    // === 뷰/모드 토글 — 변화 없음 기대 ===
    { key: 'view-toggle',      pattern: /^(가로|세로|편집 화면|미리 보기|읽기 전용|화면\s*보호)/, expect: 'no-change' },
    { key: 'mode-eraser',      pattern: /지우개|자동 글머리/,                       expect: 'no-change' },
    { key: 'mode-spellcheck',  pattern: /맞춤법/,                                   expect: 'no-change' },

    // === 파일/IO — 별도 처리 ===
    { key: 'file-reload',      pattern: /새로 (고침|만들기)|새 문서/,               expect: 'no-change' },

    // === 사이드바/어시스턴트 — unknown ===
    { key: 'sidebar-open',     pattern: /어시스턴트|작업 창|스크립트|매크로/,       expect: 'unknown' },

    // === 암호/보안 — 대화상자 취소되면 no-change ===
    { key: 'security-dialog',  pattern: /암호|보안 설정|개인정보 보호/,             expect: 'no-change' },

    // === Dialog-opener catch-all — 이름에 "설정"/"속성"/"편집" 류 ===
    { key: 'dialog-opener',    pattern: /(설정|속성|편집기|대화|옵션)$/,            expect: 'no-change' },
];

function _matchActionRule(name, tab) {
    if (!name) return null;
    for (const rule of ACTION_RULES) {
        if (rule.pattern && rule.pattern.test(name)) return rule;
        if (rule.names && rule.names.has(name)) return rule;
    }
    return null;
}

module.exports = { snapshot, summarize, diff, classify, _matchActionRule };
