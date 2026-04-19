/**
 * Stress Context — 장시간 stress 테스트 중 캐럿 위치/셀렉션 상태를 추적한다.
 *
 * HWP는 UIA로 캐럿이 본문/머리말/각주 등 어디에 있는지 직접 물어볼 수 없다.
 * runner가 "어떤 액션을 명령했는가"를 근거로 state를 추정하고, 주기적으로
 * "본문 복귀" 명령으로 drift를 보정한다.
 *
 * state 전이는 불확실하므로 이 파일은 best-guess 추적기다.
 * 정확한 검증이 필요하면 _forceReturnToBody() 후 state를 'body'로 리셋한다.
 */
'use strict';

const controller = require('./hwp-controller');

// =============================================================================
// State 정의
// =============================================================================

const STATES = Object.freeze({
    BODY:           'body',
    IN_HEADER:      'in-header',
    IN_FOOTER:      'in-footer',
    IN_FOOTNOTE:    'in-footnote',
    IN_ENDNOTE:     'in-endnote',
    IN_MEMO:        'in-memo',
    IN_TABLE_CELL:  'in-table-cell',
    IN_SHAPE_TEXT:  'in-shape-text',
});

// 각 state에서 "자연스럽게 실행 가능한" 리본 탭 힌트.
// map-to-cases가 리본 항목에 context 태그를 부여할 때 참조.
const STATE_TAB_HINTS = Object.freeze({
    [STATES.BODY]:          new Set(['편집', '입력', '서식', '쪽', '보기', '보안', '검토', '도구']),
    [STATES.IN_HEADER]:     new Set(['편집', '입력', '서식', '쪽']),
    [STATES.IN_FOOTER]:     new Set(['편집', '입력', '서식', '쪽']),
    [STATES.IN_FOOTNOTE]:   new Set(['편집', '입력', '서식']),
    [STATES.IN_ENDNOTE]:    new Set(['편집', '입력', '서식']),
    [STATES.IN_MEMO]:       new Set(['편집', '서식']),
    [STATES.IN_TABLE_CELL]: new Set(['편집', '서식', '표 디자인', '표 레이아웃', '입력']),
    [STATES.IN_SHAPE_TEXT]: new Set(['편집', '서식', '도형']),
});

// =============================================================================
// StressContext
// =============================================================================

class StressContext {
    constructor() {
        this.state = STATES.BODY;
        this._history = [];      // 최근 20개 state 전이 기록 (디버깅용)
        this._lastForceBodyAt = Date.now();
    }

    get tabHints() {
        return STATE_TAB_HINTS[this.state] || STATE_TAB_HINTS[STATES.BODY];
    }

    /**
     * 외부에서 state 전이를 통지한다. 실제 동작이 성공했을 때만 호출할 것.
     */
    transition(newState, reason = '') {
        if (!Object.values(STATES).includes(newState)) {
            throw new Error(`unknown state: ${newState}`);
        }
        const prev = this.state;
        this.state = newState;
        this._history.push({ t: Date.now(), from: prev, to: newState, reason });
        if (this._history.length > 20) this._history.shift();
    }

    /**
     * 일정 시간/iter마다 drift 방지용으로 본문 복귀 강제.
     * Esc 여러 번 + Ctrl+Home으로 안전하게 body 최상단으로 이동.
     */
    async forceReturnToBody() {
        try { await controller.pressKeys({ keys: 'Escape' }); } catch (_) {}
        try { await controller.pressKeys({ keys: 'Escape' }); } catch (_) {}
        try { await controller.pressKeys({ keys: 'Escape' }); } catch (_) {}
        try { await controller.pressKeys({ keys: 'Ctrl+Home' }); } catch (_) {}
        this.transition(STATES.BODY, 'forceReturnToBody');
        this._lastForceBodyAt = Date.now();
    }

    /** 마지막 강제 복귀 이후 경과 ms */
    msSinceForceBody() {
        return Date.now() - this._lastForceBodyAt;
    }

    snapshot() {
        return {
            state: this.state,
            history: [...this._history],
            msSinceForceBody: this.msSinceForceBody(),
        };
    }
}

module.exports = { StressContext, STATES, STATE_TAB_HINTS };
