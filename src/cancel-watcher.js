/**
 * Cancel Watcher — ESC 더블탭으로 장기 실행 프로세스 안전 중단
 *
 * GetAsyncKeyState로 ESC를 폴링하며, 사용자가 정해진 윈도 내에
 * 두 번 탭하면 onCancel 콜백을 호출한다. 우리 코드가 SendInput으로
 * 보낸 ESC는 win32.getLastEscSentAt() 쿨다운으로 구분.
 *
 * 사용: menu-mapper, stress-runner 양쪽에서 공유.
 */
'use strict';

const win32 = require('./win32');

const CANCEL = {
    pollMs:        75,     // GetAsyncKeyState 폴링 주기
    doubleTapMs:   1000,   // 첫 ESC 이후 두 번째 ESC 인정 윈도
    selfEscMs:     250,    // 우리가 보낸 ESC 무시 쿨다운 (2차 안전망)
    minHeldPolls:  2,      // 사용자 탭 판정에 필요한 연속 pressed 폴링 수.
                           // pollMs=75ms × 2 = 150ms 이상 유지돼야 "의도적 탭".
                           // SendInput ESC는 keydown+keyup이 수 ms 내 끝나 1 폴링만 스침 → 차단됨.
};

class CancelledError extends Error {
    constructor(reason = 'cancelled by user (ESC x2)') {
        super(reason);
        this.name = 'CancelledError';
    }
}

/**
 * ESC 더블탭 감시자. setInterval로 폴링하면서 ESC의 "down-edge"를 감지.
 * - 우리가 내부에서 SendInput으로 보낸 ESC는 selfEscMs 쿨다운 이내면 무시.
 * - 첫 탭 이후 doubleTapMs 이내에 두 번째 탭이 오면 onCancel 호출.
 */
class CancelWatcher {
    constructor(onCancel) {
        this._onCancel = onCancel;
        this._timer = null;
        this._heldPolls = 0;
        this._tapCommitted = false;
        this._firstTapAt = 0;
        this._fired = false;
    }

    start() {
        if (this._timer) return;
        this._heldPolls = 0;
        this._tapCommitted = false;
        this._firstTapAt = 0;
        this._fired = false;
        this._timer = setInterval(() => this._tick(), CANCEL.pollMs);
    }

    stop() {
        if (this._timer) {
            clearInterval(this._timer);
            this._timer = null;
        }
    }

    _tick() {
        if (this._fired) return;
        const pressed = win32.isKeyPressed(win32.VK_ESCAPE);
        const now = Date.now();

        if (pressed) {
            this._heldPolls++;
        } else {
            this._heldPolls = 0;
            this._tapCommitted = false; // ESC가 떼어져야 다음 탭 인정
        }

        // 진짜 사용자 탭 판정 조건:
        //   1) 연속 minHeldPolls회 이상 pressed (SendInput ESC는 도달 불가)
        //   2) 우리가 보낸 ESC 쿨다운 밖 (2차 안전망)
        //   3) 같은 press를 중복 카운트하지 않도록 latch
        const sinceSelfEsc = now - win32.getLastEscSentAt();
        if (this._heldPolls >= CANCEL.minHeldPolls
            && sinceSelfEsc >= CANCEL.selfEscMs
            && !this._tapCommitted) {
            this._tapCommitted = true;
            if (this._firstTapAt && (now - this._firstTapAt) <= CANCEL.doubleTapMs) {
                this._fired = true;
                try { this._onCancel(); } catch (_) {}
            } else {
                this._firstTapAt = now;
            }
        }

        // doubleTap 윈도 만료
        if (this._firstTapAt && (now - this._firstTapAt) > CANCEL.doubleTapMs) {
            this._firstTapAt = 0;
        }
    }
}

module.exports = { CancelledError, CancelWatcher, CANCEL };
