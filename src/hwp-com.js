/**
 * HWP COM Automation Wrapper — winax를 통한 한/글 OLE Automation 접근.
 *
 * 목적:
 *   - 실시간 caret/selection 상태 조회 (UIA/추정으로는 부정확)
 *   - ParentCtrl / CurSelectedCtrl / IsEmpty / PageCount 등의 메타데이터
 *
 * 원칙:
 *   - best-effort: winax 미설치/HWP 미실행이면 null 반환, 러너는 추정 방식 폴백
 *   - 단일 싱글톤 HwpObject — 세션 전체에서 공유
 *   - COM 호출 실패 시 throw 없이 null 반환
 *
 * 참고: 한컴 han-ax 프로젝트의 hwp-com-api.ts 패턴을 CJS로 이식.
 */
'use strict';

let _winax = null;
let _hwp = null;
let _initAttempted = false;

function _tryRequireWinax() {
    try { return require('winax'); } catch (_) { return null; }
}

/**
 * winax + HwpObject를 lazy 초기화. 최초 호출 시 HwpFrame.HwpObject.2 생성.
 * HWP가 이미 실행 중이면 그 인스턴스에 attach, 아니면 새 프로세스 기동.
 *
 * @returns {object|null} HwpComObject 또는 실패 시 null
 */
function init() {
    if (_hwp) return _hwp;
    if (_initAttempted) return null; // 이전 시도가 실패했으면 재시도하지 않음
    _initAttempted = true;

    _winax = _tryRequireWinax();
    if (!_winax) return null;

    try {
        _hwp = new _winax.Object('HwpFrame.HwpObject.2');
        // 메시지박스 자동 응답 (무인 자동화 필수 — 0x10000 = 기본 응답)
        try { _hwp.SetMessageBoxMode(0x10000); } catch (_) {}
        // 보안 모듈 — 파일 접근 승인 팝업 자동 처리
        try { _hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample'); } catch (_) {}
        // 창 보이기 — COM이 새 HWP를 띄웠을 경우 대비
        try {
            if (_hwp.XHwpWindows) {
                const w = _hwp.XHwpWindows.Item(0);
                if (w) w.Visible = true;
            }
        } catch (_) {}
        return _hwp;
    } catch (e) {
        process.stderr.write(`[hwp-com] init failed: ${e.message}\n`);
        _hwp = null;
        return null;
    }
}

/** winax + HWP COM 이용 가능 여부. */
function isAvailable() {
    return _hwp !== null;
}

/**
 * 현재 HWP의 caret / 선택 상태를 읽는다.
 * 반환값 예시:
 *   { parentCtrlId: 'tbl', selectedCtrlId: null, isEmpty: false, pageCount: 3, error: null }
 *   null — winax 미설치 또는 COM 호출 전부 실패
 */
function getStateSnapshot() {
    if (!_hwp) return null;
    const out = {
        parentCtrlId: null,
        selectedCtrlId: null,
        isEmpty: null,
        pageCount: null,
        error: null,
    };
    try {
        try {
            const p = _hwp.ParentCtrl;
            out.parentCtrlId = (p && p.CtrlID) ? String(p.CtrlID).trim() : null;
        } catch (_) {}
        try {
            const s = _hwp.CurSelectedCtrl;
            out.selectedCtrlId = (s && s.CtrlID) ? String(s.CtrlID).trim() : null;
        } catch (_) {}
        try { out.isEmpty   = Boolean(_hwp.IsEmpty); }   catch (_) {}
        try { out.pageCount = Number(_hwp.PageCount); }  catch (_) {}
    } catch (e) {
        out.error = e.message;
    }
    return out;
}

/**
 * HwpCtrlCode.CtrlID → stress-context STATE 매핑.
 * HWP의 CtrlID 주요 값:
 *   'tbl'   — 표 (Table)
 *   'gso'   — 일반 도형 개체 (GenShapeObject) — 도형·글상자 포함
 *   'plate' — 글상자 내부 (도형의 텍스트 영역)
 *   'pic'   — 그림 (Picture)
 *   'chart' — 차트
 *   'header'/'footer' 등은 GetCurFieldName 쪽 — 별도 처리 필요
 */
function ctrlIdToContextState(snapshot) {
    if (!snapshot) return null;

    // 선택된 개체가 있으면 그 쪽이 우선 (도형/그림/차트 "선택됨")
    const sel = snapshot.selectedCtrlId;
    if (sel === 'tbl')           return 'in-table-cell';      // 표 전체 선택도 같은 컨텍스트로 취급
    if (sel === 'gso')           return 'in-shape-text';
    if (sel === 'plate')         return 'in-shape-text';
    if (sel === 'pic')           return 'in-image-selected';
    if (sel === 'chart')         return 'in-chart-selected';

    // 캐럿이 컨테이너 내부에 있는 경우
    const par = snapshot.parentCtrlId;
    if (par === 'tbl')   return 'in-table-cell';
    if (par === 'gso')   return 'in-shape-text';
    if (par === 'plate') return 'in-shape-text';

    return null; // 명확한 단서 없음 — 호출자가 body로 간주
}

/** 싱글톤 HwpObject 직접 반환 — SetPos/MoveToField 등 고수준 조작에 사용. */
function getHwpObject() {
    return _hwp;
}

/**
 * caret을 특정 list로 이동. HWP의 "list" 개념:
 *   0 = 본문
 *   1+ = 표/도형 내부 (삽입 순서대로 번호 부여)
 * list 1에 SetPos(1, 0, 0)이면 첫 번째 비본문 컨테이너(주로 표)의 첫 셀 시작점.
 * 반환: 성공 여부 (설정한 list가 존재하지 않으면 실패)
 */
function moveToList(listIndex) {
    if (!_hwp) return false;
    try {
        _hwp.SetPos(listIndex, 0, 0);
        return true;
    } catch (_) {
        return false;
    }
}

function release() {
    if (_winax && _hwp) {
        try { _winax.release(_hwp); } catch (_) {}
    }
    _hwp = null;
}

module.exports = {
    init,
    isAvailable,
    getStateSnapshot,
    ctrlIdToContextState,
    getHwpObject,
    moveToList,
    release,
};
