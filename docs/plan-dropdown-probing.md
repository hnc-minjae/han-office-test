# 드롭다운 메뉴 프로빙 구현 계획

> 한컴오피스 2027 메뉴 매퍼(`src/menu-mapper.js`) 확장 — 리본 드롭다운 메뉴의 내부 구조를 자동 수집하기 위한 계획서.  
> 본 문서는 **새 세션에서 맥락 없이 바로 구현 시작** 가능하도록 자체 완결적으로 작성되었다.

---

## 1. 배경 & 현재 상태

### 1.1 프로젝트 위치
- 저장소: `C:\Users\lmj79\dev\han-office-test` (GitHub: `hnc-minjae/han-office-test`)
- 브랜치: `master` — 최신 커밋 `fc5edae` "멀티제품 지원 및 메뉴 자동 매핑 추가"
- 현재 메뉴 맵: `maps/hwp-menu-map.json` (9탭, 168 리본 항목, **21 다이얼로그, 472 컨트롤**)

### 1.2 현재 JSON에 수집된 것 / 빠진 것
| 구분 | 내용 | 상태 |
|------|------|------|
| ✅ 수집됨 | 리본 탭 이름·좌표·액세스 키 | 11개 탭 |
| ✅ 수집됨 | 리본 항목 이름·좌표·타입 | 168개 |
| ✅ 수집됨 | 다이얼로그 내부 (탭·컨트롤) | 21개 다이얼로그, 472 컨트롤 |
| ✅ 수집됨 | ToolBar / StatusBar | 22 + 17 |
| ❌ **빠짐** | **드롭다운 내부 메뉴** | **65개 — 본 계획의 대상** |
| ❌ 빠짐 | 파일 탭 Backstage | 별도 계획 필요 |
| ❌ 빠짐 | 도형 탭 (selection 의존) | 별도 계획 필요 |
| ❌ 빠짐 | Selection 다른 상태 (캐럿/도형/표) | 별도 계획 필요 |

### 1.3 과거 실패 분석
과거 세션에서 "메뉴트리 타고 들어가기" 방식으로 시도했으나 **"표" 드롭다운(행×열 그리드)에서 실패**. 원인:
- 표 삽입 드롭다운은 마우스 호버로 크기를 미리보기하는 **그리드 UI**
- 각 그리드 셀이 UIA 트리에 MenuItem/Button으로 **노출되지 않음**
- 한컴이 커스텀 렌더링으로 구현

### 1.4 오늘 세션에서 확인된 한컴 UIA 특성
| 컨트롤 | UIA invoke | 조작 방법 |
|--------|-----------|-----------|
| `MainMenuItemImpl` (리본 탭) | ❌ no-op | 마우스 클릭 (boundingRect 왼쪽 1/3) |
| `ButtonImpl` (리본 버튼) | ❌ no-op | 마우스 클릭 (boundingRect 상단 1/3) |
| `DialogButtonImpl` | ✅ | 마우스 클릭 권장 |
| `ComboImpl` | ✅ | 더블클릭 + Ctrl+A + 타이핑 |
| `SpinImpl` | ❌ | UIA 조작 불가 — WM_SETTEXT 필요 |

**중요 교훈:** UIA 패턴 지원이 **선언만 되고 no-op**인 경우가 많으므로, "UIA 상태 보고 vs 실제 동작"을 교차 검증해야 함.

### 1.5 다이얼로그 위치 규칙 (재사용)
한컴 다이얼로그는 **`FrameWindowImpl`의 직계 자식**:
```
FrameWindowImpl (main window)
├── className === 'DialogImpl' ← 이게 다이얼로그
├── HwpMainEditWnd
├── MenuBarImpl
├── ToolBoxImpl
├── ToolBarImpl
└── StatusBarImpl
```
드롭다운 팝업도 비슷한 위치에 나타날 가능성이 높으나 **Discovery로 검증 필요**.

---

## 2. 드롭다운 분류 (가설)

드롭다운은 최소 4가지 타입. 타입별로 수집 전략 달라짐:

| 타입 | 예시 | UIA 구조 (가설) | 수집 난이도 |
|------|------|----------------|------------|
| **1. 단순 메뉴** | 붙이기 → 붙이기/골라 붙이기 | MenuItem children 명확 | 쉬움 |
| **2. 갤러리(텍스트)** | 쪽 여백 → 좁게/보통/넓게/사용자 정의 | List/Button 자식 + 이름 | 중간 |
| **3. 갤러리(그리드)** | 표 (N×M), 색 팔레트 | 그리드 셀 UIA 노출 안 됨 | 어려움 |
| **4. 복합 메뉴** | 글머리표 (갤러리 + "다른 문단 번호 모양...") | 섞인 구조 + 다이얼로그 진입점 | 중간 |

---

## 3. 새 전략 — 3단계 접근

### Strategy 1: UIA 스냅샷 차분 (dropdown popup 탐지)
드롭다운을 열기 전/후에 UIA 트리 스냅샷을 떠서 **"새로 생긴 요소"**를 팝업으로 식별.  
팝업이 루트(Desktop) 자식일지, `FrameWindowImpl` 자식일지, 혹은 다른 위치일지 사전엔 알 수 없음 → Discovery로 확인.

### Strategy 2: 트리 휴리스틱 분류
팝업 내용을 재귀 순회(visited set + depth limit + node limit)하여 아래 규칙으로 분류:

| 조건 | 분류 | 수집 |
|------|------|------|
| MenuItem 자식 N개 (이름 있음) | `menu` | 이름·단축키 나열 |
| List/Button 자식 (이름 있음) | `gallery-text` | 이름 나열 |
| 자식 0개 또는 이름 없는 커스텀 자식 | `gallery-visual` | "그리드 UI" 플래그만 기록 |
| "다른/기타 XX 설정..." 항목 존재 | `mixed` | 항목 나열 + 다이얼로그 진입점 식별 |
| 이름 끝 `...` | 개별 항목 → dialog opener로 표시 | `hasDialogOpener: true` |

### Strategy 3: ExpandCollapseState 교차 검증
각 메뉴 항목의 `ExpandCollapseState`를 읽어 **서브메뉴 유무 선언값 수집**, 실제 호버 후 서브메뉴 등장 여부와 비교.

---

## 4. UIA 확장 필요 사항

### 4.1 `get_CurrentExpandCollapseState` 추가
**파일:** `src/uia.js`

```javascript
// VTABLE.IExpandCollapsePattern에 추가
IExpandCollapsePattern: {
    Release: 2,
    Expand: 3,
    Collapse: 4,
    get_CurrentExpandCollapseState: 5,  // ← 추가
},
```

UIElement에 getter 추가:
```javascript
get expandCollapseState() {
    const pp = [null];
    const hr = comCall(this._ptr, VTABLE.IUIAutomationElement.GetCurrentPattern,
                       proto.GetCurrentPattern, PatternId.ExpandCollapse, pp);
    if (hr !== 0 || !pp[0]) return null;
    const out = [0];
    comCall(pp[0], VTABLE.IExpandCollapsePattern.get_CurrentExpandCollapseState,
            proto.GetInt32Prop, out);
    comRelease(pp[0]);
    return out[0];  // 0=Collapsed, 1=Expanded, 2=Partially, 3=Leaf
}
```

**ExpandCollapseState 상수 export:**
```javascript
const ExpandCollapseState = { Collapsed: 0, Expanded: 1, PartiallyExpanded: 2, LeafNode: 3 };
```

### 4.2 (선택) `GetCurrentPropertyValue` 일반 프로퍼티 읽기
PROPERTYID 기반 접근이 가능하면 더 유연함. 단 VARIANT 핸들링이 복잡해서 필요할 때만.

### 4.3 (선택) `runtimeId` — 트리 순회 visited 키
```javascript
get runtimeId() {
    // GetRuntimeId (vtable[4])가 SAFEARRAY를 반환
    // 복잡하니 대안으로 포인터 주소(this._ptr을 16진수 hex 문자열로)를 키로 사용
}
```
**간단 대안:** `this._ptr`의 포인터 주소값을 visited 키로 사용. UIA는 같은 물리 요소에 대해 다른 COM 포인터를 줄 수 있지만, **같은 순회 내에서는 안정적**.

---

## 5. 무한루프 방어 7계층

### Layer 1: 상수 한도
```javascript
const DD_LIMITS = {
    closeAttempts: 5,              // Escape 최대 5회
    popupTreeMaxNodes: 500,
    popupTreeMaxDepth: 8,
    popupWaitMs: 1500,             // 팝업 등장 대기 데드라인
    recoveryAttempts: 2,
    perDropdownBudgetMs: 10_000,   // 드롭다운 1개 전체 예산
    perTabBudgetMs: 120_000,       // 탭 전체 예산
    totalProbingBudgetMs: 900_000, // 전체 15분
};
```

### Layer 2: no-progress break (상태 fingerprint 비교)
```javascript
function fingerprint(session) {
    const top = session.hwpElement.findAllChildren();
    const hash = top.map(c => c.className + ':' + c.controlTypeName).join('|');
    top.forEach(c => c.release());
    return hash;
}
// 루프 내에서 fingerprint 바뀌지 않으면 break
```

### Layer 3: 트리 순회 visited + depth + count 가드
```javascript
function walkTree(el, visited, depth) {
    if (depth > DD_LIMITS.popupTreeMaxDepth) return null;
    if (visited.size > DD_LIMITS.popupTreeMaxNodes) throw new Error('TreeSizeExceeded');
    const key = String(el._ptr);  // 포인터 주소를 키로
    if (visited.has(key)) return null;
    visited.add(key);
    // ...
}
```

### Layer 4: 시간 예산 워치독
```javascript
function budgeted(deadline) {
    return () => {
        if (Date.now() > deadline) throw new Error('BudgetExceeded');
    };
}

async function probeDropdown(item) {
    const check = budgeted(Date.now() + DD_LIMITS.perDropdownBudgetMs);
    check(); await openDropdown(item);
    check(); const popup = await waitForPopup();
    check(); const items = walkTree(popup, new Set(), 0);
    check(); await closeDropdown();
    return items;
}
```

### Layer 5: 재귀 복구 차단
```javascript
let _recoveryInProgress = false;
async _ensureUiHealthy() {
    if (_recoveryInProgress) return;
    _recoveryInProgress = true;
    try { /* 복구 로직 */ }
    finally { _recoveryInProgress = false; }
}
```

### Layer 6: 드롭다운 블랙리스트
```javascript
const DANGEROUS_DROPDOWN_PATTERNS = [
    // Discovery 후 발견되는 것 추가
    // 예: 매크로 실행계열, 파일 열기 등
];
```

### Layer 7: 에러 시 격리 + 다음 항목 진행
드롭다운 하나 실패 = 에러 기록 + `_recover()` + `_ensureUiHealthy()` + **다음 항목으로 이동**.  
**절대 재시도 루프 금지**.

---

## 6. 구현 단계 (Phase 1~5)

### Phase 1: UIA 확장 (코드 없이 시작)
**파일:** `src/uia.js`
- [ ] `ExpandCollapseState` 상수 추가 (4값)
- [ ] `IExpandCollapsePattern.get_CurrentExpandCollapseState` 추가 (vtable[5])
- [ ] `UIElement.expandCollapseState` getter 추가
- [ ] export에 `ExpandCollapseState` 포함

### Phase 2: Discovery — `test-dropdown-probe.js` 작성
**신규 파일:** `test-dropdown-probe.js` (프로젝트 루트)

**목표:** 4종 샘플로 UIA 노출 방식 경험적 확인
- 붙이기 (ALT+P) — 단순 메뉴
- 쪽 여백 (ALT+J) — 갤러리(텍스트)  
- 글머리표 (ALT+L, 편집 탭 ALT+S 제외) — 복합
- 표 (ALT+B) — 그리드 (과거 실패 지점)

**안전장치 (Discovery도 방어 필요):**
- 각 드롭다운 5초 타임아웃
- 재귀 없음 — 관찰만 (depth 3 1회 순회)
- 테스트 간 Escape 2회 + 상태 확인

**수집 항목:**
- 클릭 전 스냅샷 (루트 자식, FrameWindowImpl 자식)
- 클릭 후 스냅샷 (차분으로 팝업 위치 발견)
- 팝업의 className, controlTypeName
- 팝업 자식 트리 (depth 3, 최대 100노드)
- 각 자식의 ExpandCollapseState (Phase 1에서 추가한 속성)
- 특이사항 (비어있음, 이름 없음, 그리드 힌트)

**출력:** `docs/discovery-dropdown-results.md` (표로 정리)

### Phase 3: Classification 규칙 구체화
Discovery 결과로 분류 규칙을 **실측 기반**으로 재정의. 이 시점에서 계획서의 "가설 분류"를 확정된 알고리즘으로 대체.

예상 성과물: `_classifyDropdown(popup)` 함수 구현체 스펙

### Phase 4: menu-mapper 통합
**파일:** `src/menu-mapper.js`

**신규 메서드:**
```
_probeAllDropdowns()
├── 탭별 순회 (SKIP_PROBING_TABS 제외)
├── hasDropdown=true 항목만 대상
├── 매 항목 전 _resetCaretContext()
├── _probeDropdown(item) 호출
└── 매 항목 후 _ensureNoPopup() + _ensureUiHealthy()

_probeDropdown(item)
├── 워치독 시작 (10초)
├── 스냅샷 before
├── 마우스 클릭 — 드롭다운 화살표 영역 (boundingRect 오른쪽 하단 1/4)
├── 팝업 감지 (waitForPopupWithDeadline)
├── 팝업 분류 + 수집 (walkTree with visited set)
├── Escape로 닫기 (최대 5회, no-progress break)
└── 결과 반환

_classifyDropdown(popup)
└── Discovery 결과에 따라 구체화

_ensureNoPopup()
└── FrameWindowImpl 자식 중 Popup/Menu 있으면 Escape 2회
```

**호출 위치:** `run()`의 Step 4 (`_probeAllDialogs`) 다음에 **Step 5: `_probeAllDropdowns`** 추가.

### Phase 5: Validation & 재실행
- 블랙리스트 없이 편집 탭만 먼저 테스트
- 문제 없으면 전체 실행 (예산 내 완료 기대)
- 결과 `maps/hwp-menu-map.json` diff 확인
- errors 통계 확인, 실패 항목 분석

---

## 7. JSON 스키마 확장

기존 드롭다운 항목에 `dropdown` 필드 추가:

```json
{
  "name": "쪽 여백 : ALT+J",
  "controlType": "MenuItem",
  "hasDropdown": true,
  "type": "dropdown",
  "clickX": 1234,
  "clickY": 159,
  "dropdown": {
    "classification": "gallery-text",
    "popupClass": "PopupMenuImpl",
    "itemCount": 4,
    "items": [
      {
        "name": "좁게",
        "controlType": "MenuItem",
        "expandCollapseState": 3,
        "hasSubmenu": false,
        "isDialogOpener": false
      },
      {
        "name": "사용자 정의...",
        "controlType": "MenuItem",
        "expandCollapseState": 3,
        "hasSubmenu": false,
        "isDialogOpener": true
      }
    ],
    "probeDurationMs": 823,
    "notes": []
  }
}
```

`notes` 예: `["tree-truncated"]`, `["visual-grid-detected"]`, `["popup-position-unusual"]`

### stats 확장
```json
"stats": {
  "totalTabs": 9,
  "totalRibbonItems": 168,
  "totalDialogs": 21,
  "totalControls": 472,
  "totalDropdowns": 65,
  "totalDropdownItems": 0,
  "dropdownErrors": 0,
  "errors": 0
}
```

---

## 8. 테스트 시나리오 (Validation)

### Discovery 단계 (Phase 2 검증)
- [ ] 붙이기: 단순 메뉴 자식 2개 이상 감지
- [ ] 쪽 여백: 갤러리 텍스트 항목 수집
- [ ] 글머리표: 복합 구조 — 갤러리 + "다른..." 발견
- [ ] 표: 그리드 감지 — 자식 0 또는 무의미한 구조 확인 (성공적으로 **안전하게 실패**)

### 방어 계층 검증 (Phase 4 단위 테스트)
- [ ] 닫히지 않는 팝업 시뮬레이션 → 5회 Escape 후 포기하는지
- [ ] 트리가 깊이 10까지 있는 경우 → depth 8에서 멈추는지
- [ ] 고의로 타임아웃 발생 → 다음 항목으로 격리 진행하는지
- [ ] 복구 중 또 복구 트리거 → 재귀 차단되는지

### 전체 통합 (Phase 5)
- [ ] errors = 0
- [ ] totalDropdownItems > 100 (예상)
- [ ] 전체 실행 시간 ≤ 15분
- [ ] 최소 편집/입력/서식/쪽/도구 탭 각 하나 이상의 드롭다운 완전 수집

---

## 9. 성공 기준

| 항목 | 기준 |
|------|------|
| 단순 메뉴 드롭다운 (type: `menu`) | ≥ 10개 완전 수집 |
| 갤러리 텍스트 드롭다운 (type: `gallery-text`) | ≥ 5개 수집 |
| 그리드 드롭다운 (type: `gallery-visual`) | 안전하게 감지만, 무한루프 없이 스킵 |
| 복합 드롭다운 (type: `mixed`) | 다이얼로그 진입점 식별 |
| 무한루프 | 0건 |
| 에러 복구 연쇄 실패 | 0건 |
| 총 실행 시간 | 기존 맵 생성 시간 + 10분 이내 |

---

## 10. 다음 세션 체크리스트 (순서대로)

1. [ ] 한글(Hwp.exe) 실행 중인지 확인 — 없으면 사용자에게 실행 요청
2. [ ] `git status` — working tree 깨끗한지 확인
3. [ ] 이 문서 + `maps/hwp-menu-map.json` + `src/menu-mapper.js` 읽어 맥락 복원
4. [ ] **Phase 1 먼저** — `src/uia.js`에 `ExpandCollapseState` + getter 추가. 노드 스크립트로 간단 smoke test.
5. [ ] **Phase 2 진행** — `test-dropdown-probe.js` 작성 후 실행. 결과를 `docs/discovery-dropdown-results.md`로 정리.
6. [ ] Discovery 결과 검토 후 **Phase 3 진행** — 분류 규칙 확정.
7. [ ] **Phase 4 구현** — `menu-mapper.js`에 `_probeAllDropdowns`, `_probeDropdown`, `_classifyDropdown` 추가.
8. [ ] **편집 탭 단독 테스트** (test-probe-edit-tab.js 패턴 재사용).
9. [ ] **Phase 5 전체 재실행** — 백그라운드로 `node src/menu-mapper.js hwp > maps/full-run3.log 2>&1`.
10. [ ] 결과 검증 후 **커밋 + 푸시**.

---

## 11. 참고 — 오늘 세션의 교훈 요약

- **UIA invoke/expand는 한컴 커스텀 컨트롤에서 no-op** → 마우스 클릭이 필수
- **boundingRect는 vtable[43]** — RECT {left,top,right,bottom} int32×4
- **다이얼로그는 FrameWindowImpl 직계 자식**, className='DialogImpl'
- **탭/버튼의 클릭 위치**: 왼쪽 1/3 (탭), 상단 1/3 (버튼) — 오른쪽은 드롭다운 화살표
- **요소 검색은 반드시 `{controlType, name}` 쌍으로** — Label과 ComboBox 이름 충돌 주의
- **캐럿 컨텍스트 통일 필수** — Escape + Ctrl+Home + Ctrl+A
- **UI 건강 체크 필수** — ToolBoxImpl 없으면 Ctrl+F1, "(숫자)" 창 있으면 Ctrl+F4
- **Selection 종류(캐럿/글자/도형/표)에 따라 메뉴 동작 달라짐** — 본 프로빙은 글자 선택 하나의 상태만 반영

---

## 12. 관련 파일 빠른 참조

| 파일 | 역할 |
|------|------|
| `src/uia.js` | UIAutomation + UIElement 래퍼 (확장 대상) |
| `src/win32.js` | 마우스 클릭, 키보드 입력, boundingRect 호환 |
| `src/hwp-controller.js` | 고수준 컨트롤러 (pressKeys, findElement 등) |
| `src/session.js` | 싱글톤 세션 관리 |
| `src/menu-mapper.js` | 메뉴 크롤러 (확장 대상) |
| `maps/hwp-menu-map.json` | 현재 결과 — Phase 5 후 갱신됨 |
| `test-probe-edit-tab.js` | 편집 탭 단독 프로빙 (패턴 재사용 참고) |
| `test-dialog.js` | 다이얼로그 조작 E2E 테스트 (패턴 참고) |
| `~/.claude/projects/.../memory/` | 프로젝트 기억 — 새 세션에서 자동 로드 |
