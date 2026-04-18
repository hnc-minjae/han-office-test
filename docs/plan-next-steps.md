# 한컴오피스 메뉴 맵 — 다음 작업 가이드

> 본 문서는 새 세션에서 **맥락 복원 없이 바로 이어서 작업** 가능하도록 작성된 핸드오프 문서.
> 최신 상태: 2026-04-18, 최신 커밋 `3833387` (Phase G 완료).

---

## 1. 현재 상태 (Phase G까지 완료)

### 1.1 완료된 커밋 체인
```
3833387 feat: Phase G — 객체별 우클릭 컨텍스트 메뉴 확장
ef608dc feat: Selection 상태 프로빙 Phase F — 그림 컨텍스트 탭 지원
0c3b078 fix: setForeground 후 창이 축소되지 않도록 재최대화
ed3b5b7 feat: Phase E 우클릭 컨텍스트 메뉴 + 속도/안정성 개선
cab70a0 feat: Selection 상태 프로빙 Phase D — 차트 컨텍스트 탭 지원
72eeef9 fix: 드롭다운 화살표 위치 aspect-ratio 적응 클릭
db62a9c feat: Selection 상태 프로빙 Phase B — 표 컨텍스트 탭 지원
e772f63 feat: Selection 상태 프로빙 Phase A — 도형 탭 지원
f9a171a feat: 리본 드롭다운 프로빙 + 포커스 안전 가드 추가
fc5edae feat: 한컴오피스 2027 멀티제품 지원 및 메뉴 자동 매핑 추가
```

### 1.2 현재 맵 통계 (`maps/hwp-menu-map.json`)
- **15 tabs** (파일=backstage, 9 일반, 5 컨텍스트)
- **283 ribbon items**
- **21 dialogs / 510 controls**
- **94 dropdowns / 1125 dropdown items**
- **6 contextMenus** (caret, text-selected, shape-selected, table-inside, chart-selected, image-selected)
- **0 errors**

### 1.3 Step 1-10 구조 (`src/menu-mapper.js`)
| Step | 내용 | 구현 위치 |
|------|------|----------|
| 1 | 메뉴 탭 수집 | `_collectMenuTabs` |
| 2 | 탭별 리본 항목 수집 | `_collectRibbonItems` |
| 3 | ToolBar / StatusBar | `_collectFixedBars` |
| 4 | 다이얼로그 프로빙 | `_probeAllDialogs` |
| 5 | 드롭다운 프로빙 | `_probeAllDropdowns` |
| 6 | 도형 컨텍스트 탭 (Phase A) | `_probeShapeTab` |
| 7 | 표 컨텍스트 탭 (Phase B) | `_probeTableTabs` |
| 8 | 차트 컨텍스트 탭 (Phase D) | `_probeChartTabs` |
| 9 | 그림 컨텍스트 탭 (Phase F) | `_probeImageTab` |
| 10 | 우클릭 컨텍스트 메뉴 (Phase E+G) | `_probeContextMenus` |

---

## 2. 다음 작업 — Phase H: 파일 Backstage

### 2.1 배경
파일 탭은 **풀스크린 Backstage 패널**로 열린다. 일반 리본 구조와 달라 별도 처리 필요. 현재 `map.tabs['파일'] = { note: 'backstage', ribbonItems: [] }`로 비어있음.

### 2.2 Discovery 결과 (진단 완료)
파일 탭 클릭 후 `FrameWindowImpl` 자식 구조:
```
Popup | Window (좌표: 585,408-617,440 — 내부 placeholder)
HwpMainEditWnd | Edit
HwpRulerWnd | Pane
...
TitleBarImpl
MenuBarImpl | Menu | MenuBar
ToolBoxImpl | ToolTip | ToolBox
ToolBarImpl
StatusBarImpl
WorkspaceImpl | ToolBar | 다른 문서 활용
msctls_progress32 | ProgressBar
```

**관찰**: `BackstageImpl`/`FileMenuImpl` 같은 명확한 컨테이너가 frame 자식으로 나타나지 **않음**. Backstage가 기존 ToolBoxImpl이나 MainEditWnd를 덮어쓰는 방식일 가능성.

**가설**:
1. Backstage는 `ToolBoxImpl`의 컨텐츠를 교체 (리본 → 파일 메뉴 패널)
2. 또는 `HwpMainEditWnd`가 비활성되고 Backstage 패널이 모달 역할
3. 또는 Desktop root의 별도 윈도우로 떠있음

### 2.3 Phase H 구현 방향

**A. 확인 스냅샷** (진단 필요):
- 파일 클릭 전/후 `FrameWindowImpl`의 **모든 descendants** 비교 (depth 5)
- Desktop root children 비교 (새 top-level 윈도우 발생 여부)
- `ToolBoxImpl` 자식이 일반 리본 → Backstage 메뉴로 바뀌는지

**B. 네비게이션 설계**:
- Backstage 내 카테고리 버튼(새 문서/열기/저장/인쇄/옵션 등) 클릭 → 우측 패널에 상세 UI 표시
- 각 카테고리마다:
  - 제목
  - 하위 액션 버튼 / 다이얼로그 오프너
  - 최근 문서 리스트 (열기 카테고리)

**C. 구현 스켈레톤** (`src/menu-mapper.js`):
```javascript
async _probeFileBackstage() {
    // 1. 파일 탭 클릭
    const menuTabs = await this._collectMenuTabs();
    const fileTab = menuTabs.find(t => t.name === '파일');
    await this._safeClick(fileTab.clickX, fileTab.clickY);
    await sleep(DELAY.dialog);

    // 2. Backstage 컨테이너 찾기 — ToolBoxImpl 또는 신규 요소
    const backstage = this._findBackstageContainer();

    // 3. 카테고리 수집 (왼쪽 메뉴)
    const categories = this._collectBackstageCategories(backstage);

    // 4. 각 카테고리 순회
    const result = { categories: {} };
    for (const cat of categories) {
        await this._safeClick(cat.clickX, cat.clickY);
        await sleep(DELAY.long);
        result.categories[cat.name] = this._collectBackstagePanel();
    }

    // 5. 정리: Escape
    await this._safeKeys('Escape');
    await sleep(DELAY.long);

    this.map.tabs['파일'] = {
        ...this.map.tabs['파일'],
        note: 'backstage',
        backstage: result,
    };
}
```

**D. 예상 카테고리** (HWP 2027 일반):
- 새 문서 / 열기 / 최근 문서
- 저장 / 다른 이름으로 저장 / PDF로 저장
- 인쇄 / 보내기
- 문서 정보 / 보안
- 옵션 (큰 다이얼로그)
- 도움말
- 종료

**E. 주의사항**:
- **"옵션"은 대형 다이얼로그** — 다이얼로그 프로빙 infrastructure 재사용 가능
- **"인쇄" 카테고리**는 프린터 설정 — 실제 인쇄 방지 필수 (시뮬레이션 금지)
- **"종료" 카테고리**는 절대 클릭하면 안 됨 — DANGEROUS_CATEGORIES 블랙리스트 필요
- **"열기"/"저장"**: 파일 대화상자 트리거 — Escape로 닫아야
- Backstage 종료는 **Escape 1회**로 가능 (확인 완료)

### 2.4 Phase H 구현 단계
1. **Discovery 테스트** (`test-backstage-discovery.js`): frame의 모든 descendants depth 5 덤프 + 파일 클릭 전/후 비교 + 카테고리 좌표/이름 식별
2. **smoke test** (`test-backstage-smoke.js`): 카테고리 하나(예: "새 문서") 클릭 후 상세 패널 구조 확인
3. **menu-mapper 통합**: `_probeFileBackstage()` + Step 11
4. **위험 카테고리 블랙리스트**: 종료, 인쇄(실행 방지)
5. **전체 재실행 + 커밋**

---

## 3. 남은 추가 Phase들 (Phase H 이후)

### 3.1 Phase I: 수식 편집기
- 입력 → 수식 → standalone 수식 편집기 윈도우
- 별도 UIA 트리 (수식 편집기는 독립 윈도우)
- ESC로 종료

### 3.2 Phase J: 드롭다운 flakiness 개선
- 현재 몇몇 드롭다운이 context-lost로 실패 (특히 차트 서식 재실행 시)
- 삽입 직후 UI 안정화 대기 시간 증가 + 재시도 강화
- 주요 개선점: chart cleanup 후 다음 Phase 진입 전 문서 상태 검증

### 3.3 Phase K: 글상자(Text Box) 상태
- 도형 탭의 "글상자 여백·방향·정렬·연결" 등 4개 no-popup 드롭다운은 **직사각형이 아닌 가로 글상자**를 삽입해야 활성화
- `_insertShapeForProbing`에 shape 타입 파라미터 추가 (default='직사각형', 추가='가로 글상자')
- 글상자 모드로 Step 6 한 번 더 실행

---

## 4. 핵심 기술 인사이트 (새 세션에서 필수 숙지)

### 4.1 한컴 UIA 특이점
| 특징 | 내용 |
|------|------|
| 팝업 위치 | `FrameWindowImpl` 직계 자식 `className='Popup', controlType='Window'` |
| 동일 fingerprint 중복 | rect 포함 fingerprint 필수 (baseline popup 32×32가 항상 존재) |
| ExpandCollapse 패턴 | **no-op** — 호출 성공하지만 팝업 열리지 않음. 마우스 클릭만 사용 |
| Invoke 패턴 | 마찬가지로 no-op in 대부분의 ribbon 요소 |
| koffi void* | `String(ptr)` 변환 불가 → visited 키는 `(className, name, rect)` 사용 |

### 4.2 split-button 드롭다운 클릭
- **Wide (width > height)**: 오른쪽 끝 화살표 → `(right-8, center-y)`
- **Tall (height ≥ width)**: 하단 화살표 → `(left+width*0.8, top+height*0.8)`
- 중앙 클릭은 action 트리거 (드롭다운 미개방)

### 4.3 포커스 & 창 상태
- **setForeground 호출이 창을 축소시킴** (SW_RESTORE) → 매번 `win32.maximizeWindow(hwnd)` 재호출 필수
- `_setForegroundMaximized()` 헬퍼 사용
- `_ensureForeground()`가 이벤트 전 HWP 포커스 검증 (다른 앱 누설 차단)
- `win32.getForegroundPid()`로 현재 foreground PID 확인

### 4.4 객체 상태 전환 패턴
| 객체 | 삽입 방법 | 선택 유지 방법 |
|------|----------|----------------|
| 도형 (직사각형) | 편집/도형 드롭다운 → 직사각형 → 캔버스 단일 클릭 | 이전 좌표 재클릭 |
| 표 (5×5) | 입력/표/표 만들기 → Enter | Ctrl+End (표 안 진입) |
| 차트 | 입력/차트/묶은 가로 막대형 → Escape(데이터 대화상자) | 캔버스 중앙 재클릭 |
| 그림 | PS로 이미지→클립보드 → Ctrl+V → **F11**(개체 선택) | F11 재실행 |

### 4.5 런처 해제
HWP 시작 시 `LauncherImpl`가 있으면 메뉴바가 없음. `_dismissLauncher()`가 "새 문서" ListItem을 **더블클릭**해서 빈 문서로 진입. 단일 클릭이나 Escape는 불충분.

### 4.6 timing 튜닝
```javascript
const DELAY = { short: 50, medium: 120, long: 250, dialog: 450 };
const DD_LIMITS = {
    closeAttempts: 2,          // 드롭다운 닫기 Escape 횟수
    popupWaitMs: 350,          // 팝업 등장 대기
    perDropdownBudgetMs: 8_000,
    perTabBudgetMs: 180_000,
    totalProbingBudgetMs: 900_000,
};
```
더 줄이면 lazy-load popup 데이터 손실 (차트 89 items 등).

---

## 5. 새 세션 시작 체크리스트

1. [ ] `git log --oneline -10` — 최신 커밋 확인
2. [ ] 이 문서(`docs/plan-next-steps.md`) 읽기
3. [ ] `maps/hwp-menu-map.json`의 stats 확인
4. [ ] HWP 실행 확인 (`tasklist | grep Hwp`) — 실행 중이면 OK, 아니면 `node -e "require('./src/hwp-controller').launch({product:'hwp'}).then(r => console.log(r.pid))"`
5. [ ] Phase H 시작 — `test-backstage-discovery.js` 작성부터
6. [ ] 각 Phase 완료 시 커밋 (format: `feat: Phase X — ...`)

---

## 6. 파일 맵

| 파일 | 역할 |
|------|------|
| `src/menu-mapper.js` | 메인 크롤러, Step 1-10 구현 |
| `src/uia.js` | UIA COM 래퍼 (`ExpandCollapseState` 포함) |
| `src/win32.js` | Win32 API, 마우스/키보드, `maximizeWindow`, `getForegroundPid` |
| `src/hwp-controller.js` | 고수준 controller (`attach`, `launch`, `pressKeys` 등) |
| `src/session.js` | 싱글톤 세션 (`getSession`, `refreshHwpElement`, `getUia`) |
| `maps/hwp-menu-map.json` | 현재 맵 (commit에 포함) |
| `maps/fixtures/probe-test.png` | Phase F용 테스트 PNG |
| `docs/plan-dropdown-probing.md` | 드롭다운 프로빙 원본 계획서 |
| `docs/plan-selection-probing.md` | Phase A 계획서 |
| `docs/plan-selection-probing-phase-b.md` | Phase B 계획서 |
| `docs/discovery-dropdown-results.md` | 드롭다운 Discovery 결과 + 분류 규칙 |
| `test-*.js` | smoke/discovery 테스트 (재실행 가능) |

---

## 7. 실행 커맨드 레퍼런스

```bash
# 전체 맵 생성 (Step 1-10)
node src/menu-mapper.js hwp > maps/full-run.log 2>&1

# 편집 탭만 드롭다운 테스트 (~40s)
node test-probe-dropdown-edit-tab.js > maps/speed-test.json 2> maps/speed-test.log

# 개별 smoke tests
node test-shape-insert.js         # Phase A
node test-table-insert.js         # Phase B
node test-chart-insert.js         # Phase D
node test-image-insert.js         # Phase F
node test-right-click.js          # Phase E
node test-dropdown-probe.js       # Discovery (원본)

# HWP 재시작
taskkill //F //IM Hwp.exe
node -e "require('./src/hwp-controller').launch({product:'hwp'}).then(r => console.log(r.pid))"

# 커밋 (identity override)
git -c user.email=minjae@hancom.com -c user.name=minjae commit -m "feat: Phase X — ..."
```
