# 드롭다운 Discovery 결과

- 생성: 2026-04-18T02:37:34.483Z
- 대상: 붙이기 : ALT+P, 쪽 여백 : ALT+J, 글머리표 : ALT+L, 표 : ALT+B
- 제한: depth=3, maxNodes=100, popupWait=1200ms, probeTimeout=12000ms

## 요약

| 대상 | 탭 | 리본 controlType | 팝업 | 위치 | 자식수 | 분류 힌트 | ms |
|------|----|----|------|------|------|-----------|-----|
| 붙이기 : ALT+P | 편집 | MenuItem | ❌ | - | 0 | no-popup | 5435 |
| 쪽 여백 : ALT+J | 편집 | MenuItem | ✅ | frame | 9 | menu | 5330 |
| 글머리표 : ALT+L | 서식 | MenuItem | ✅ | frame | 34 | menu | 5681 |
| 표 : ALT+B | 편집 | MenuItem | ✅ | frame | 7 | menu | 5571 |

## 붙이기 : ALT+P (편집)

- 리본 항목: `ButtonImpl` / MenuItem
- 사전 ExpandCollapseState: Collapsed
- 클릭 좌표: (223, 162)
- 버튼 rect: 185,101-233,177 (48×76, refreshed)
- 팝업 컨테이너: `N/A` (N/A)
- 팝업 위치: N/A
- notes: no-popup-detected

### 진단 (팝업 미감지 시)
- frame 자식: 11 → 11
- desktop 자식: 8 → 8
- frame 자식 전체 (after):
  - Popup|Window||233,426,265,458
  - HwpRulerTabButtonWnd|Pane||57,233,85,261
  - HwpRulerWnd|Pane||85,233,1237,261
  - HwpRulerWnd|Pane||57,261,85,911
  - HwpMainEditWnd|Edit||85,261,1237,911
  - TitleBarImpl|Unknown(50037)|빈 문서 1 - 글|57,18,1254,62
  - MenuBarImpl|Menu|MenuBar|57,62,1254,95
  - ToolBoxImpl|ToolTip|ToolBox|67,96,1243,186
  - ToolBarImpl|ToolTip|ToolBar|67,186,1243,218
  - StatusBarImpl|StatusBar|StatusBar|57,940,1254,965
  - msctls_progress32|ProgressBar||0,0,0,0

## 쪽 여백 : ALT+J (편집)

- 리본 항목: `ButtonImpl` / MenuItem
- 사전 ExpandCollapseState: Collapsed
- 클릭 좌표: (627, 162)
- 버튼 rect: 581,101-639,177 (58×76, refreshed)
- 팝업 컨테이너: `Popup` (Window)
- 팝업 위치: frame
- notes: -

| depth | name | controlType | className | ExpState | dialogOpener |
|-------|------|-------------|-----------|----------|--------------|
| 0 |  | Window | Popup | - |  |
| 1 |  | List | ListImpl | - |  |
| 2 |  | ListItem | ListBoxItem | - |  |
| 2 |  | ListItem | ListBoxItem | - |  |
| 2 |  | ListItem | ListBoxItem | - |  |
| 2 |  | ListItem | ListBoxItem | - |  |
| 2 |  | ListItem | ListBoxItem | - |  |
| 2 |  | ListItem | ListBoxItem | - |  |
| 1 | 쪽 여백 설정 : ALT+A | MenuItem | ButtonImpl | - |  |

## 글머리표 : ALT+L (서식)

- 리본 항목: `ButtonImpl` / MenuItem
- 사전 ExpandCollapseState: Collapsed
- 클릭 좌표: (1209, 129)
- 버튼 rect: 1176,109-1217,134 (41×25, refreshed)
- 팝업 컨테이너: `Popup` (Window)
- 팝업 위치: frame
- notes: -

| depth | name | controlType | className | ExpState | dialogOpener |
|-------|------|-------------|-----------|----------|--------------|
| 0 |  | Window | Popup | - |  |
| 1 | Scroll | ScrollBar | ScrollImpl | - |  |
| 2 | 최근 사용한 목록 | Unknown(50038) | SeparatorImpl | - |  |
| 2 | 최근 사용한 목록 :  | List | ListImpl | - |  |
| 2 | 글머리표 | Unknown(50038) | SeparatorImpl | - |  |
| 2 | 글머리표 :  | List | ListImpl | - |  |
| 3 | (없음) | ListItem | ListBoxItem | - |  |
| 3 |  | ListItem | ListBoxItem | - |  |
| 3 |  | ListItem | ListBoxItem | - |  |
| 3 |  | ListItem | ListBoxItem | - |  |
| 3 |  | ListItem | ListBoxItem | - |  |
| 3 |  | ListItem | ListBoxItem | - |  |
| 3 |  | ListItem | ListBoxItem | - |  |
| 3 |  | ListItem | ListBoxItem | - |  |
| 3 |  | ListItem | ListBoxItem | - |  |
| 3 |  | ListItem | ListBoxItem | - |  |
| 3 |  | ListItem | ListBoxItem | - |  |
| 3 |  | ListItem | ListBoxItem | - |  |
| 3 |  | ListItem | ListBoxItem | - |  |
| 3 |  | ListItem | ListBoxItem | - |  |
| 3 |  | ListItem | ListBoxItem | - |  |
| 3 |  | ListItem | ListBoxItem | - |  |
| 3 |  | ListItem | ListBoxItem | - |  |
| 3 |  | ListItem | ListBoxItem | - |  |
| 3 |  | ListItem | ListBoxItem | - |  |
| 3 | 사용자 정의 | ListItem | ListBoxItem | - |  |
| 2 | 확인용 글머리표 | Unknown(50038) | SeparatorImpl | - |  |
| 2 |  | List | ListImpl | - |  |
| 3 |  | ListItem | ListBoxItem | - |  |
| 3 |  | ListItem | ListBoxItem | - |  |
| 3 |  | ListItem | ListBoxItem | - |  |
| 3 |  | ListItem | ListBoxItem | - |  |
| 3 | 사용자 정의 | ListItem | ListBoxItem | - |  |
| 1 | 글머리표 모양 : ALT+N | MenuItem | ButtonImpl | - |  |

## 표 : ALT+B (편집)

- 리본 항목: `ButtonImpl` / MenuItem
- 사전 ExpandCollapseState: Collapsed
- 클릭 좌표: (859, 162)
- 버튼 rect: 827,101-867,177 (40×76, refreshed)
- 팝업 컨테이너: `Popup` (Window)
- 팝업 위치: frame
- notes: -

| depth | name | controlType | className | ExpState | dialogOpener |
|-------|------|-------------|-----------|----------|--------------|
| 0 |  | Window | Popup | - |  |
| 1 | 취소 | Text | RowColImpl | - |  |
| 1 | 표 만들기 : ALT+T | MenuItem | ButtonImpl | - |  |
| 1 | 표 그리기 : ALT+W | MenuItem | ButtonImpl | - |  |
| 1 | 표 지우개 : ALT+D | MenuItem | ButtonImpl | - |  |
| 1 | 문자열을 표로 : ALT+L | MenuItem | ButtonImpl | - |  |
| 1 | 표를 문자열로 : ALT+X | MenuItem | ButtonImpl | - |  |

---

## 확정된 분류 규칙 (Phase 3 산출)

### 팝업 위치
- 한컴 리본 드롭다운은 예외 없이 **FrameWindowImpl 직계 자식**으로 생성됨.
- 팝업 컨테이너: `className='Popup'`, `controlType='Window'`, `name=''`.
- 동일 fingerprint Popup이 여러 개 존재할 수 있으므로 **rect 포함 fingerprint로 식별**.

### depth-1 구조 분석
| 자식 구성 | 타입 | 수집 전략 |
|-----------|------|----------|
| ScrollImpl 존재 | `scrollable-gallery` | SeparatorImpl로 섹션 구분, 각 ListImpl 카운트 |
| RowColImpl 존재 | `visual-grid` | 격자 UI (셀 미노출) — 이름 있는 MenuItem 형제만 수집 |
| ListImpl + MenuItem | `gallery-mixed` | ListBoxItem 카운트 + MenuItem 수집 |
| MenuItem only | `menu` | MenuItem 모두 수집 |
| ListImpl only | `gallery-pure` | ListBoxItem 카운트만 (이름 없음이 정상) |
| 기타 / 비어있음 | `unknown` | notes 기록 |

### Dialog-opener 식별
- Plan §2의 "이름 끝 `...`" 가정은 **한글 UI에서 사용 안 됨**.
- 실제 패턴: `/설정 :|사용자 정의|모양 :/` (현재 관찰된 한글 다이얼로그 오프너 표기)
- ListItem의 "사용자 정의"도 dialog-opener (글머리표 Discovery에서 확인)

### 예외 처리
- 팝업이 열리지 않는 항목 (붙이기 사례): clipboard 빈 상태에서 dropdown arrow 비활성 가능성.  
  → `no-popup-detected`로 기록 후 다음 항목 진행 (블로킹 아님).
- rect가 toolbox 범위 밖 (글머리표 사례): 창이 좁아 off-screen overflow.  
  → refresh 후에도 동일하면 notes에 `off-screen` 기록.

### 검증된 방어 설계
- **visited 키**: koffi pointer `String()` 불가 → `(className, name, rect)` 튜플 사용
- **Expand() 패턴**: 한컴 MenuItem에서 **no-op** (호출은 성공하지만 팝업 미개설) — **사용 금지**
- **팝업 탐지**: before/after 스냅샷 diff (fingerprint with rect)
- **닫기**: Escape 4회 (성공률 높음, 다이얼로그와 호환)
