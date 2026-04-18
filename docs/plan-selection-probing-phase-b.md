# Phase B: 표 Selection 상태 프로빙

> Phase A(도형)의 인프라를 재사용해 표(table) 삽입 후 표 컨텍스트 탭을 수집.

---

## 1. 배경

Phase A 완료 후 (`maps/hwp-menu-map.json`):
- 도형 탭: 35 ribbon items, 5/17 dropdowns 성공
- 0 errors
- 커밋 `e772f63`

**표 탭 상태**: 표 안 캐럿 상태에서만 활성화되는 컨텍스트 탭. 현재 맵에 존재하지 않음 (inactive 탭은 `_collectMenuTabs`가 수집 안함).

---

## 2. 삽입 전략

표 삽입 방법 비교:

| 방법 | 장점 | 단점 |
|------|------|------|
| Ctrl+N,T chord | 짧음 | 정확한 chord 타이밍 필요 |
| **입력 → 표 드롭다운 → 표 만들기 → Enter(기본값)** | Phase A 패턴 재사용 | 다이얼로그 + 2단계 클릭 |
| COM 자동화 | 정확 | 새 infra |

**선택: 2번** — `_insertShapeForProbing`과 구조 동일. 다이얼로그 1개만 추가.

### 2.1 사이클
```
1. 문서 끝으로 이동 (Ctrl+End)
2. 입력 탭 활성화
3. "표 : ALT+T" 드롭다운 열기
4. popup에서 "표 만들기 : ALT+T" MenuItem 클릭
5. "표 만들기" 다이얼로그 열림
6. Enter → 기본값(5줄×5칸)으로 표 생성
7. 커서 자동으로 표 안으로 이동
8. 표 컨텍스트 탭 활성화 확인
9. 리본 수집 + 드롭다운 프로빙
10. 정리: Escape + Ctrl+Z×5 (다이얼로그 + 표 삽입 롤백)
```

---

## 3. 구현 범위

**`src/menu-mapper.js`:**
- `_insertTableForProbing()` — 새 method (Phase A 패턴)
- `_cleanupTable()` — 새 method
- `_isTableTabActive()` — 새 method (표 탭 명확히 모름 → 여러 후보 탭 체크)
- `_probeTableTab()` — 새 method
- `run()`에 Step 7 추가 (Step 6 다음)
- options에 `probeTableTab` 플래그 추가

**탭 이름 탐지**: HWP의 표 컨텍스트 탭 이름이 "표" 또는 "표 레이아웃" 등일 수 있음 — 삽입 후 `_collectMenuTabs`에서 **새로 등장한 탭**을 동적으로 탐지 (기존 9탭 목록과 diff).

---

## 4. 구현 단계

### B-1: smoke test (`test-table-insert.js`)
표 삽입 + 새 탭 감지 + 정리. 실제 리본/드롭다운 프로빙 없음.

검증:
- [ ] 표 삽입 후 새 컨텍스트 탭 1개 이상 등장
- [ ] Ctrl+Z×5 후 표 삭제, 새 탭 사라짐

### B-2: `_probeTableTab()` 구현 + `test-probe-table-tab.js` 단독 테스트

### B-3: Step 7 통합 + 전체 재실행

---

## 5. 성공 기준

| 기준 | 값 |
|------|-----|
| 표 컨텍스트 탭 감지 | ≥ 1개 |
| 새 탭 리본 항목 수집 | ≥ 10개 |
| 상태 오염 | 0건 |
| 에러 | 0건 |
| 증가 실행 시간 | ≤ 3분 |

---

## 6. 예상 이슈

- **다이얼로그 Enter 반응**: 표 만들기 다이얼로그의 기본 포커스가 "만들기" 버튼이 아닐 수 있음 → 대체: UIA로 "만들기" 버튼 찾아 클릭
- **Ctrl+Z 부족**: 표 생성은 여러 undo 단위일 수 있음 → Ctrl+Z × 5-8 횟수 조정
- **표 탭 이름 불확실**: `_collectMenuTabs`로 기존 탭과 diff해서 새 탭 찾기
