# Selection 상태 프로빙 구현 계획

> 한컴오피스 2027 메뉴 매퍼 확장 — Selection 상태에 따라 활성화되는 컨텍스트 탭/메뉴를 수집.
> 본 문서는 Phase A(도형 탭)만 다룬다. Phase B/C(표, 그림, 수식)는 A 검증 후 확장.

---

## 1. 배경

### 1.1 현재 상태 (2026-04-18)
- 최신 맵: `maps/hwp-menu-map.json` — 9탭, 168 리본 항목, 21 다이얼로그, 508 컨트롤, **51 드롭다운 / 593 items**
- 최근 커밋 `f9a171a` "리본 드롭다운 프로빙 + 포커스 안전 가드 추가"
- `docs/discovery-dropdown-results.md`에 드롭다운 구조 분석 완료

### 1.2 남은 빈틈 (plan-dropdown-probing §1.2)
현재 "글자 선택" 상태 1개만 반영. **"도형" 탭은 비활성으로 남아있음** (map에 `note: 'disabled'`):
```json
"도형": { "accessKey": "...", "ribbonItems": [], "note": "disabled" }
```
도형이 문서에 선택되어 있을 때만 활성화되는 컨텍스트 탭이라 일반 프로빙으로는 접근 불가.

### 1.3 목표 (Phase A)
도형 탭의 리본 항목 + 드롭다운 구조를 수집해 `maps/hwp-menu-map.json`에 채워넣기.

---

## 2. 접근 전략

### 2.1 도형 삽입 방법 선정

| 방법 | 장점 | 단점 | 선정 |
|------|------|------|------|
| A) 입력 탭 → 도형 드롭다운 → 사각형 → 마우스 드래그 | UI 사실상 그대로 | 드래그 좌표 계산 복잡 | |
| B) 단축키로 기본 도형 삽입 | 간단 | HWP가 적절한 단축키 제공하는지 불확실 | |
| **C) 입력 탭 → 도형 드롭다운 → 사각형 → 단일 클릭** | 단일 클릭으로 기본 크기 도형 삽입 가능 (HWP 기본 동작) | 검증 필요 | ✅ |
| D) COM 자동화로 직접 도형 객체 삽입 | 100% 안정 | 새 infra 필요 | |

**C) 채택** — 대부분의 Office 제품은 도형 도구 선택 후 단일 클릭 시 기본 크기 도형 삽입함. 실패 시 B로 폴백.

### 2.2 상태 전환 사이클

```
[시작: 글자 선택 상태]
 ├─ 1. 문서 맨 끝으로 이동 (Ctrl+End)
 ├─ 2. 입력 탭으로 전환
 ├─ 3. "도형" 드롭다운 열기 (이미 맵 보유: 입력 탭)
 ├─ 4. 드롭다운에서 "직사각형" 메뉴 아이템 클릭
 ├─ 5. HwpMainEditWnd 영역 내 안전한 좌표에서 마우스 단일 클릭
 │     (또는 클릭-드래그로 작은 도형 생성)
 ├─ 6. 도형 자동 선택 여부 확인
 │     └─ `_collectMenuTabs()` 재호출하여 "도형" 탭이 isEnabled=true가 되는지 확인
 ├─ 7. 도형 탭으로 전환
 ├─ 8. 리본 항목 수집 + 드롭다운 프로빙
 └─ 9. 정리: Escape → Ctrl+Z × 3 (도형 선택 해제 + 삽입 롤백)

[종료: 글자 선택 상태 복귀]
```

### 2.3 실패 케이스 방어
- **도형 탭이 나타나지 않음** → 삽입 실패로 간주, Escape + Ctrl+Z + 다음 시도
- **도형이 너무 크게 삽입됨** → 삭제 후 재시도 또는 수용
- **원치 않는 상태 오염** (매크로 실행 등) → `_recover()` 호출 + 다음 상태 진행 (격리)

---

## 3. 구현 범위

### 3.1 신규 코드

**`src/menu-mapper.js`:**
- `_ensureDocumentReady()` — 문서가 편집 가능한 상태인지 확인, 아니면 새 문서 열기
- `_insertShapeForProbing()` — 도형 삽입 + 선택 보장
- `_probeShapeTab()` — 도형 탭 프로빙 (리본 수집 + 드롭다운 프로빙)
- `_cleanupShape()` — Ctrl+Z로 삽입 롤백
- `run()`에 **Step 6** 추가 (Step 5 뒤)

**`src/win32.js`:**
- `mouseDrag(x1, y1, x2, y2)` — 단일 클릭으로 부족할 때 폴백

### 3.2 map 스키마 확장

도형 탭 데이터는 기존 `map.tabs['도형']`에 저장 (`note: 'disabled'` 제거):
```json
"도형": {
  "accessKey": "...",
  "uiaName": "도형 ...",
  "ribbonItems": [...],         // ← 새로 채워짐
  "contextState": "shape-selected"  // ← 상태 표시용 신규 필드
}
```

### 3.3 stats 확장
```json
"stats": {
  ...
  "contextTabs": 1,          // 새 필드
  "contextRibbonItems": N    // 새 필드
}
```

---

## 4. 구현 단계 (Phase A, 5 단계)

### Phase A-1: 도형 삽입 Helper 작성 & smoke test (single file)
**파일:** `test-shape-insert.js` (신규)

도형 1개만 삽입 + 탭 활성화 확인 + Ctrl+Z 정리. 다른 기능 없음.

검증 항목:
- [ ] 도형 삽입 후 "도형" 탭이 isEnabled=true
- [ ] Ctrl+Z 3회 후 "도형" 탭이 isEnabled=false로 복귀
- [ ] 문서가 깨끗한 상태 유지

### Phase A-2: `_probeShapeTab()` 구현
**파일:** `src/menu-mapper.js`

smoke test 통과 후 실제 프로빙 로직 추가. 기존 `_collectRibbonItems`, `_probeAllDropdowns`의 하위 로직 재사용.

### Phase A-3: run()에 Step 6 통합
Step 5 (드롭다운 프로빙) 완료 후 Step 6 호출.

### Phase A-4: 편집 탭 단독 테스트 (재현 가능한 패턴)
`test-probe-shape-tab.js` (신규) — 전체 run 없이 도형 탭만 프로빙.

### Phase A-5: 전체 재실행
`node src/menu-mapper.js hwp` 전체 플로우. 도형 탭이 채워지는지 확인.

---

## 5. 성공 기준

| 항목 | 기준 |
|------|------|
| 도형 탭 프로빙 성공 | `map.tabs['도형'].ribbonItems.length > 0` |
| 드롭다운 프로빙 | 도형 탭 내 드롭다운 ≥ 1개 수집 |
| 상태 오염 | 0건 (모든 작업 Ctrl+Z로 롤백) |
| 전체 실행 시간 증가 | ≤ 2분 (기존 15분 + 2분 이내) |
| 에러 | 0건 (실패 시 스킵 + 다음 진행) |

---

## 6. Phase B/C 예고 (구현 안 함)

- **Phase B (표)**: 표 삽입 → 셀 선택 → 프로빙 → 표 삭제
- **Phase C (그림)**: 클립보드에서 이미지 붙여넣기 → 선택 → 프로빙 → 삭제
- **Phase D (수식)**: 수식 편집기 열기 → 빈 수식 생성 → 프로빙 → 삭제

Phase A의 인프라(`_insertX`, `_probeContextTab`, `_cleanupX`) 재사용으로 각 Phase는 짧아질 것.

---

## 7. 선행 확인 사항

- [ ] HWP 실행 중 + 새 문서 1 빈 상태
- [ ] 입력 탭의 "도형 : ALT+P" 드롭다운이 map에 존재 (✅ 확인됨)
- [ ] 드롭다운 내 "직사각형" 등 기본 도형 아이템 존재 (Discovery로 확인 필요)
