# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

한컴오피스 2027 UI 테스트 자동화 MCP 서버. Windows UI Automation(UIA) COM 인터페이스를 koffi로 직접 호출하여 한컴 앱의 UI 요소를 탐색하고 조작한다. MCP(Model Context Protocol) stdio 서버로 동작하며, AI 에이전트가 22개 도구를 통해 4개 제품(한글, 한워드, 한쇼, 한셀)을 제어할 수 있다.

## Commands

```bash
npm start              # MCP 서버 실행 (stdio transport)
npm run test:hwp       # 한글 실행 + UIA 조작 통합 테스트 (node test-hwp.js)
npm install            # 의존성 설치 (koffi 네이티브 빌드 포함)
```

## Architecture

### Layer Structure

```
mcp-server.js          MCP 도구 등록 + stdio transport (진입점)
  └─ hwp-controller.js   고수준 자동화 API (launch/attach/close/UI탐색/인터랙션)
       ├─ uia.js            Windows UIA COM vtable 직접 호출 래퍼 (UIAutomation, UIElement 클래스)
       ├─ win32.js           Win32 API 래퍼 (SendInput, 키보드, 윈도우 관리)
       └─ session.js         싱글톤 세션 상태 (프로세스/윈도우/캐시 관리)
  ├─ action-logger.js    도구 실행 로깅 + 크래시 감지 데코레이터
  ├─ crash-monitor.js    프로세스 하트비트 + 크래시 유형 판별
  └─ jira-reporter.js    Jira REST API 버그 자동 리포팅 (ADF 포맷)
```

### Key Design Decisions

- **COM vtable 직접 호출**: `uia.js`에서 koffi로 IUIAutomation/IUIAutomationElement의 vtable 인덱스를 직접 참조하여 메서드 호출. vtable 인덱스 상수(`VTABLE`)가 COM 인터페이스 버전에 종속됨.
- **Lazy-require 패턴**: `mcp-server.js`에서 모듈을 함수로 감싸 지연 로딩 (`getController()`, `getCrashMonitor()` 등). 개별 모듈 실패 시에도 서버 시작 가능.
- **싱글톤 세션**: `session.js`의 모듈-레벨 객체가 유일한 세션 상태. `initSession(options)` 객체 형태로 초기화하며 세션 객체를 반환. `endSession()`으로 정리. UIAutomation 인스턴스는 세션 간 재사용.
- **멀티 제품 지원**: `session.js`의 `PRODUCTS` 맵에 4개 제품(hwp/hword/hshow/hcell) 경로와 이름 정의. `launch()`/`attach()`의 `product` 파라미터로 선택.
- **검색 결과 캐싱**: `findElement()`의 결과가 `_cachedSearchResults`에 UIElement 포인터로 캐시되며, `clickElement(index)`로 참조. 새 검색 시 이전 캐시 해제 필수.
- **koffi 메모리 관리**: UIElement는 사용 후 반드시 `.release()` 호출. 누락 시 COM 객체 누수 발생.

### MCP Tools (22개)

- **세션**: hwp_launch, hwp_attach, hwp_close, hwp_status
- **UI 탐색**: hwp_get_ui_tree, hwp_find_element, hwp_get_window_info, hwp_get_focused_element
- **인터랙션**: hwp_click_menu, hwp_click_button, hwp_click_element, hwp_type_text, hwp_press_keys, hwp_handle_dialog, hwp_set_foreground, hwp_take_screenshot (미구현)
- **로깅**: hwp_get_action_log, hwp_export_report, hwp_clear_log
- **버그 리포트**: hwp_report_bug, hwp_get_crash_history, hwp_configure_jira

## Environment Requirements

- **Windows x64 전용** (koffi COM 바인딩, user32/kernel32 DLL 직접 호출)
- Node.js 18+ (built-in fetch 사용)
- 한컴오피스 2027 설치 필요 (기본 경로: `C:\Program Files (x86)\Hnc\Office 2027\HOffice140\Bin\`)
  - 지원 제품: `Hwp.exe`(한글), `Hword.exe`(한워드), `HShow.exe`(한쇼), `HCell.exe`(한셀)
- Jira 연동 시 `.env` 파일 필요 (`.env.example` 참조)

## Important Conventions

- 한컴오피스 메인 윈도우 클래스명은 `FrameWindowImpl` — 4개 제품 모두 동일. 윈도우 탐색 시 이 클래스명으로 필터링
- 한글 런처(시작 화면)는 ESC 키로 닫음. SendInput 실패 시 PostMessage 폴백 사용
- 키 표현식은 `"Ctrl+S"`, `"Alt+F4"`, `"F5"` 형식 (`win32.parseKeyExpression()`)
- 한글/유니코드 텍스트 입력은 반드시 `useClipboard: true` (클립보드+Ctrl+V 방식)
- `mcp-server.js`의 로깅 도구(hwp_get_action_log, hwp_export_report, hwp_clear_log)는 `withLogging` 래퍼로 감싸지 않음 (재귀 방지)
