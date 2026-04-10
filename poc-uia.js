/**
 * PoC: koffi를 통한 Windows UI Automation COM API 직접 호출
 * pywinauto와 동일한 기술을 Node.js에서 구현
 */
const koffi = require('koffi');

// === DLL 로드 ===
const ole32 = koffi.load('ole32.dll');
const oleaut32 = koffi.load('oleaut32.dll');
const user32 = koffi.load('user32.dll');

// === 타입 정의 ===
const GUID = koffi.struct('GUID', {
    Data1: 'uint32',
    Data2: 'uint16',
    Data3: 'uint16',
    Data4: koffi.array('uint8', 8)
});

// === COM 기본 함수 ===
const CoInitializeEx = ole32.func('int32 __stdcall CoInitializeEx(void *, uint32)');
const CoCreateInstance = ole32.func('int32 __stdcall CoCreateInstance(GUID *, void *, uint32, GUID *, _Out_ void **)');
const CoUninitialize = ole32.func('void __stdcall CoUninitialize()');
const SysFreeString = oleaut32.func('void __stdcall SysFreeString(void *)');
const SysStringLen = oleaut32.func('uint32 __stdcall SysStringLen(void *)');

// === user32 함수 (윈도우 관리용) ===
const FindWindowW = user32.func('void * __stdcall FindWindowW(str16 *, str16 *)');
const IsWindow = user32.func('int32 __stdcall IsWindow(void *)');
const IsHungAppWindow = user32.func('int32 __stdcall IsHungAppWindow(void *)');

// === COM vtable 호출 헬퍼 ===
const PTR_SIZE = 8; // x64

// COM 메서드 프로토타입 정의
const proto_QueryInterface = koffi.proto('int32 __stdcall p_QI(void *, GUID *, _Out_ void **)');
const proto_AddRef = koffi.proto('uint32 __stdcall p_AddRef(void *)');
const proto_Release = koffi.proto('uint32 __stdcall p_Release(void *)');

// IUIAutomation 메서드
const proto_GetRootElement = koffi.proto('int32 __stdcall p_GetRootElement(void *, _Out_ void **)');
const proto_ElementFromHandle = koffi.proto('int32 __stdcall p_ElementFromHandle(void *, void *, _Out_ void **)');
const proto_GetFocusedElement = koffi.proto('int32 __stdcall p_GetFocusedElement(void *, _Out_ void **)');
const proto_CreatePropertyCondition = koffi.proto('int32 __stdcall p_CreatePropertyCondition(void *, int32, void *, _Out_ void **)');
const proto_CreateTrueCondition = koffi.proto('int32 __stdcall p_CreateTrueCondition(void *, _Out_ void **)');

// IUIAutomationElement 메서드
const proto_FindFirst = koffi.proto('int32 __stdcall p_FindFirst(void *, int32, void *, _Out_ void **)');
const proto_FindAll = koffi.proto('int32 __stdcall p_FindAll(void *, int32, void *, _Out_ void **)');
const proto_GetCurrentName = koffi.proto('int32 __stdcall p_GetCurrentName(void *, _Out_ void **)');
const proto_GetCurrentControlType = koffi.proto('int32 __stdcall p_GetCurrentControlType(void *, _Out_ int32 *)');
const proto_GetCurrentClassName = koffi.proto('int32 __stdcall p_GetCurrentClassName(void *, _Out_ void **)');
const proto_GetCurrentProcessId = koffi.proto('int32 __stdcall p_GetCurrentProcessId(void *, _Out_ int32 *)');
const proto_SetFocus = koffi.proto('int32 __stdcall p_SetFocus(void *)');

// IUIAutomationElementArray 메서드
const proto_GetLength = koffi.proto('int32 __stdcall p_GetLength(void *, _Out_ int32 *)');
const proto_GetElement = koffi.proto('int32 __stdcall p_GetElement(void *, int32, _Out_ void **)');

/**
 * COM vtable에서 메서드를 가져와 호출하는 헬퍼
 * COM 인터페이스 메모리 레이아웃:
 *   pInterface -> [vtable_ptr] -> [fn0_ptr, fn1_ptr, fn2_ptr, ...]
 */
function comCall(pInterface, vtableIndex, proto, ...args) {
    // 1) vtable 포인터 읽기 (인터페이스의 첫 번째 포인터)
    const vtablePtr = koffi.decode(pInterface, 'void *');
    // 2) vtable에서 메서드 함수 포인터 읽기
    const fnPtr = koffi.decode(vtablePtr, vtableIndex * PTR_SIZE, 'void *');
    // 3) 함수 포인터를 callable로 변환
    const fn = koffi.decode(fnPtr, proto);
    return fn(pInterface, ...args);
}

/**
 * BSTR을 JavaScript 문자열로 변환
 * koffi.decode(ptr, 'str16')가 segfault를 유발하므로
 * SysStringLen + uint16 배열로 안전하게 읽음
 */
function bstrToString(bstr) {
    if (!bstr) return '';
    try {
        const len = SysStringLen(bstr);
        if (len === 0) {
            SysFreeString(bstr);
            return '';
        }
        const chars = koffi.decode(bstr, koffi.array('uint16', len));
        SysFreeString(bstr);
        return String.fromCharCode(...chars);
    } catch (e) {
        try { SysFreeString(bstr); } catch (_) {}
        return `(decode error: ${e.message})`;
    }
}

// === UIA ControlType 상수 ===
const UIA_ControlType = {
    50000: 'Button',
    50001: 'Calendar',
    50002: 'CheckBox',
    50003: 'ComboBox',
    50004: 'Edit',
    50005: 'Hyperlink',
    50006: 'Image',
    50007: 'ListItem',
    50008: 'List',
    50009: 'Menu',
    50010: 'MenuBar',
    50011: 'MenuItem',
    50012: 'ProgressBar',
    50013: 'RadioButton',
    50014: 'ScrollBar',
    50015: 'Slider',
    50016: 'Spinner',
    50017: 'StatusBar',
    50018: 'Tab',
    50019: 'TabItem',
    50020: 'Text',
    50021: 'ToolBar',
    50025: 'ToolTip',
    50026: 'Tree',
    50027: 'TreeItem',
    50032: 'Window',
    50033: 'Pane',
    50040: 'Custom',
};

// === IUIAutomation vtable indices ===
const UIA = {
    // IUnknown
    QueryInterface: 0,
    AddRef: 1,
    Release: 2,
    // IUIAutomation
    CompareElements: 3,
    CompareRuntimeIds: 4,
    GetRootElement: 5,
    ElementFromHandle: 6,
    ElementFromPoint: 7,
    GetFocusedElement: 8,
    CreateTreeWalker: 9,
    get_ControlViewWalker: 10,
    get_ContentViewWalker: 11,
    get_RawViewWalker: 12,
    get_RawViewCondition: 13,
    get_ControlViewCondition: 14,
    get_ContentViewCondition: 15,
    CreateCacheRequest: 16,
    CreateTrueCondition: 17,
    CreateFalseCondition: 18,
    CreatePropertyCondition: 19,
};

// === IUIAutomationElement vtable indices ===
// IUnknown(3) + IUIAutomationElement 메서드 순서 (UIAutomation.h 기준)
const ELEM = {
    QueryInterface: 0,
    AddRef: 1,
    Release: 2,
    SetFocus: 3,
    GetRuntimeId: 4,
    FindFirst: 5,
    FindAll: 6,
    FindFirstBuildCache: 7,
    FindAllBuildCache: 8,
    BuildUpdatedCache: 9,
    GetCurrentPropertyValue: 10,
    GetCurrentPropertyValueEx: 11,
    GetCachedPropertyValue: 12,
    GetCachedPropertyValueEx: 13,
    GetCurrentPatternAs: 14,
    GetCachedPatternAs: 15,
    GetCurrentPattern: 16,
    GetCachedPattern: 17,
    GetCachedParent: 18,
    GetCachedChildren: 19,
    get_CurrentProcessId: 20,
    get_CurrentControlType: 21,
    get_CurrentLocalizedControlType: 22,
    get_CurrentName: 23,
    get_CurrentAcceleratorKey: 24,
    get_CurrentAccessKey: 25,
    get_CurrentHasKeyboardFocus: 26,
    get_CurrentIsKeyboardFocusable: 27,
    get_CurrentIsEnabled: 28,
    get_CurrentAutomationId: 29,
    get_CurrentClassName: 30,
};

// === IUIAutomationElementArray vtable indices ===
const ARR = {
    QueryInterface: 0,
    AddRef: 1,
    Release: 2,
    get_Length: 3,
    GetElement: 4,
};

// TreeScope enum
const TreeScope = {
    Element: 0x1,
    Children: 0x2,
    Descendants: 0x4,
    Subtree: 0x7,
};

// =============================================
// 메인 PoC 실행
// =============================================
console.log('=== Windows UI Automation PoC via koffi ===\n');

// 1. COM 초기화
let hr = CoInitializeEx(null, 0x2); // COINIT_APARTMENTTHREADED
console.log(`[1] CoInitializeEx: 0x${(hr >>> 0).toString(16)}`);

// 2. IUIAutomation 인스턴스 생성
const CLSID_CUIAutomation = {
    Data1: 0xFF48DBA4, Data2: 0x60EF, Data3: 0x4201,
    Data4: [0xAA, 0x87, 0x54, 0x10, 0x3E, 0xEF, 0x59, 0x4E]
};
const IID_IUIAutomation = {
    Data1: 0x30CBE57D, Data2: 0xD9D0, Data3: 0x452A,
    Data4: [0xAB, 0x13, 0x7A, 0xC5, 0xAC, 0x48, 0x25, 0xEE]
};

const ppAutomation = [null];
hr = CoCreateInstance(CLSID_CUIAutomation, null, 1, IID_IUIAutomation, ppAutomation);
console.log(`[2] CoCreateInstance: 0x${(hr >>> 0).toString(16)}`);

if (hr !== 0) {
    console.error('Failed to create IUIAutomation');
    CoUninitialize();
    process.exit(1);
}
const pAutomation = ppAutomation[0];
console.log(`    pAutomation: [OK]\n`);

// 3. 데스크톱 루트 요소 가져오기
const ppRoot = [null];
hr = comCall(pAutomation, UIA.GetRootElement, proto_GetRootElement, ppRoot);
console.log(`[3] GetRootElement: 0x${(hr >>> 0).toString(16)}`);

if (hr === 0 && ppRoot[0]) {
    const pRoot = ppRoot[0];

    // 디버그: 먼저 ControlType 읽기 (int32 반환, BSTR보다 안전)
    console.log('    Trying get_CurrentControlType (index 21)...');
    const ctOut = [0];
    hr = comCall(pRoot, ELEM.get_CurrentControlType, proto_GetCurrentControlType, ctOut);
    console.log(`    ControlType hr=0x${(hr >>> 0).toString(16)} val=${ctOut[0]} (${UIA_ControlType[ctOut[0]] || 'unknown'})`);

    // ProcessId 읽기
    console.log('    Trying get_CurrentProcessId (index 20)...');
    const pidOut = [0];
    hr = comCall(pRoot, ELEM.get_CurrentProcessId, proto_GetCurrentProcessId, pidOut);
    console.log(`    ProcessId hr=0x${(hr >>> 0).toString(16)} val=${pidOut[0]}`);

    // Name 읽기 - Buffer에서 raw 주소 직접 추출
    console.log('    Trying get_CurrentName (index 23) with Buffer...');
    const proto_GetName_raw = koffi.proto('int32 __stdcall p_GetNameRaw(void *, void *)');
    const nameBuf = Buffer.alloc(PTR_SIZE);
    hr = comCall(pRoot, ELEM.get_CurrentName, proto_GetName_raw, nameBuf);
    console.log(`    Name hr=0x${(hr >>> 0).toString(16)}`);
    const bstrAddr = nameBuf.readBigUInt64LE(0);
    console.log(`    BSTR address: 0x${bstrAddr.toString(16)}`);
    if (bstrAddr !== 0n) {
        const bstrPtr = koffi.decode(nameBuf, 'void *');
        console.log(`    Name value: "${bstrToString(bstrPtr)}"`);
    } else {
        console.log('    Name is empty');
    }

    // 4. TrueCondition 생성 (모든 자식 검색용)
    const ppCondition = [null];
    hr = comCall(pAutomation, UIA.CreateTrueCondition, proto_CreateTrueCondition, ppCondition);
    console.log(`\n[4] CreateTrueCondition: 0x${(hr >>> 0).toString(16)}`);

    if (hr === 0 && ppCondition[0]) {
        // 5. 루트의 자식 윈도우들 검색
        const ppArray = [null];
        hr = comCall(pRoot, ELEM.FindAll, proto_FindAll, TreeScope.Children, ppCondition[0], ppArray);
        console.log(`[5] FindAll (children): 0x${(hr >>> 0).toString(16)}`);

        if (hr === 0 && ppArray[0]) {
            const pArray = ppArray[0];
            const pLength = [0];
            comCall(pArray, ARR.get_Length, proto_GetLength, pLength);
            console.log(`    Found ${pLength[0]} top-level windows\n`);

            // 각 윈도우 정보 출력 (최대 15개)
            const count = Math.min(pLength[0], 15);
            for (let i = 0; i < count; i++) {
                try {
                    const ppChild = [null];
                    comCall(pArray, ARR.GetElement, proto_GetElement, i, ppChild);

                    if (ppChild[0]) {
                        const ctOut = [0];
                        comCall(ppChild[0], ELEM.get_CurrentControlType, proto_GetCurrentControlType, ctOut);
                        const ctName = UIA_ControlType[ctOut[0]] || ctOut[0];

                        // BSTR은 Buffer 방식으로 읽기
                        const nameBuf2 = Buffer.alloc(PTR_SIZE);
                        comCall(ppChild[0], ELEM.get_CurrentName, proto_GetName_raw, nameBuf2);
                        const addr = nameBuf2.readBigUInt64LE(0);
                        let name = '';
                        if (addr !== 0n) {
                            const ptr = koffi.decode(nameBuf2, 'void *');
                            name = bstrToString(ptr);
                        }

                        console.log(`    [${ctName}] "${name}"`);
                        comCall(ppChild[0], 2, proto_Release);
                    }
                } catch (e) {
                    console.log(`    [${i}] Error: ${e.message}`);
                }
            }

            comCall(pArray, 2, proto_Release);
        }

        comCall(ppCondition[0], 2, proto_Release);
    }

    comCall(pRoot, 2, proto_Release);
}

// 정리
comCall(pAutomation, 2, proto_Release);
CoUninitialize();
console.log('\n=== PoC 완료 ===');
