/**
 * Windows UI Automation (UIA) wrapper via koffi
 * COM vtable 직접 호출 방식으로 pywinauto와 동등한 기능 제공
 */
const koffi = require('koffi');

// === DLL 로드 ===
const ole32 = koffi.load('ole32.dll');
const oleaut32 = koffi.load('oleaut32.dll');

// === 기본 타입 ===
const GUID = koffi.struct('GUID', {
    Data1: 'uint32',
    Data2: 'uint16',
    Data3: 'uint16',
    Data4: koffi.array('uint8', 8)
});

const PTR_SIZE = 8; // x64

// === COM 기본 함수 ===
const CoInitializeEx = ole32.func('int32 __stdcall CoInitializeEx(void *, uint32)');
const CoCreateInstance = ole32.func('int32 __stdcall CoCreateInstance(GUID *, void *, uint32, GUID *, _Out_ void **)');
const CoUninitialize = ole32.func('void __stdcall CoUninitialize()');
const SysFreeString = oleaut32.func('void __stdcall SysFreeString(void *)');
const SysStringLen = oleaut32.func('uint32 __stdcall SysStringLen(void *)');

// === COM vtable 메서드 프로토타입 ===
const proto = {
    Release: koffi.proto('uint32 __stdcall p_Release(void *)'),
    // IUIAutomation
    GetRootElement: koffi.proto('int32 __stdcall p_GetRootElement(void *, _Out_ void **)'),
    ElementFromHandle: koffi.proto('int32 __stdcall p_ElementFromHandle(void *, void *, _Out_ void **)'),
    GetFocusedElement: koffi.proto('int32 __stdcall p_GetFocusedElement(void *, _Out_ void **)'),
    CreateTrueCondition: koffi.proto('int32 __stdcall p_CreateTrueCondition(void *, _Out_ void **)'),
    CreatePropertyCondition: koffi.proto('int32 __stdcall p_CreatePropertyCondition(void *, int32, int64, _Out_ void **)'),
    CreateAndCondition: koffi.proto('int32 __stdcall p_CreateAndCondition(void *, void *, void *, _Out_ void **)'),
    // IUIAutomationElement
    FindFirst: koffi.proto('int32 __stdcall p_FindFirst(void *, int32, void *, _Out_ void **)'),
    FindAll: koffi.proto('int32 __stdcall p_FindAll(void *, int32, void *, _Out_ void **)'),
    SetFocus: koffi.proto('int32 __stdcall p_SetFocus(void *)'),
    GetInt32Prop: koffi.proto('int32 __stdcall p_GetInt32Prop(void *, _Out_ int32 *)'),
    GetBstrProp: koffi.proto('int32 __stdcall p_GetBstrProp(void *, void *)'), // Buffer 방식
    // IUIAutomationElementArray
    GetLength: koffi.proto('int32 __stdcall p_GetLength(void *, _Out_ int32 *)'),
    GetElement: koffi.proto('int32 __stdcall p_GetElement(void *, int32, _Out_ void **)'),
    // InvokePattern
    Invoke: koffi.proto('int32 __stdcall p_Invoke(void *)'),
    // ExpandCollapsePattern
    Expand: koffi.proto('int32 __stdcall p_Expand(void *)'),
    Collapse: koffi.proto('int32 __stdcall p_Collapse(void *)'),
    // ValuePattern
    SetValue: koffi.proto('int32 __stdcall p_SetValue(void *, str16)'),
    // GetCurrentPattern
    GetCurrentPattern: koffi.proto('int32 __stdcall p_GetCurrentPattern(void *, int32, _Out_ void **)'),
};

// === vtable 인덱스 ===
const VTABLE = {
    IUIAutomation: {
        Release: 2,
        GetRootElement: 5,
        ElementFromHandle: 6,
        GetFocusedElement: 8,
        CreateTrueCondition: 17,
        CreatePropertyCondition: 19,
        CreateAndCondition: 21,
    },
    IUIAutomationElement: {
        Release: 2,
        SetFocus: 3,
        FindFirst: 5,
        FindAll: 6,
        GetCurrentPattern: 16,
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
        get_CurrentNativeWindowHandle: 36,
    },
    IUIAutomationElementArray: {
        Release: 2,
        get_Length: 3,
        GetElement: 4,
    },
    IInvokePattern: {
        Release: 2,
        Invoke: 3,
    },
    IExpandCollapsePattern: {
        Release: 2,
        Expand: 3,
        Collapse: 4,
    },
    IValuePattern: {
        Release: 2,
        SetValue: 3,
    },
};

// === UIA 상수 ===
const TreeScope = { Element: 0x1, Children: 0x2, Descendants: 0x4, Subtree: 0x7 };

const ControlTypeId = {
    Button: 50000, Calendar: 50001, CheckBox: 50002, ComboBox: 50003,
    Edit: 50004, Hyperlink: 50005, Image: 50006, ListItem: 50007,
    List: 50008, Menu: 50009, MenuBar: 50010, MenuItem: 50011,
    ProgressBar: 50012, RadioButton: 50013, ScrollBar: 50014,
    Slider: 50015, Spinner: 50016, StatusBar: 50017, Tab: 50018,
    TabItem: 50019, Text: 50020, ToolBar: 50021, ToolTip: 50025,
    Tree: 50026, TreeItem: 50027, Window: 50032, Pane: 50033,
    Custom: 50040,
};

const ControlTypeName = Object.fromEntries(
    Object.entries(ControlTypeId).map(([k, v]) => [v, k])
);

const PatternId = {
    Invoke: 10000,
    Selection: 10001,
    Value: 10002,
    RangeValue: 10003,
    Scroll: 10004,
    ExpandCollapse: 10005,
    Grid: 10006,
    Toggle: 10015,
};

const PropertyId = {
    Name: 30005,
    ControlType: 30003,
    AutomationId: 30011,
    ClassName: 30012,
    IsEnabled: 30010,
};

// === COM vtable 호출 헬퍼 ===
function comCall(pInterface, vtableIndex, fnProto, ...args) {
    const vtablePtr = koffi.decode(pInterface, 'void *');
    const fnPtr = koffi.decode(vtablePtr, vtableIndex * PTR_SIZE, 'void *');
    const fn = koffi.decode(fnPtr, fnProto);
    return fn(pInterface, ...args);
}

function comRelease(pInterface) {
    if (pInterface) comCall(pInterface, 2, proto.Release);
}

// === BSTR 변환 ===
function bstrToString(bstr) {
    if (!bstr) return '';
    try {
        const len = SysStringLen(bstr);
        if (len === 0) { SysFreeString(bstr); return ''; }
        const chars = koffi.decode(bstr, koffi.array('uint16', len));
        SysFreeString(bstr);
        return String.fromCharCode(...chars);
    } catch (e) {
        try { SysFreeString(bstr); } catch (_) {}
        return '';
    }
}

function readBstr(pElement, vtableIndex) {
    const buf = Buffer.alloc(PTR_SIZE);
    comCall(pElement, vtableIndex, proto.GetBstrProp, buf);
    const addr = buf.readBigUInt64LE(0);
    if (addr === 0n) return '';
    return bstrToString(koffi.decode(buf, 'void *'));
}

function readInt32(pElement, vtableIndex) {
    const out = [0];
    comCall(pElement, vtableIndex, proto.GetInt32Prop, out);
    return out[0];
}

// ===========================================================
// UIAutomation 클래스
// ===========================================================
class UIAutomation {
    constructor() {
        this._pAutomation = null;
        this._initialized = false;
    }

    init() {
        if (this._initialized) return;
        const hr1 = CoInitializeEx(null, 0x2); // COINIT_APARTMENTTHREADED
        if (hr1 !== 0 && hr1 !== 1) { // S_OK or S_FALSE (already initialized)
            throw new Error(`CoInitializeEx failed: 0x${(hr1 >>> 0).toString(16)}`);
        }

        const CLSID = { Data1: 0xFF48DBA4, Data2: 0x60EF, Data3: 0x4201, Data4: [0xAA, 0x87, 0x54, 0x10, 0x3E, 0xEF, 0x59, 0x4E] };
        const IID = { Data1: 0x30CBE57D, Data2: 0xD9D0, Data3: 0x452A, Data4: [0xAB, 0x13, 0x7A, 0xC5, 0xAC, 0x48, 0x25, 0xEE] };

        const pp = [null];
        const hr2 = CoCreateInstance(CLSID, null, 1, IID, pp);
        if (hr2 !== 0) throw new Error(`CoCreateInstance failed: 0x${(hr2 >>> 0).toString(16)}`);

        this._pAutomation = pp[0];
        this._initialized = true;
    }

    destroy() {
        if (this._pAutomation) {
            comRelease(this._pAutomation);
            this._pAutomation = null;
        }
        CoUninitialize();
        this._initialized = false;
    }

    getRootElement() {
        const pp = [null];
        const hr = comCall(this._pAutomation, VTABLE.IUIAutomation.GetRootElement, proto.GetRootElement, pp);
        if (hr !== 0 || !pp[0]) return null;
        return new UIElement(pp[0], this);
    }

    elementFromHandle(hwnd) {
        const pp = [null];
        const hr = comCall(this._pAutomation, VTABLE.IUIAutomation.ElementFromHandle, proto.ElementFromHandle, hwnd, pp);
        if (hr !== 0 || !pp[0]) return null;
        return new UIElement(pp[0], this);
    }

    getFocusedElement() {
        const pp = [null];
        const hr = comCall(this._pAutomation, VTABLE.IUIAutomation.GetFocusedElement, proto.GetFocusedElement, pp);
        if (hr !== 0 || !pp[0]) return null;
        return new UIElement(pp[0], this);
    }

    createTrueCondition() {
        const pp = [null];
        comCall(this._pAutomation, VTABLE.IUIAutomation.CreateTrueCondition, proto.CreateTrueCondition, pp);
        return pp[0];
    }
}

// ===========================================================
// UIElement 클래스
// ===========================================================
class UIElement {
    constructor(pElement, automation) {
        this._ptr = pElement;
        this._uia = automation;
    }

    release() {
        if (this._ptr) { comRelease(this._ptr); this._ptr = null; }
    }

    // --- 속성 읽기 ---
    get name() { return readBstr(this._ptr, VTABLE.IUIAutomationElement.get_CurrentName); }
    get controlType() { return readInt32(this._ptr, VTABLE.IUIAutomationElement.get_CurrentControlType); }
    get controlTypeName() { return ControlTypeName[this.controlType] || `Unknown(${this.controlType})`; }
    get className() { return readBstr(this._ptr, VTABLE.IUIAutomationElement.get_CurrentClassName); }
    get automationId() { return readBstr(this._ptr, VTABLE.IUIAutomationElement.get_CurrentAutomationId); }
    get processId() { return readInt32(this._ptr, VTABLE.IUIAutomationElement.get_CurrentProcessId); }
    get isEnabled() { return readInt32(this._ptr, VTABLE.IUIAutomationElement.get_CurrentIsEnabled) !== 0; }
    get acceleratorKey() { return readBstr(this._ptr, VTABLE.IUIAutomationElement.get_CurrentAcceleratorKey); }
    get localizedControlType() { return readBstr(this._ptr, VTABLE.IUIAutomationElement.get_CurrentLocalizedControlType); }

    get nativeWindowHandle() {
        const buf = Buffer.alloc(PTR_SIZE);
        comCall(this._ptr, VTABLE.IUIAutomationElement.get_CurrentNativeWindowHandle, proto.GetBstrProp, buf);
        return buf.readBigUInt64LE(0);
    }

    // --- 요소 정보 객체 ---
    toInfo() {
        return {
            name: this.name,
            controlType: this.controlTypeName,
            className: this.className,
            automationId: this.automationId,
            isEnabled: this.isEnabled,
        };
    }

    // --- 검색 ---
    findAll(scope = TreeScope.Children) {
        const cond = this._uia.createTrueCondition();
        if (!cond) return [];
        const ppArray = [null];
        const hr = comCall(this._ptr, VTABLE.IUIAutomationElement.FindAll, proto.FindAll, scope, cond, ppArray);
        comRelease(cond);
        if (hr !== 0 || !ppArray[0]) return [];
        return this._arrayToElements(ppArray[0]);
    }

    findFirst(scope, condition) {
        const pp = [null];
        const hr = comCall(this._ptr, VTABLE.IUIAutomationElement.FindFirst, proto.FindFirst, scope, condition, pp);
        if (hr !== 0 || !pp[0]) return null;
        return new UIElement(pp[0], this._uia);
    }

    findAllChildren() { return this.findAll(TreeScope.Children); }
    findAllDescendants() { return this.findAll(TreeScope.Descendants); }

    /**
     * 이름 또는 컨트롤 타입으로 자식 요소 검색
     */
    findByName(name, scope = TreeScope.Descendants) {
        const children = this.findAll(scope);
        const results = [];
        for (const child of children) {
            if (child.name === name) {
                results.push(child);
            } else {
                child.release();
            }
        }
        return results;
    }

    findByControlType(controlType, scope = TreeScope.Descendants) {
        const ctId = typeof controlType === 'string' ? ControlTypeId[controlType] : controlType;
        const children = this.findAll(scope);
        const results = [];
        for (const child of children) {
            if (child.controlType === ctId) {
                results.push(child);
            } else {
                child.release();
            }
        }
        return results;
    }

    // --- 액션 ---
    setFocus() {
        comCall(this._ptr, VTABLE.IUIAutomationElement.SetFocus, proto.SetFocus);
    }

    invoke() {
        const pp = [null];
        const hr = comCall(this._ptr, VTABLE.IUIAutomationElement.GetCurrentPattern, proto.GetCurrentPattern, PatternId.Invoke, pp);
        if (hr === 0 && pp[0]) {
            comCall(pp[0], VTABLE.IInvokePattern.Invoke, proto.Invoke);
            comRelease(pp[0]);
            return true;
        }
        return false;
    }

    expand() {
        const pp = [null];
        const hr = comCall(this._ptr, VTABLE.IUIAutomationElement.GetCurrentPattern, proto.GetCurrentPattern, PatternId.ExpandCollapse, pp);
        if (hr === 0 && pp[0]) {
            comCall(pp[0], VTABLE.IExpandCollapsePattern.Expand, proto.Expand);
            comRelease(pp[0]);
            return true;
        }
        return false;
    }

    collapse() {
        const pp = [null];
        const hr = comCall(this._ptr, VTABLE.IUIAutomationElement.GetCurrentPattern, proto.GetCurrentPattern, PatternId.ExpandCollapse, pp);
        if (hr === 0 && pp[0]) {
            comCall(pp[0], VTABLE.IExpandCollapsePattern.Collapse, proto.Collapse);
            comRelease(pp[0]);
            return true;
        }
        return false;
    }

    // --- 내부 헬퍼 ---
    _arrayToElements(pArray) {
        const lenOut = [0];
        comCall(pArray, VTABLE.IUIAutomationElementArray.get_Length, proto.GetLength, lenOut);
        const elements = [];
        for (let i = 0; i < lenOut[0]; i++) {
            const pp = [null];
            comCall(pArray, VTABLE.IUIAutomationElementArray.GetElement, proto.GetElement, i, pp);
            if (pp[0]) elements.push(new UIElement(pp[0], this._uia));
        }
        comRelease(pArray);
        return elements;
    }
}

module.exports = {
    UIAutomation,
    UIElement,
    TreeScope,
    ControlTypeId,
    ControlTypeName,
    PatternId,
    PropertyId,
};
