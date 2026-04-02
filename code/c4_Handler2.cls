VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c4_Handler2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''Public Class c4_Handler2
'Option Explicit
'
'#If Win64 Then
'Public Event Invoked(ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr)
'
'' --- API Declarations (for x64) ---
'Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
'Private Declare PtrSafe Function VirtualFree Lib "kernel32" (ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, ByVal dwFreeType As Long) As Long
'Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As LongPtr, ByVal source As LongPtr, ByVal Length As LongPtr)
'
'' VirtualAlloc constants
'Private Const MEM_COMMIT As Long = &H1000
'Private Const MEM_RESERVE As Long = &H2000
'Private Const PAGE_EXECUTE_READWRITE As Long = &H40
'Private Const MEM_RELEASE As Long = &H8000
'
'' --- Internal Variables ---
'Private mVTablePtr As LongPtr      ' Address of the faked VTable (array of function pointers)
'Private mFakeObjPtr As LongPtr     ' Address that "impersonates" an object to pass to WebView2
'Private mThunkAddr As LongPtr      ' Address of the machine code thunk (translator)
'Public token As LongPtr            ' Token for unregistering WebView2
'
'' Allocation size constants
'Private Const VTABLE_SIZE As Long = 32 ' IUnknown(3) + Invoke(1) = 4 methods * 8 bytes
'Private Const THUNK_SIZE As Long = 32  ' Machine code space size
'
'Public ppWebView2 As LongPtr
'
'' --- Add to Class Internal Variables ---
'Private mThunkTimer As LongPtr     ' Address for timer-specific thunk
'Private m_TimerID As LongPtr       ' ID for stopping the timer
'
'' --- Public Exposure: Pointer to be passed to WebView2 ---
'Public Property Get Pointer() As LongPtr
'    Pointer = mFakeObjPtr
'End Property
'
'' ==========================================================
''  Main: Method invoked by WebView2
'' ==========================================================
'' Since registers are aligned by the thunk, arguments can be received as per the signature.
'' pThis: Faked object (mFakeObjPtr), sender: WebView2 instance, args: Event arguments
'Public Function cls_Invoke(ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
'    On Error Resume Next ' Fail-safe for external calls
'
'    RaiseEvent Invoked(pThis, sender, args)
''    Debug.Print "=== WebView2 Event Fired ==="
''    Debug.Print "Sender Ptr:", sender
''    Debug.Print "Args Ptr:", args
''    Debug.Print "ppWebView2:", ppWebView2
'    ' Safe to access properties of Me (this class instance) here.
'    ' e.g., Debug.Print Me.Token
'
'    cls_Invoke = 0 ' S_OK
'End Function
'' ==========================================================
''  Initialization: Memory Allocation and Machine Code Injection
'' ==========================================================
'Private Sub Class_Initialize()
'    ' 1. Allocate memory (VTable, Thunk, FakeInstance)
'    mVTablePtr = VirtualAlloc(0, VTABLE_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
'    mThunkAddr = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
'    mFakeObjPtr = VirtualAlloc(0, 8, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
'
'    ' 2. Retrieve the actual address of the class method "cls_Invoke"
'    Dim pTargetFunc As LongPtr, offset As Long
'    pTargetFunc = GetClassMethodPtr(Me, "cls_Invoke", offset) ' * Uses the existing utility function
'
'    ' 3. Create machine code thunk for x64
'    ' Details: Shift WebView2 arguments (RCX, RDX, R8) to the right (RDX, R8, R9),
'    '          set "Me" (this class instance) into RCX, and jump to the target function.
'    Dim thunkBytes(31) As Byte
'    Dim i As Long: i = 0
'
'    ' mov r9, r8   (args -> shift to 4th argument)
'    thunkBytes(i) = &H4D: thunkBytes(i + 1) = &H89: thunkBytes(i + 2) = &HC1: i = i + 3
'    ' mov r8, rdx  (sender -> shift to 3rd argument)
'    thunkBytes(i) = &H49: thunkBytes(i + 1) = &H89: thunkBytes(i + 2) = &HD0: i = i + 3
'    ' mov rdx, rcx (pThis -> shift to 2nd argument)
'    thunkBytes(i) = &H48: thunkBytes(i + 1) = &H89: thunkBytes(i + 2) = &HCA: i = i + 3
'    ' mov rcx, [ObjPtr(Me)] (Inject Me -> 1st argument)
'    thunkBytes(i) = &H48: thunkBytes(i + 1) = &HB9: i = i + 2
'    RtlMoveMemory VarPtr(thunkBytes(i)), VarPtr(ObjPtr(Me)), 8: i = i + 8
'    ' mov rax, [pTargetFunc] (Set jump target)
'    thunkBytes(i) = &H48: thunkBytes(i + 1) = &HB8: i = i + 2
'    RtlMoveMemory VarPtr(thunkBytes(i)), VarPtr(pTargetFunc), 8: i = i + 8
'    ' jmp rax (Transfer execution)
'    thunkBytes(i) = &HFF: thunkBytes(i + 1) = &HE0
'
'    RtlMoveMemory mThunkAddr, VarPtr(thunkBytes(0)), THUNK_SIZE
'
'    ' 4. Construct VTable (Index 0-2 are standard module dummies, 3 is the thunk)
'    ' * Assumes Handler_QI/AddRef/Release are defined in standard modules.
'    Dim pQI As LongPtr, pAddRef As LongPtr, pRelease As LongPtr
'    pQI = GetAddr(AddressOf Handler_QueryInterface)
'    pAddRef = GetAddr(AddressOf Handler_AddRef)
'    pRelease = GetAddr(AddressOf Handler_Release)
'
'    RtlMoveMemory mVTablePtr + (0 * 8), VarPtr(pQI), 8
'    RtlMoveMemory mVTablePtr + (1 * 8), VarPtr(pAddRef), 8
'    RtlMoveMemory mVTablePtr + (2 * 8), VarPtr(pRelease), 8
'    RtlMoveMemory mVTablePtr + (3 * 8), VarPtr(mThunkAddr), 8
'
'    ' 5. Write the VTable pointer at the beginning of the FakeInstance
'    RtlMoveMemory mFakeObjPtr, VarPtr(mVTablePtr), 8
'
'    Debug.Print "Handler Ready. FakeObjPtr: " & Hex(mFakeObjPtr)
'End Sub
'
'Private Sub Class_Terminate()
'    ' Release memory
''    If mVTablePtr <> 0 Then VirtualFree mVTablePtr, 0, MEM_RELEASE
''    If mThunkAddr <> 0 Then VirtualFree mThunkAddr, 0, MEM_RELEASE
''    If mFakeObjPtr <> 0 Then VirtualFree mFakeObjPtr, 0, MEM_RELEASE
''    Debug.Print "c4_Handler2 Memory Released."
'End Sub
'#End If
