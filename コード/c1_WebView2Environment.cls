VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c1_WebView2Environment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Standard Class: WebView2Environment
Option Explicit

Public Event ControllerCreationCompleted()

Private vTable(3) As LongPtr
Private pVTable As LongPtr
Private pControllerHandler As LongPtr

Public Sub CreateWebView2Controller(ByVal pEnvironment As LongPtr)
    
    ' 1. Retrieve the latest function pointers (required every time to handle re-compilation)
    vTable(0) = GetAddr(AddressOf Handler_QueryInterface)
    vTable(1) = GetAddr(AddressOf Handler_AddRef)
    vTable(2) = GetAddr(AddressOf Handler_Release)
    vTable(3) = GetAddr(AddressOf ControllerHandler_Invoke)

    ' 2. Allocate heap memory for the custom COM object if not already allocated
    If pVTable = 0 Then
        pVTable = GlobalAlloc(GPTR, (UBound(vTable) + 1) * LenB(vTable(0)))
        pControllerHandler = GlobalAlloc(GPTR, LenB(pVTable))
        
        ' Point the Object instance (pControllerHandler) to its VTable (pVTable). This only needs to be done once.
        CopyMemory ByVal pControllerHandler, pVTable, LenB(pVTable)
    End If

    ' 3. Synchronize the VTable content with the latest function pointers.
    ' This ensures it works correctly even if the VBA functions shifted memory addresses during a reset/re-compile.
    CopyMemory ByVal pVTable, vTable(0), (UBound(vTable) + 1) * LenB(vTable(0))

    ' 4. Invoke the API via DispCallFunc to create the WebView2 Controller
    Dim hr As Long
    hr = CallCreateController(pEnvironment, TargetHwnd, pControllerHandler)
    
    Debug.Print "CreateCoreWebView2Controller Call HRESULT: " & hr
End Sub

Public Sub OnControllerCreationCompleted()
    RaiseEvent ControllerCreationCompleted
End Sub
