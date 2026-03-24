Attribute VB_Name = "Module1"
' Standard Module: Module1
Option Explicit

Public myWidth As Long
Public myHeight As Long
Public TargetHwnd As LongPtr

Public Sub ShowUserForm()
    UserForm1.Show
End Sub

' ---------------------------------------------------------
'  COM (DispCallFunc) Layer
' ---------------------------------------------------------

' Wraps DispCallFunc to call CreateCoreWebView2Controller (ICoreWebView2Environment Index 3)
Public Function CallCreateController(ByVal pEnv As LongPtr, ByVal hwnd As LongPtr, ByVal pHandler As LongPtr) As Long
    Dim vTable As LongPtr
    Dim pFunc As LongPtr
    Dim hr As Long
    Dim args(2) As Variant
    Dim argTypes(2) As Integer
    Dim argPtrs(2) As LongPtr
    Dim result As Variant

    ' ICoreWebView2Environment::CreateCoreWebView2Controller is VTable Index 3
    ' (QueryInterface=0, AddRef=1, Release=2)

    args(0) = hwnd
    args(1) = pHandler

    ' Explicitly set as 64-bit LongPtr compatibility
    argTypes(0) = vbLongLong
    argTypes(1) = vbLongLong

    argPtrs(0) = VarPtr(args(0))
    argPtrs(1) = VarPtr(args(1))

    CallCreateController = DispCallFunc(pEnv, 3 * LenB(pEnv), CC_STDCALL, vbLong, 2, argTypes(0), argPtrs(0), result)
End Function

' Calls IUnknown::AddRef (VTable Index 1) to increment the reference count
Public Function CallAddRef(ByVal pUnk As LongPtr) As Long
    Dim res As Variant
    CallAddRef = DispCallFunc(pUnk, 1 * LenB(pUnk), CC_STDCALL, vbLong, 0, 0, 0, res)
End Function


' ---------------------------------------------------------
'  Win32 Window Sizing & Layout Layer
' ---------------------------------------------------------

' Resizes the WebView2 window to match the inner dimensions of the UserForm Frame
Public Sub Resize()
    ResizeWebView2Force TargetHwnd, PtsToPx(UserForm1.Frame1.InsideWidth, False), PtsToPx(UserForm1.Frame1.InsideHeight, True)
End Sub

' Recursively traverses Chromium's internal Win32 hierarchy to forcefully resize drawing canvases
Public Sub ResizeWebView2Force(ByVal parentHwnd As LongPtr, ByVal w As Long, ByVal h As Long)
    Dim child0 As LongPtr
    Dim child1 As LongPtr
    Dim childD3D As LongPtr
    
    ' 1. Chrome_WidgetWin_0 (Outer Shell)
    child0 = FindWindowEx(parentHwnd, 0, "Chrome_WidgetWin_0", vbNullString)
    If child0 = 0 Then Exit Sub
    MoveWindow child0, 0, 0, w, h, 1
    
    ' 2. Chrome_WidgetWin_1 (Renderer Canvas)
    child1 = FindWindowEx(child0, 0, "Chrome_WidgetWin_1", vbNullString)
    If child1 <> 0 Then
        MoveWindow child1, 0, 0, w, h, 1
        
        ' 3. Intermediate D3D Window (Direct3D Drawing Surface)
        childD3D = FindWindowEx(child1, 0, "Intermediate D3D Window", vbNullString)
        If childD3D <> 0 Then
            MoveWindow childD3D, 0, 0, w, h, 1
        End If
    End If
    
    Debug.Print "Win32 Recursive Resize Completed."
End Sub


' ---------------------------------------------------------
'  DPI / GDI Coordinate Conversion Layer
' ---------------------------------------------------------

' Converts VBA Points (twips/points) to Pixel coordinates based on monitor DPI scaling
Public Function PtsToPx(ByVal pts As Single, ByVal isVertical As Boolean) As Long
    Dim hdc As LongPtr
    Dim dpi As Long
    
    hdc = GetDC(0) ' Get DC for the primary desktop monitor
    If isVertical Then
        dpi = GetDeviceCaps(hdc, LOGPIXELSY)
    Else
        dpi = GetDeviceCaps(hdc, LOGPIXELSX)
    End If
    ReleaseDC 0, hdc
    
    ' Formula: Points * (Current DPI / 72)
    ' e.g., 150% scaling will yield dpi = 144, resulting in points * 2
    PtsToPx = pts * (dpi / 72)
End Function

' Dummy Sub to prevent VBA compiler optimization / address loss
Public Sub RegisterNavigationCompleted_()
    Static vTable As LongPtr
    vTable = GetAddr(AddressOf Handler_QueryInterface)
End Sub
