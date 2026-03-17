Attribute VB_Name = "Module1"
'標準モジュール Module1

Option Explicit

Public myWidth As Long
Public myHeight As Long
Public TargetHwnd As LongPtr

Public Sub フォーム表示()
    UserForm1.Show
End Sub


' DispCallFuncをラップする関数 (可変引数はVBAでは難しいため、今回は引数固定で実装)
Public Function CallCreateController(ByVal pEnv As LongPtr, ByVal hWnd As LongPtr, ByVal pHandler As LongPtr) As Long
    Dim vTable As LongPtr
    Dim pFunc As LongPtr
    Dim hr As Long
    Dim args(2) As Variant 'LongPtr
    Dim argTypes(2) As Integer
    Dim argPtrs(2) As LongPtr
    Dim result As Variant

    ' ICoreWebView2Environment::CreateCoreWebView2Controller は vTable の index 3
    ' (QueryInterface, AddRef, Release が 0,1,2 なので)

    args(0) = hWnd
    args(1) = pHandler

    argTypes(0) = vbLongLong
    argTypes(1) = vbLongLong

    argPtrs(0) = VarPtr(args(0))
    argPtrs(1) = VarPtr(args(1))

    CallCreateController = DispCallFunc(pEnv, 3 * LenB(pEnv), CC_STDCALL, vbLong, 2, argTypes(0), argPtrs(0), result)
End Function

Public Function CallAddRef(ByVal pUnk As LongPtr) As Long
    Dim res As Variant
    ' Index 1: IUnknown::AddRef
    ' 引数なし、戻り値は新しい参照カウント（Long）
    CallAddRef = DispCallFunc(pUnk, 1 * LenB(pUnk), CC_STDCALL, vbLong, 0, 0, 0, res)
End Function


' サイズ調整関連
Public Sub サイズ調整()

    ResizeWebView2Force TargetHwnd, PtsToPx(UserForm1.Frame1.InsideWidth, False), PtsToPx(UserForm1.Frame1.InsideHeight, True)

End Sub

Public Sub ResizeWebView2Force(ByVal parentHwnd As LongPtr, ByVal w As Long, ByVal h As Long)
    Dim child0 As LongPtr
    Dim child1 As LongPtr
    Dim childD3D As LongPtr
    
    ' 1. Chrome_WidgetWin_0 (器)
    child0 = FindWindowEx(parentHwnd, 0, "Chrome_WidgetWin_0", vbNullString)
    If child0 = 0 Then Exit Sub
    MoveWindow child0, 0, 0, w, h, 1
    
    ' 2. Chrome_WidgetWin_1 (レンダラ)
    child1 = FindWindowEx(child0, 0, "Chrome_WidgetWin_1", vbNullString)
    If child1 <> 0 Then
        MoveWindow child1, 0, 0, w, h, 1
        
        ' 3. Intermediate D3D Window (実際の描画面)
        childD3D = FindWindowEx(child1, 0, "Intermediate D3D Window", vbNullString)
        If childD3D <> 0 Then
            MoveWindow childD3D, 0, 0, w, h, 1
        End If
    End If
    
    Debug.Print "Win32 Recursive Resize Done."
End Sub

Public Function PtsToPx(ByVal pts As Single, ByVal isVertical As Boolean) As Long
    Dim hdc As LongPtr
    Dim dpi As Long
    
    ' デスクトップ全体のデバイスコンテキストを取得して現在のDPIを調査
    hdc = GetDC(0)
    If isVertical Then
        dpi = GetDeviceCaps(hdc, LOGPIXELSY)
    Else
        dpi = GetDeviceCaps(hdc, LOGPIXELSX)
    End If
    ReleaseDC 0, hdc
    
    ' 計算式: ポイント * (現在のDPI / 72)
    ' 150%設定なら、dpiには 144 が入るので、pts * 2 となり正しく変換される
    PtsToPx = pts * (dpi / 72)
End Function


'全く使ってないのに、このプロシージャを消したり、大部分をコメントアウトすると
'WebView2初期化処理の際、c0_WebView2Loader内のCreateCoreWebView2EnvironmentWithOptions
'を呼び出した瞬間にExcelがクラッシュする
Public Sub RegisterNavigationCompleted_()
    Static vTable As LongPtr
    vTable = GetAddr(AddressOf Handler_QueryInterface)
End Sub
