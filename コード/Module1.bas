Attribute VB_Name = "Module1"
'標準モジュール Module1

Option Explicit

' 宣言

' 標準モジュールにて
Public Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long


' 定数
Public Const LOGPIXELSX As Long = 88 ' 横方向のDPI
Public Const LOGPIXELSY As Long = 90 ' 縦方向のDPI

Public myWidth As Long
Public myHeight As Long


' 寿命を永続化させるために Public かつ Static に近い扱いで保持
Public pKeepEnv As LongPtr        ' Environmentポインタ保存用
Public pKeepController As LongPtr   ' Controllerポインタ保存用
Public pKeepWebView As LongPtr      ' WebViewポインタ保存用

Public Declare PtrSafe Function CallWindowProcW Lib "user32" (ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, ByVal Msg As LongPtr, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

Public Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWndParent As LongPtr, _
    ByVal hWndChildAfter As LongPtr, _
    ByVal lpszClass As String, _
    ByVal lpszWindow As String) As LongPtr

Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Public Const SW_SHOW As Long = 5

Public Declare PtrSafe Function GetWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal uCmd As Long) As LongPtr
Public Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Const GW_CHILD As Long = 5

Public Declare PtrSafe Function SetWindowPos Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
    ByVal uFlags As Long) As Long

Public Const HWND_TOP As LongPtr = 0
Public Const SWP_SHOWWINDOW As Long = &H40

Public TargetHwnd As LongPtr

'' WebView2Loader.dll のエントリポイント
'Private Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" ( _
'    ByVal browserExecutableFolder As LongPtr, _
'    ByVal userDataFolder As LongPtr, _
'    ByVal additionalBrowserArguments As LongPtr, _
'    ByVal environmentCreatedHandler As LongPtr) As Long

' メモリ確保用
Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As LongPtr)
Public Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr

Public Const GPTR As Long = &H40
Public Const S_OK As Long = 0

Public Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" ( _
    ByVal pvInstance As LongPtr, _
    ByVal oVft As LongPtr, _
    ByVal cc As Long, _
    ByVal vtReturn As Integer, _
    ByVal cArgs As Long, _
    ByRef rgvt As Integer, _
    ByRef rgpvarg As LongPtr, _
    ByRef pvargResult As Variant) As Long

' 定数
Public Const CC_STDCALL As Long = 4

Public Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Public WV2Loader As New c0_WebView2Loader
Public WV2Controller As New c1_WebView2Controller
Public WV2 As c2_WebView2

Sub a()

    '事前にフォームを表示してウィンドウハンドルを取得する
    UserForm1.Show

    '独自に作っているUIA ラッパークラスを使ってフォームを取得
    Dim win As uia_e
    Set win = e.getRoot.ffDescendants(c.ClsName("ThunderDFrame"), 5)
        
    Dim fr As uia_e
    Set fr = win.ffDescendants(c.Type_(Group))

    TargetHwnd = fr.prHwnd
    Debug.Print TargetHwnd
    
    Call WV2Loader.CreateWebView2Environment

End Sub

'フォームにWebView2を生成する処理
Public Sub WebView2錬成()
    '独自に作っているUIA ラッパークラスを使ってフォームを取得
    Dim win As uia_e
    Set win = e.getRoot.ffDescendants(c.ClsName("ThunderDFrame"), 5)
        
    Dim fr As uia_e
    Set fr = win.ffDescendants(c.Type_(Group))

    TargetHwnd = fr.prHwnd
    Debug.Print TargetHwnd
    
    Call WV2Loader.CreateWebView2Environment
    'Call WV2Loader.DebugLoader
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
Public Sub RegisterNavigationCompleted_(ByVal pWebView As LongPtr)
    Dim vTable(3) As LongPtr
    Dim pVTable As LongPtr
    Static pNavHandler As LongPtr

    'vTable 構築 (QueryInterface, AddRef, Release は共通でOK)
    vTable(0) = GetAddr(AddressOf Handler_QueryInterface)
    vTable(1) = GetAddr(AddressOf Handler_AddRef)
    vTable(2) = GetAddr(AddressOf Handler_Release)
    vTable(3) = GetAddr(AddressOf NavCompleted_Invoke) ' 新しいInvoke

    If pVTable = 0 Then
        pVTable = GlobalAlloc(GPTR, 4 * LenB(vTable(0)))
        CopyMemory ByVal pVTable, vTable(0), 4 * LenB(vTable(0))
        pNavHandler = GlobalAlloc(GPTR, LenB(pVTable))
        CopyMemory ByVal pNavHandler, pVTable, LenB(pVTable)
    End If

    ' ICoreWebView2::add_NavigationCompleted (Index 15)
    ' 第1引数: [this], 第2引数: [eventHandler], 第3引数: [token(受け取り用アウト引数)]
    Dim token As LongLong
    Dim hr As Long
    Dim res As Variant

    Dim args(1) As Variant
    Dim argTypes(1) As Integer
    Dim argPtrs(1) As LongPtr

    args(0) = pNavHandler
    args(1) = VarPtr(token) ' 登録解除に使うトークンを受け取る
    argTypes(0) = 20: argTypes(1) = 20
    argPtrs(0) = VarPtr(args(0)): argPtrs(1) = VarPtr(args(1))

    'hr = DispCallFunc(pWebView, 15 * LenB(pWebView), CC_STDCALL, vbLong, 2, argTypes(0), argPtrs(0), res)
    'Debug.Print "Add_NavigationCompleted Result: " & hr
End Sub
