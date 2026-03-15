VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c2_WebView2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Class c2_WebView2

Option Explicit

Public ppWebView2 As LongPtr

#If Win64 Then
    Private Const vbLongPtr = vbLongLong
#Else
    Private Const vbLongPtr = vbLong
#End If

Private vTable(3) As LongPtr
Private pVTable As LongPtr
Private pNavHandler As LongPtr

Private pExeScriptHandler As LongPtr
Private vExeScriptTable(3) As LongPtr
Private pExeScriptVTable As LongPtr

'イベントハンドラの登録
Public Sub WebView2_RegisterNavigationCompleted() 'ByVal pWebView As LongPtr)

    ' vTable 構築 (QueryInterface, AddRef, Release は共通でOK)
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
    Dim token As LongPtr
    Dim hr As Long
    Dim res As Variant

    Dim args(1) As Variant
    Dim argTypes(1) As Integer
    Dim argPtrs(1) As LongPtr

    args(0) = pNavHandler
    args(1) = VarPtr(token) ' 登録解除に使うトークンを受け取る
    argTypes(0) = 20: argTypes(1) = 20
    argPtrs(0) = VarPtr(args(0)): argPtrs(1) = VarPtr(args(1))

    hr = DispCallFunc(ppWebView2, 15 * LenB(ppWebView2), CC_STDCALL, vbLong, 2, argTypes(0), argPtrs(0), res)
    If hr = 0 Then
        If res = 0 Then
            Debug.Print "Add_NavigationCompleted成功：" & res
        Else
            Debug.Print "Add_NavigationCompleted失敗：" & res
        End If
    Else
        Debug.Print "Add_NavigationCompletedのDsipCallfunc失敗：" & hr
    End If
End Sub

' 指定したURLへ移動する
Public Function Navigate(ByVal url As String) As Long
    Dim hr As Long
    Dim res As Variant
    
    If ppWebView2 = 0 Then
        Debug.Print "WebView2_Navigate: CoreWebView2 pointer is null."
        Navigate = -1
        Exit Function
    End If
    
    ' 引数0: URLの文字列ポインタ (StrPtr)
    Dim nArgs(0) As Variant: nArgs(0) = StrPtr(url)
    Dim nTypes(0) As Integer: nTypes(0) = vbLongPtr
    Dim nPtrs(0) As LongPtr: nPtrs(0) = VarPtr(nArgs(0))
    
    ' ICoreWebView2::Navigate は vTable Index 5
    hr = DispCallFunc(ppWebView2, 5 * LenB(ppWebView2), CC_STDCALL, vbLong, 1, nTypes(0), nPtrs(0), res)
    
    Debug.Print "Navigate to [" & url & "] Result: " & hr
    Navigate = hr
    
End Function


' JavaScriptを実行する
Public Sub ExecuteScript(ByVal script As String)

    
    ' 初回呼び出し時にハンドラーを作成（例の「消すと落ちる」対策と同じ構造）
    If pExeScriptHandler = 0 Then

        vExeScriptTable(0) = GetAddr(AddressOf Handler_QueryInterface)
        vExeScriptTable(1) = GetAddr(AddressOf Handler_AddRef)
        vExeScriptTable(2) = GetAddr(AddressOf Handler_Release)
        vExeScriptTable(3) = GetAddr(AddressOf ExecuteScript_Invoke)
        
        pExeScriptVTable = GlobalAlloc(GPTR, 4 * LenB(pKeepEnv))
        CopyMemory ByVal pExeScriptVTable, vExeScriptTable(0), 4 * LenB(pKeepEnv)
        pExeScriptHandler = GlobalAlloc(GPTR, LenB(pKeepEnv))
        CopyMemory ByVal pExeScriptHandler, pExeScriptVTable, LenB(pKeepEnv)
    End If

    Dim hr As Long
    Dim res As Variant
    Dim args(1) As Variant
    Dim argTypes(1) As Integer
    Dim argPtrs(1) As LongPtr

    ' ICoreWebView2::ExecuteScript (Index 29)
    ' 第1引数: [in] LPCWSTR script
    ' 第2引数: [in] ICoreWebView2ExecuteScriptCompletedHandler* handler
    args(0) = StrPtr(script)
    args(1) = pExeScriptHandler
    
    argTypes(0) = vbLongLong ' 64bit
    argTypes(1) = vbLongLong
    argPtrs(0) = VarPtr(args(0))
    argPtrs(1) = VarPtr(args(1))

    hr = DispCallFunc(ppWebView2, 29 * LenB(ppWebView2), CC_STDCALL, vbLong, 2, argTypes(0), argPtrs(0), res)
    
    If hr <> 0 Then Debug.Print "ExecuteScript DispCallFunc Error: " & hr
End Sub
