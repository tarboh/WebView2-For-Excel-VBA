VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c3_WebView2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Class c3_WebView2

Option Explicit

Public ppWebView2 As LongPtr
Public Event NavigationStarting()
Public Event NavigationCompleted()
Public Event ReceiveScriptResult(ByVal result As String)
Public LastScriptResult As String


Private Col_Handler As New Collection
Private m_IsNavigating As Boolean
Private m_IsScriptRunning As Boolean

' --- イベントハンドラ登録 ---
Public Sub WebView2_RegisterNavigationCompleted()
    Dim NavHandler As c4_Handler: Set NavHandler = New c4_Handler
    NavHandler.CreateVTble GetAddr(AddressOf NavCompleted_Invoke), ppWebView2
    Col_Handler.Add NavHandler

    Dim Token As LongPtr, hr As Long, res As Variant
    Dim args(1) As Variant, argTypes(1) As Integer, argPtrs(1) As LongPtr

    args(0) = NavHandler.Pointer
    args(1) = VarPtr(Token)
    argTypes(0) = vbLongPtr: argTypes(1) = vbLongPtr
    argPtrs(0) = VarPtr(args(0)): argPtrs(1) = VarPtr(args(1))

    hr = DispCallFunc(ppWebView2, 15 * LenB(ppWebView2), CC_STDCALL, vbLong, 2, argTypes(0), argPtrs(0), res)
    
    If hr = 0 And res = 0 Then
        NavHandler.Token = Token
        RegisterInstance NavHandler.Pointer, Me
        Debug.Print "NavigationHandler Registered."
    End If
End Sub

Public Function NavigateToString(htmlContent As String) As Long
    '    virtual HRESULT STDMETHODCALLTYPE NavigateToString(
    '        /* [in] */ LPCWSTR htmlContent) = 0;
    Dim nArgs(0) As Variant: nArgs(0) = StrPtr(htmlContent)
    Dim nTypes(0) As Integer: nTypes(0) = vbLongPtr
    Dim nPtrs(0) As LongPtr: nPtrs(0) = VarPtr(nArgs(0))
    Dim hr As Long
    Dim res As Variant
    
    hr = DispCallFunc(ppWebView2, 6 * LenB(ppWebView2), CC_STDCALL, vbLong, 1, nTypes(0), nPtrs(0), res)
    
    If hr = 0 Then
        Debug.Print "DispCallFunc成功"
        If res = 0 Then
            Debug.Print "NavigateToString成功"
            NavigateToString = 0
        Else
            Debug.Print "NavigateToString失敗", res
        End If
    Else
        Debug.Print "DispCallFunc失敗", hr
    End If

End Function

Public Function add_NavigationStarting()
    Dim NavHandler As c4_Handler: Set NavHandler = New c4_Handler
    NavHandler.CreateVTble GetAddr(AddressOf NavigationStarting_Invoke), ppWebView2
    Col_Handler.Add NavHandler

    Dim Token As LongPtr, hr As Long, res As Variant
    Dim args(1) As Variant, argTypes(1) As Integer, argPtrs(1) As LongPtr

    args(0) = NavHandler.Pointer
    args(1) = VarPtr(Token)
    argTypes(0) = vbLongPtr: argTypes(1) = vbLongPtr
    argPtrs(0) = VarPtr(args(0)): argPtrs(1) = VarPtr(args(1))

    hr = DispCallFunc(ppWebView2, 7 * LenB(ppWebView2), CC_STDCALL, vbLong, 2, argTypes(0), argPtrs(0), res)
    
    If hr = 0 And res = 0 Then
        NavHandler.Token = Token
        RegisterInstance NavHandler.Pointer, Me
        Debug.Print "NavigationHandler Registered."
    End If
End Function

' --- 現在のURLを取得する関数 ---
' ICoreWebView2::get_Source (Index 4)
Public Property Get Source() As String

    Dim pUri As LongPtr, hr As Long, res As Variant
    Dim args(0) As Variant, argTypes(0) As Integer, argPtrs(0) As LongPtr

    args(0) = VarPtr(pUri)
    argTypes(0) = vbLongPtr
    argPtrs(0) = VarPtr(args(0))

    hr = DispCallFunc(ppWebView2, 4 * LenB(ppWebView2), CC_STDCALL, vbLong, 1, argTypes(0), argPtrs(0), res)
    
    If hr = 0 Then
        
        If res = 0 Then
        
            Debug.Print "get_source成功！！！"
            If pUri <> 0 Then
                Debug.Print "アドレス取得成功！！！pUri:" & pUri
                Dim url As String
                url = PtrToStrW(pUri)
                Source = url
                Debug.Print "URL:" & url
                Call CoTaskMemFree(pUri)
            Else
                Debug.Print "アドレス取得失敗、pUri:" & pUri
            End If
            
        Else
            Debug.Print "get_source失敗、res:" & res
        
        End If
    Else
        Debug.Print "dispcallfunc失敗、hr:" & hr
    End If

End Property



' --- Navigate 関連 ---
' 非同期版：メソッドを呼んだらすぐに制御が戻る
Public Function NavigateAsync(ByVal url As String) As Long
    Dim nArgs(0) As Variant: nArgs(0) = StrPtr(url)
    Dim nTypes(0) As Integer: nTypes(0) = vbLongPtr
    Dim nPtrs(0) As LongPtr: nPtrs(0) = VarPtr(nArgs(0))
    Dim res As Variant
    
    NavigateAsync = DispCallFunc(ppWebView2, 5 * LenB(ppWebView2), CC_STDCALL, vbLong, 1, nTypes(0), nPtrs(0), res)
End Function

' 同期版：読み込み完了まで待機する
Public Sub NavigateSync(ByVal url As String, Optional ByVal TimeoutSec As Double = 10)
    m_IsNavigating = True
    If NavigateAsync(url) = 0 Then
        WaitNavigation TimeoutSec
    End If
End Sub

' --- ExecuteScript 関連 ---
' 非同期版：実行を依頼するだけで戻る。結果はイベントで受け取る
Public Sub ExecuteScriptAsync(ByVal script As String)
    Dim ScriptHandler As c4_Handler: Set ScriptHandler = New c4_Handler
    ScriptHandler.CreateVTble GetAddr(AddressOf ExecuteScript_Invoke), ppWebView2
    Col_Handler.Add ScriptHandler
    RegisterInstance ScriptHandler.Pointer, Me

    Dim args(1) As Variant, argTypes(1) As Integer, argPtrs(1) As LongPtr
    Dim res As Variant
    
    args(0) = StrPtr(script)
    args(1) = ScriptHandler.Pointer
    argTypes(0) = vbLongPtr: argTypes(1) = vbLongPtr
    argPtrs(0) = VarPtr(args(0)): argPtrs(1) = VarPtr(args(1))

    Call DispCallFunc(ppWebView2, 29 * LenB(ppWebView2), CC_STDCALL, vbLong, 2, argTypes(0), argPtrs(0), res)
End Sub

' 同期版：実行結果を戻り値として直接受け取る
Public Function ExecuteScriptSync(ByVal script As String, Optional ByVal TimeoutSec As Double = 5) As String
    m_IsScriptRunning = True
    ExecuteScriptAsync script
    
    ' 専用の待機ループ
    Dim start As Double: start = Timer
    Do While m_IsScriptRunning
        Debug.Print "waiting ExecuteScriptSync"
        DoEvents
        If Timer - start > TimeoutSec Then
            Debug.Print "ExecuteScriptSync: Timeout!"
            Exit Do
        End If
    Loop
    
    ExecuteScriptSync = LastScriptResult
End Function

' --- 内部通知用 ---
Public Sub NotifyNavigationStarting()
    RaiseEvent NavigationStarting
End Sub
Public Sub NotifyNavigationCompleted()
    m_IsNavigating = False
    RaiseEvent NavigationCompleted
End Sub

Public Sub NotifyExecuteScriptCompleted(ByVal jsonResult As String)
    LastScriptResult = jsonResult
    m_IsScriptRunning = False
    RaiseEvent ReceiveScriptResult(jsonResult)
End Sub

' --- 後始末 ---
Public Sub Finalize()
    Dim h As c4_Handler
    On Error Resume Next
    For Each h In Col_Handler
        UnregisterInstance h.Pointer
    Next h
    Set Col_Handler = New Collection
End Sub

Private Sub Class_Terminate()
    Finalize
    Debug.Print "c3_WebView2 Terminated."
End Sub

' --- 待機用ヘルパー ---
Private Sub WaitNavigation(ByVal TimeoutSec As Double)
    Dim start As Double: start = Timer
    Do While m_IsNavigating
        DoEvents
        If Timer - start > TimeoutSec Then Exit Do
    Loop
End Sub
