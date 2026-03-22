Attribute VB_Name = "mHandlers"
'標準モジュール mHandlers

Option Explicit

' AddressOf を LongPtr で受け取るための補助関数
Public Function GetAddr(ByVal addr As LongPtr) As LongPtr
    GetAddr = addr
End Function

'IUnknown:: QueryInterface
Public Function Handler_QueryInterface(ByVal This As LongPtr, ByVal riid As LongPtr, ByRef ppvObject As LongPtr) As Long
    ' 本来はGUIDを判定するが、今回は自分自身を返す
    Debug.Print "クエリ！"
    ppvObject = This
    Handler_QueryInterface = S_OK
End Function

' IUnknown::AddRef / Release (簡易的に1を返す)
Public Function Handler_AddRef(ByVal This As LongPtr) As Long: Handler_AddRef = 1: End Function
Public Function Handler_Release(ByVal This As LongPtr) As Long: Handler_Release = 1: End Function

' ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler::Invoke
' ここにWebView2から初期化結果が届く
Public Function Handler_Invoke(ByVal This As LongPtr, ByVal errorCode As Long, ByVal pEnvironment As LongPtr) As Long
    Debug.Print "WebView2 Environment Created. ErrorCode: " & errorCode

    If errorCode = 0 Then

        Call UserForm1.WV2Environment.CreateWebView2Controller(pEnvironment)

    End If

    Handler_Invoke = 0
End Function


'コントローラーの作成完了時にWebView2から呼び出されるコールバック関数
Public Function ControllerHandler_Invoke(ByVal This As LongPtr, ByVal errorCode As Long, ByVal pController As LongPtr) As Long
    
    Debug.Print "ControllerHandler_Invoke called. pController: " & pController
    
    If errorCode <> 0 Then Exit Function
    
    ' --- 最重要：WebView2が消えるのを阻止する ---
    CallAddRef pController
    
    Set UserForm1.WV2Controller = New c2_WebView2Controller
    
    'ポインタの登録
    UserForm1.WV2Controller.pController = pController
    
    '可視化
    UserForm1.WV2Controller.IsVisible = True

    'WebView2本体オブジェクトの取得
    'Set WV2 = WV2Controller.GetWebView2
    Call UserForm1.WV2Controller.GetWebView2
    Set UserForm1.WV2 = UserForm1.WV2Controller.WebView2
    
    ' Get Settings
    Call UserForm1.WV2Controller.WebView2.get_Settings
    
    ' Set ScriptDialogsEnabled Property
    UserForm1.WV2Controller.WebView2.Settings.AreDefaultScriptDialogsEnabled = True
    
    ' NavigationCompleted イベントの登録
    Call UserForm1.WV2Controller.WebView2.add_NavigationStarting
    Call UserForm1.WV2Controller.WebView2.add_ContentLoading
    Call UserForm1.WV2Controller.WebView2.add_SourceChanged
    Call UserForm1.WV2Controller.WebView2.add_HistoryChanged
    Call UserForm1.WV2Controller.WebView2.add_NavigationCompleted
    Call UserForm1.WV2Controller.WebView2.add_FrameNavigationStarting
    Call UserForm1.WV2Controller.WebView2.add_FrameNavigationCompleted
    Debug.Print "add_ScriptDialogOpening:", UserForm1.WV2Controller.WebView2.add_ScriptDialogOpening
    
    'Handler2を使う方式のイベント登録
    Call UserForm1.WV2Controller.WebView2.AddNavigationCompletedHandler(UserForm1.NavigationCompletedHandler)
    
    Debug.Print "ppWebView2:", UserForm1.WV2Controller.WebView2.ppWebView2
    
    'Navigateメソッドの実行
    'WV2.Navigate "Https://www.google.co.jp/"
    'Call WV2Controller.WebView2.NavigateSync("Https://www.google.co.jp/")

    ' 4. Win32 API で「力技」の可視化
    DoEvents
    Dim childHwnd As LongPtr
    ' 前回の調査で判明した「Chrome_WidgetWin_0」を直接操作
    childHwnd = FindWindowEx(TargetHwnd, 0, "Chrome_WidgetWin_0", vbNullString)

    If childHwnd <> 0 Then
        ' WebView2内部の put_Bounds が失敗していても、
        ' ウィンドウハンドルさえあれば OS レベルでサイズをねじ込めます
        MoveWindow childHwnd, 0, 0, 800, 600, 1
        Debug.Print "Final Sync via Win32 API. childHwnd: " & childHwnd & "(" & Hex(childHwnd) & ")"
    End If

    Call サイズ調整
    
    Call UserForm1.WV2Controller.ReadyCompleted

    ControllerHandler_Invoke = 0
End Function

Public Function NavigationStarting_Invoke(ByVal This As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    
    Dim target As c3_WebView2
    Set target = GetInstance(This)
    
    If Not target Is Nothing Then
        ' クラス側のメソッドを叩く
        target.NotifyNavigationStarting
    Else
        ' 【重要】もしターゲットが見つからない（クラスが破棄された後）なら
        ' WebView2側に残っている「幽霊ハンドラ」の可能性があるので
        ' 辞書からこのポインタを掃除しておく（念のため）
        UnregisterInstance This
    End If
    
    NavigationStarting_Invoke = 0

End Function

Public Function ContentLoading_Invoke(ByVal This As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(This)
    
    If Not target Is Nothing Then
        ' クラス側のメソッドを叩く
        target.NotifyContentLoading
    Else
        ' 【重要】もしターゲットが見つからない（クラスが破棄された後）なら
        ' WebView2側に残っている「幽霊ハンドラ」の可能性があるので
        ' 辞書からこのポインタを掃除しておく（念のため）
        UnregisterInstance This
    End If
    
    ContentLoading_Invoke = 0
    
End Function

Public Function SourceChanged_Invoke(ByVal This As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(This)
    
    If Not target Is Nothing Then
        ' クラス側のメソッドを叩く
        target.NotifySourceChanged
    Else
        ' 【重要】もしターゲットが見つからない（クラスが破棄された後）なら
        ' WebView2側に残っている「幽霊ハンドラ」の可能性があるので
        ' 辞書からこのポインタを掃除しておく（念のため）
        UnregisterInstance This
    End If
    
    SourceChanged_Invoke = 0
End Function

Public Function HistoryChanged_Invoke(ByVal This As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(This)
    
    If Not target Is Nothing Then
        ' クラス側のメソッドを叩く
        target.NotifyHistoryChanged
    Else
        ' 【重要】もしターゲットが見つからない（クラスが破棄された後）なら
        ' WebView2側に残っている「幽霊ハンドラ」の可能性があるので
        ' 辞書からこのポインタを掃除しておく（念のため）
        UnregisterInstance This
    End If
    
    HistoryChanged_Invoke = 0
End Function
Public Function NavCompleted_Invoke(ByVal This As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    
    Dim target As c3_WebView2
    Set target = GetInstance(This)
    
    If Not target Is Nothing Then
        ' クラス側のメソッドを叩く
        target.NotifyNavigationCompleted
    Else
        ' 【重要】もしターゲットが見つからない（クラスが破棄された後）なら
        ' WebView2側に残っている「幽霊ハンドラ」の可能性があるので
        ' 辞書からこのポインタを掃除しておく（念のため）
        UnregisterInstance This
    End If
    
    NavCompleted_Invoke = 0
End Function

Public Function FrameNavigationStarting_Invoke(ByVal This As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    
    Dim target As c3_WebView2
    Set target = GetInstance(This)
    
    If Not target Is Nothing Then
        ' クラス側のメソッドを叩く
        target.NotifyFrameNavigationStarting
    Else
        ' 【重要】もしターゲットが見つからない（クラスが破棄された後）なら
        ' WebView2側に残っている「幽霊ハンドラ」の可能性があるので
        ' 辞書からこのポインタを掃除しておく（念のため）
        UnregisterInstance This
    End If
    
    FrameNavigationStarting_Invoke = 0
End Function
'FrameNavigationCompleted_Invoke
Public Function FrameNavigationCompleted_Invoke(ByVal This As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    
    Dim target As c3_WebView2
    Set target = GetInstance(This)
    
    If Not target Is Nothing Then
        ' クラス側のメソッドを叩く
        target.NotifyFrameNavigationCompleted
    Else
        ' 【重要】もしターゲットが見つからない（クラスが破棄された後）なら
        ' WebView2側に残っている「幽霊ハンドラ」の可能性があるので
        ' 辞書からこのポインタを掃除しておく（念のため）
        UnregisterInstance This
    End If
    
    FrameNavigationCompleted_Invoke = 0
End Function
'ScriptDialogOpening
Public Function ScriptDialogOpening_Invoke(ByVal This As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    
    Dim target As c3_WebView2
    Set target = GetInstance(This)
    
    If Not target Is Nothing Then
        ' クラス側のメソッドを叩く
        target.NotifyScriptDialogOpening
    Else
        ' 【重要】もしターゲットが見つからない（クラスが破棄された後）なら
        ' WebView2側に残っている「幽霊ハンドラ」の可能性があるので
        ' 辞書からこのポインタを掃除しておく（念のため）
        UnregisterInstance This
    End If
    
    ScriptDialogOpening_Invoke = 0
End Function

' ExecuteScript完了時のコールバック
' Index 3: Invoke(HRESULT errorCode, LPCWSTR resultObjectAsJson)
Public Function ExecuteScript_Invoke(ByVal This As LongPtr, ByVal errorCode As Long, ByVal resultJsonPtr As LongPtr) As Long
    Dim target As c3_WebView2
    Set target = GetInstance(This)
    
    If Not target Is Nothing Then
        ' ポインタから文字列を抽出し、ターゲットに渡す
        target.NotifyExecuteScriptCompleted PtrToStrW(resultJsonPtr)
        
        ' 実行が終わったハンドラは名簿から抹消（使い捨て）
        UnregisterInstance This
    End If
    ExecuteScript_Invoke = 0
End Function

' ヘルパー：PtrToStrW (UnicodeポインタをVBA文字列へ)
Public Function PtrToStrW(ByVal pWStr As LongPtr) As String
    Dim Length As Long
    Dim buf As String

    If pWStr = 0 Then
        PtrToStrW = ""
        Exit Function
    End If

    ' 1. 文字列の長さを取得 (Unicode文字数)
    Length = lstrlenW(pWStr)

    If Length > 0 Then
        ' 2. VBAの文字列バッファを確保 (1文字=2バイト)
        buf = Space$(Length)

        ' 3. メモリからバッファへコピー
        ' VBAのStringは内部的にUnicodeなので、そのままコピー可能
        CopyMemory ByVal StrPtr(buf), ByVal pWStr, Length * 2

        PtrToStrW = buf
    Else
        PtrToStrW = ""
    End If
End Function
'Public Function PtrToStrW(ByVal pWStr As LongPtr) As String
'    Dim pBSTR As LongPtr
'    If pWStr = 0 Then Exit Function
'
'    ' UnicodeポインタからBSTRを生成
'    pBSTR = SysAllocString(pWStr)
'    If pBSTR <> 0 Then
'        ' String変数の内部ポインタに直接コピー
'        CopyMemory ByVal VarPtr(PtrToStrW), pBSTR, LenB(ppWebView2)
'        ' ※SysFreeStringはVBAの自動解放に任せる
'    End If
'End Function