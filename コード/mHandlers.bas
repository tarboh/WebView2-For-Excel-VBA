Attribute VB_Name = "mHandlers"
'標準モジュール mHandlers

Option Explicit

' 文字列の長さを取得する（Unicode版）
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
' 指定したメモリ領域をVBAのStringとしてコピーする
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As LongPtr)


' AddressOf を LongPtr で受け取るための補助関数
Public Function GetAddr(ByVal addr As LongPtr) As LongPtr
    GetAddr = addr
End Function

'IUnknown:: QueryInterface
Public Function Handler_QueryInterface(ByVal This As LongPtr, ByVal riid As LongPtr, ByRef ppvObject As LongPtr) As Long
    ' 本来はGUIDを判定するが、今回は自分自身を返す
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

    pKeepEnv = pEnvironment ' ポインタをグローバル変数に退避

    If errorCode = 0 Then

        Call WV2Controller.CreateWebView2Controller(pEnvironment)

    End If

    Handler_Invoke = 0
End Function


'コントローラーの作成完了時にWebView2から呼び出されるコールバック関数
Public Function ControllerHandler_Invoke(ByVal This As LongPtr, ByVal errorCode As Long, ByVal pController As LongPtr) As Long
    
    Debug.Print "ControllerHandler_Invoke called. pController: " & pController
    
    If errorCode <> 0 Then Exit Function
    
    ' --- 最重要：WebView2が消えるのを阻止する ---
    CallAddRef pController
    
    'ポインタの登録
    WV2Controller.pController = pController
    
    '可視化
    WV2Controller.IsVisible = True

    'WebView2本体オブジェクトの取得
    Set WV2 = WV2Controller.GetWebView2
    
    ' NavigationCompleted イベントの登録
    WV2.WebView2_RegisterNavigationCompleted
    
    'Navigateメソッドの実行
    WV2.Navigate "Https://www.google.co.jp/"
    

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

    ControllerHandler_Invoke = 0
End Function

' ICoreWebView2NavigationCompletedEventHandler::Invoke
Public Function NavCompleted_Invoke(ByVal This As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    
    On Error Resume Next
    
    Debug.Print "--- Navigation Completed! ---"

    ' ここに読み込み完了後の処理を書く
    ' 例: 検索窓に文字を入れる ExecuteScript を呼ぶなど
    
    NavCompleted_Invoke = 0
End Function


' ExecuteScript完了時のコールバック
' Index 3: Invoke(HRESULT errorCode, LPCWSTR resultObjectAsJson)
Public Function ExecuteScript_Invoke(ByVal This As LongPtr, ByVal errorCode As Long, ByVal resultObjectAsJson As LongPtr) As Long
    If resultObjectAsJson <> 0 Then
        ' 実行結果はJSON形式の文字列ポインタで返ってくる
        Debug.Print "JS Execution Result: " & PtrToStrW(resultObjectAsJson)
    End If
    ExecuteScript_Invoke = 0
End Function

' ヘルパー：PtrToStrW (UnicodeポインタをVBA文字列へ)
Private Function PtrToStrW(ByVal pWStr As LongPtr) As String
    Dim length As Long
    Dim buf As String
    
    If pWStr = 0 Then
        PtrToStrW = ""
        Exit Function
    End If
    
    ' 1. 文字列の長さを取得 (Unicode文字数)
    length = lstrlenW(pWStr)
    
    If length > 0 Then
        ' 2. VBAの文字列バッファを確保 (1文字=2バイト)
        buf = Space$(length)
        
        ' 3. メモリからバッファへコピー
        ' VBAのStringは内部的にUnicodeなので、そのままコピー可能
        CopyMemory ByVal StrPtr(buf), ByVal pWStr, length * 2
        
        PtrToStrW = buf
    Else
        PtrToStrW = ""
    End If
End Function
