VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c2_WebView2Controller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Class c2_WebView2Controller

Option Explicit

Public pController As LongPtr

Public WithEvents WebView2 As c3_WebView2
Attribute WebView2.VB_VarHelpID = -1
Public Event NavigationCompleted()
Public Event ScriptResultReceived(result As String)

Public Event WebVeiw2ReadyCompleted()


'ここから下がWevView2Controllerの関数

'get_IsVisible(VtblIndex:3)
Public Property Get IsVisible() As Boolean

    Dim value As Long ' 戻り値（BOOL）を受け取る変数
    Dim hr As Long     'DispCallFuncのHRESULT
    Dim res As Variant 'put_IsVisibleのHRESULT
    
    ' 引数0: 値を書き込んでもらうための「変数valueのアドレス」を渡す
    Dim vArgs(0) As Variant: vArgs(0) = VarPtr(value)
    Dim vTypes(0) As Integer: vTypes(0) = vbLongPtr
    Dim vPtrs(0) As LongPtr: vPtrs(0) = VarPtr(vArgs(0))
    
    ' Index 3: get_IsVisible を実行
    hr = DispCallFunc(pController, 3 * LenB(pController), CC_STDCALL, vbLong, 1, vTypes(0), vPtrs(0), res)
    
    If hr = 0 Then
        If res = 0 Then
            If value <> 0 Then
                IsVisible = True
            Else
                IsVisible = False
            End If
        Else
            'なんかエラー吐かせる
            Debug.Print "WebView2オブジェクトのIsVisibleプロパティの取得に失敗しました。"
        End If
    Else
        '致命的失敗
        Debug.Print "WebView2オブジェクトのIsVisibleプロパティを取得しようとしましたが、それ以前にDispCallFuncが失敗しました。"
    End If

End Property

'put_IsVisible(VtblIndex:4)
Public Property Let IsVisible(value As Boolean)

    Dim hr As Long     'DispCallFuncのHRESULT
    Dim res As Variant 'put_IsVisibleのHRESULT

    Dim vArgs(0) As Variant
    vArgs(0) = CLng(IIf(value, 1, 0))
    Dim vTypes(0) As Integer: vTypes(0) = vbLong
    Dim vPtrs(0) As LongPtr: vPtrs(0) = VarPtr(vArgs(0))
    hr = DispCallFunc(pController, 4 * LenB(pController), CC_STDCALL, vbLong, 1, vTypes(0), vPtrs(0), res)
    
    If hr = 0 Then
        If res = 0 Then
        Else
            Debug.Print "put_IsVisible Result: " & res
        End If
    Else
        '致命的失敗
        Debug.Print "WebView2オブジェクトのIsVisibleプロパティを取得しようとしましたが、それ以前にDispCallFuncが失敗しました。"
    End If

End Property

'put_IsVisible の本質的な役割
'このメソッドは、WebView2の「親ウィンドウ（UserForm等）に対する表示状態のスイッチ」です。
'
'1. 描画プロセスの起動スイッチ
'WebView2は、CreateCoreWebView2Controller で生成された直後は、メモリ上に存在していても「描画プロセス（Renderer）」がフル稼働していません。
'put_IsVisible = True (1) を叩いた瞬間に、WebView2は親ウィンドウのハンドル（HWND）に対して自分の描画内容を転送し始めます。
'
'2. 親ウィンドウとの「親子関係」の確定
'これを True にすることで、Win32 APIレベルでの「子ウィンドウ」としての振る舞いが本格化します。
'
'True (1): 親ウィンドウが再描画されるたびに、WebView2も自分の領域を再描画します。
'False (0): WebView2は描画を停止し、リソース消費を抑えます。ただし、バックグラウンドでJavaScriptなどは動き続けます。
'実装上の重要な挙動
'① put_Bounds との順序
'通常、put_Bounds（サイズ設定）を先に呼び出してから put_IsVisible(True) を呼び出します。
'もしサイズが 0x0 の状態で Visible = True にしても何も見えません。逆に、サイズを設定した後に Visible = True にしないと、領域が真っ白（または親フォームの背景色のまま）になります。
'
'② フォーカスとの連動
'put_IsVisible(True) にすると、WebView2はキーボードフォーカスを受け取れる状態になります。逆に False にすると、マウスイベントやキーボードイベントを一切受け付けなくなります。
'
'③ ウィンドウメッセージのハンドリング
'WebView2は内部で WM_SHOWWINDOW などのメッセージをこのメソッド経由で処理しています。VBAで UserForm.Hide などをした際、WebView2側が「自分も隠れた」と正しく認識させるためには、本来はこのメソッドを False にするのが正攻法です。


' get_CoreWebView2 (Index 25) 専用：ポインタのポインタを渡して実体を取得
Public Function GetWebView2() As c3_WebView2

    Dim ppWebView2 As LongPtr
    Dim args(0) As Variant
    Dim argTypes(0) As Integer
    Dim argPtrs(0) As LongPtr
    
    Dim hr As Long
    Dim result As Variant

    ' 受け取り用変数のアドレスを渡す
    args(0) = VarPtr(ppWebView2)
    argTypes(0) = vbLongPtr ' 64bitならvbLongLong(20), 32bitならvbLong(3)
    argPtrs(0) = VarPtr(args(0))
 
    hr = DispCallFunc(pController, 25 * LenB(pController), CC_STDCALL, vbLong, 1, argTypes(0), argPtrs(0), result)
    If hr = 0 Then
    
        If result = 0 Then
        
            Dim WV2 As New c3_WebView2
            WV2.ppWebView2 = ppWebView2
            'Set GetWebView2 = WV2
            Set WebView2 = WV2
        
        Else
        
            Debug.Print "get_CoreWebView2が失敗しました。result:", result
        
        End If
    
    Else
    
        Debug.Print "get_CoreWebView2を実行しようとしましたが、そもそもDispCallFuncが失敗しました。HR:", hr
    
    End If

End Function


Public Sub ReadyCompleted()
    RaiseEvent WebVeiw2ReadyCompleted
End Sub

Public Sub CloseWebView2()
    Dim hr As Long
    Dim result As Variant

    ' pController は Environment::CreateCoreWebView2Controller
    ' の完了ハンドラ (Invoke) で受け取ったポインタです。
    If pController = 0 Then Exit Sub

    ' ICoreWebView2Controller::Close (Index 24) を実行
    ' これにより、関連する msedgewebview2.exe プロセスが終了します。
    hr = DispCallFunc(pController, 24 * LenB(pController), CC_STDCALL, vbLong, 0, 0, 0, result)

    If hr = 0 Then
        Debug.Print "WebView2 Controller Closed."
    Else
        Debug.Print "Close Failed. HR: " & hr
    End If

    ' ポインタを破棄して、VBA側からも参照を完全に切る
    ' これを忘れると、VBA側がポインタを「握ったまま」になり、不安定になります。
    Set WebView2 = Nothing 'オブジェクト自体の破棄によりポインタも消える・・・はず
    pController = 0
    'pKeepEnv = 0
End Sub

Private Sub WebView2_NavigationCompleted()
    RaiseEvent NavigationCompleted
End Sub

Private Sub WebView2_ReceiveScriptResult(ByVal result As String)
    RaiseEvent ScriptResultReceived(result)
End Sub

