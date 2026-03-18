WebView2自体のフレームワーク？アーキテクチャ？は以下のような構成でした。

WebView2Loader (WebView2.exeを起動するためのプログラム。VBAからDeclareステートメントで呼び出し可能）

　┗ WebView2Environment (WebView2.exe全体。WebView2Controller及びWebView2のオブジェクト生成などを担当）
 
　　　┗ WebView2Controller（外側ウィンドウであるChrome_WidgetWin_0の制御及びWebView2本体の取得関連）
   
　　　┗ WebView2  (ブラウザ本体であるChrome_WidgetWin_1の制御関連。NavigateとかExecuteScriptとか、各種イベントハンドラの登録とか）

Loader以外は全部IUnknownベースのインターフェースしか持っていない（＝IDispatchが無くてObjectとして扱えない）ので、
各機能をVBAから呼び出す時は全てDispCallFuncを経由する必要があります。
UIautomationで遊んで得たスキルがそのまんま転用できてて楽しいです。

WebView2EnvironmentとWebView2Controllerはオブジェクト生成が完了したことをコールバックで通知してくるので、
その通知を受け取るためのハンドラーを用意した上で、
IUnknownインターフェースの３メソッド（QueryInterface,AddRef,Release）とハンドラーの関数ポインタをまとめてVTbleを作成し、
それをWebView2側に渡す処理が必要になります。
WebView2本体の各種ハンドラーの登録も同じ手法になっています。

VBA側からDispCallFuncでWebViewに制御を投げる

　↓
 
処理を終えたWebView2がVBA側のハンドラを叩く

　↓
 
ハンドラ内で次の処理をWebView2に投げる

　↓
 
（以下繰り返し）

という処理のリレーによって非同期通信を実現する設計思想になっています。
