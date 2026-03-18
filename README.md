やばいです。ブレイクスルーです。

IEサポート終了の話が出で以降、Excel VBA開発者たちの中では「モダンブラウザをVBAから以下に制御するか？」は大きな課題となっています。

昨今のスタンダードはSelenium VBAによる外部ブラウザの制御かと思いますが、この度、ユーザーフォーム上にWebView2コントロールを配置するというアプローチにおいて、大きな進展がありました。

これまではWebView2を使うと言っても.NET系アプリに実装したWebView2コントロールを、そのアプリのCOMインターフェースを用意してタイプライブラリを作成し、それをインストールした上で参照設定を通すというような非現実的な手法しかありませんでした。

そんな中、https://eschamali.github.io/StarterWebScrapingKit/#userform-powershell<br>
こちらの記事で紹介されていますが、もっくんさんが重要な事実を発見しました。

その中で今回のプロジェクトの直接的なきっかけとなったのが、<br>
C:\Program Files\Microsoft Office\root\Office16\ADDINS\Microsoft Power Query for Excel Integrated\bin<br>
こちらのパス内に、WebView2のプログラム群(dll)が最初からインストール済みであるという事実です。

少し使い方を調べてみると、WebView2Loader.dllの関数はDecleareステートメントによってVBAから直接呼び出せることが分かりました。

後はひたすら泥臭いコーディング＆テストを重ねた末、
以下のような手順で、フォーム上に配置したWebVeiw2を制御できるようになりました。

＜WebView2オブジェクト取得まで＞<br>
Declareステートメント宣言でWebView2Loader.dllのCreateCoreWebView2EnvironmentWithOptionsをコール<br>
↓<br>
Microsoft.Web.WebView2.Core.dllがロードされ、<br>
WebView2Environmentオブジェクトが作成される<br>
↓<br>
標準モジュールに準備したWebView2Environmentの<br>
作成完了通知を受け取る関数がコールされる<br>
↓<br>
VBAからWebView2Environment.CreateCoreWebView2Controllerメソッドを呼ぶ<br>
↓<br>
WebView2.exe内でWebView2Controllerが作成される<br>
↓<br>
標準モジュールに準備したWebView2Controllerの<br>
作成完了通知を受け取る関数がコールされる<br>
↓<br>
VBAからWebView2Controller.GetWebView2メソッドでWebView2オブジェクトを取得<br>
↓<br>
WindowsAPI等でChrome_WidgetWin_0クラスウィンドウの位置とサイズを調整<br>
<br>
＜WebView2の制御：イベントハンドラの登録例＞<br>
WebView2.add_NavigationCompletedを実行<br>
（実行時にAddress of演算子でコールバック関数のポインタを渡す）<br>
↓<br>
WebView2.Navigateメソッドを実行<br>
↓<br>
ページ読み込みが完了したら、<br>
標準モジュールに配置したイベントハンドラがコールされる<br>
<br>
なお、WebView2Environmentオブジェクト取得以降の関数は、<br>
全部DispCallFuncを使用してコールする必要があるため、<br>
VBA超上級者でないとライブラリ開発は難しいでしょう。<br>
<br>
このリポジトリで開発を進めていきますので、<br>
興味のある方は各種イベントハンドラの実装に挑戦してみてください。<br>
<br>
また、このプロジェクトの果てに得られる機能として<br>
・モダンなWebブラウザコントロールをユーザーフォーム配下に実装し、完全なコントロール権を得ること<br>
・ユーザーフォーム上でのマウスホイールにるスクロール操作や右クリックによるコンテキストメニューが使えるようになること<br>
・ActiveXコントロールに依存しないTreeViewやListViewコントロールを使えるようになること<br>
などを期待しています。<br>
もしかすると、ユーザーフォーム上のコントロールは全部WebView2上に描画するのが常識という時代が来るかもしれません。
