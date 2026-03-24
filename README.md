## 🚀 VBA WebView2 Integration Project
<img width="1916" height="1028" alt="image" src="https://github.com/user-attachments/assets/a2bcff2c-3443-42c7-9315-f2555433a61a" />


### 概要
本プロジェクトは、Excel VBAから外部ライブラリのインストールや参照設定を行うことなく、ユーザーフォーム上に **Microsoft Edge WebView2** を直接配置・制御することを目指す、次世代のVBA開発フレームワークです。

IE（Internet Explorer）のサポート終了に伴い、VBAにおけるブラウザ制御は大きな転換期を迎えました。本リポジトリでは、Officeに標準で内蔵されている `WebView2Loader.dll` をハックし、**DispCallFunc** を駆使した低レイヤーなCOM操作によって、VBAの限界を超えたモダンなUI/UXを提供します。



---

### 🔥 本プロジェクトのブレイクスルー
これまで、VBAからWebView2を利用するには .NET 経由のCOMラッパーを作成し、レジストリ登録（管理者権限）を行うという非現実的な壁がありました。

しかし、本プロジェクトでは以下の **「発見」** を基点に、VBA単体での動作を実現しています。

* **プリインストール資産の活用**: 
  `C:\Program Files\Microsoft Office\root\Office16\ADDINS\Microsoft Power Query for Excel Integrated\bin`
  内に存在する `WebView2Loader.dll` を `Declare` ステートメントで直接叩くことで、環境汚染なしにWebView2 Environmentを構築可能です。
* **完全なVTableハック**: 
  `WebView2Environment` 取得以降の全操作を `DispCallFunc` による関数ポインタ呼び出しで実行。参照設定ゼロでの動作を実現します。

---

### 🛠 セットアップ・フロー
WebView2オブジェクトを取得するまでの「泥臭くも精密な」プロセスは以下の通りです。

1. **Environment作成**: `WebView2Loader.dll` の `CreateCoreWebView2EnvironmentWithOptions` をコール。
2. **コールバック待機**: 標準モジュールで作成完了通知を受信。
3. **Controller生成**: `CreateCoreWebView2Controller` を実行し、ブラウザの描画領域を確保。
4. **WebView2取得**: `GetWebView2` メソッドより、本体である `ICoreWebView2` インスタンスを捕捉。
5. **ウィンドウ調整**: Windows APIを用いて、`Chrome_WidgetWin_0` クラスのウィンドウをユーザーフォームにフィッティング。

---

### 🌟 実現される未来
本プロジェクトの完成により、VBA開発者は以下の恩恵を受けることができます。

* **モダンGUIの統合**: ユーザーフォーム内に最新のHTML5/CSS3/JavaScriptによるUIを構築。
* **ActiveXからの脱却**: 動作が不安定なレガシーなTreeViewやListViewを、WebView2上の高性能なWebコンポーネントで代替。
* **フルコントロール**: イベントハンドラ（NavigationCompleted等）の自作実装による、Webスクレイピングや自動操作の完全掌握。

---

## ⚠️ Troubleshooting: Excel Crashes on Form Startup (Memory Address Table Issue)

### Problem
When you open the UserForm, Excel might crash immediately with an **"Automation error / Exception occurred"** or a silent termination. 

This happens because the VBA compiler sometimes fails to statically evaluate or initialize the memory address of standard module functions using `AddressOf` if they are not explicitly called or held elsewhere. If VBA doesn't map these addresses correctly, calling low-level Win32 APIs (like `DispCallFunc` or `CreateCoreWebView2EnvironmentWithOptions`) will lead to an access violation (Null/Invalid pointer crash).

### Solution
To force the VBA compiler to hold the `AddressOf` evaluate and keep the VTable memory address valid, add a dummy Sub at the bottom of your standard module (e.g., `Module1`):

```vba
' Dummy Sub to prevent VBA compiler optimization / address loss
Public Sub RegisterNavigationCompleted_()
    Static vTable As LongPtr
    vTable = GetAddr(AddressOf Handler_QueryInterface)
End Sub
```

By having this `AddressOf` evaluation written explicitly inside a Sub (even if it's never called!), the VBA runtime is forced to allocate the correct memory address table, bypassing the startup crash magically.

---


### 📚 謝辞
本プロジェクトの着想にあたり、重要な技術的知見を共有してくださった **もっくん** 様（[参考記事](https://eschamali.github.io/StarterWebScrapingKit/#userform-powershell)）に深く感謝いたします。
