Attribute VB_Name = "Wv2Thunks"
''''''''''''''''''''''''''''''''''
' --- Wv2Thunks.bas 第9.10a段階 (旧称: Module9_9_9.bas) ---
'
'   第9.10a の変更点 (本セッション、Sentinel MEM_FREE バグ対処 = 仕様事実 10):
'     - 定数追加: MEM_FREE (&H10000&) / PAGE_READWRITE (&H4&)
'     - Sentinel_RecoverIfNeeded を「State ガード方式」に書き換え:
'         VirtualQuery の結果 (State=MEM_COMMIT かつ AllocationBase=prevBase かつ
'         Protect が読み書き可) で「領域が健在」と確認できた場合に限り、
'         ReadLongPtr / MemLongPtr / VirtualFree を実行する。
'         健在でなければ Sentinel_ClearPrevRegion だけ撃って安全に Exit Sub。
'       → 旧実装は解放済みアドレスへ ReadLongPtr / MemLongPtr を無条件に撃ち、
'         SAFEARRAY 経由のネイティブ配列アクセスで AV (= SEH 例外) を起こして
'         いた。この AV は On Error Resume Next では捕捉できず Excel が即落ち
'         していた (仕様事実 10)。設計原則 55 として明文化。
'     - 新規 Private Function: IsProtectReadWritable(prot) As Boolean
'         (PAGE_GUARD ビット検出 + PAGE_READWRITE / PAGE_EXECUTE_READWRITE 判定)
'     - 本体ロジック (Handler_xxx / dcf / Thunks_Init 等) はその他無変更
'
'   第9.9c の変更点 (継続、補完群):
'     - Test_ExecuteScript_Parallel を新規追加 (項目 D)
'         5 個の ExecuteScript を連続発行し、callback ID と結果値の対応で
'         同時並行性 (案 a、9.8c) が正しく機能することを実機検証する。
'         連番計算 ((1+1), (2+1), ... ) を投げて callback ID と結果の対応を見る。
'     - 本体ロジック (Handler_xxx / dcf / Thunks_Init 等) は無変更
'         (9.9c は Wv2Thunks にはテスト関数 1 個追加するだけ)
'
'   第9.9b の変更点 (継続、SPA 対応):
'     - HandlerKind に 2 種類追加 (HK_HistoryChanged=9 / HK_DOMContentLoaded=10)
'     - m_iidTable の配列上限を HK_DOMContentLoaded に拡張
'     - InitIIDTable に 2 IID 追加
'     - GetLongLongProperty を新設 (#If Win64 で囲む = 64bit 専用)
'     - BuildSpaTestHtml / Test_SpaTest_Real を新設 (SPA テスト用)
'     - 案 R': BuildSpaTestHtml の doPush/doReplace に try-catch を追加
'              (about:blank origin の SecurityError を画面に表示)
'
'   ★ 第9.9 (a + b + c) の 32bit 将来拡張への配慮 (方針 X) ★
'     全段階 64bit 専用 (#Else #Error 維持) のまま実装したが、
'     新規コードの LongLong 依存箇所には「32bit 化ポイント」コメントを残す。
'     サンク機械語の x86 化という大規模課題は別段階 (将来) に切り出す。
'     EventRegistrationToken は x86 でも 64bit (LongLong) なので両対応で問題なし。
'
'   第9.9a の変更点 (継続):
'     - FillGUID を Private → Public に昇格
'         (Wv2Pane.EnsureView2 がローカル GUID 構造体を初期化するために必要)
'     - ICoreWebView2_2 取り出し基盤を Wv2Pane 側に実装 (m_pView2 + EnsureView2)
'
'   第9.8c の変更点 (継続):
'     - HandlerKind に HK_ExecuteScriptCompleted = 8 を追加
'         (1 ショット系、同時並行可、案 a 採用)
'     - m_iidTable の配列上限を HK_NewWindowRequested → HK_ExecuteScriptCompleted に拡張
'     - InitIIDTable に IID_ICoreWebView2ExecuteScriptCompletedHandler を追加
'         (49511172-CC67-4BCA-9923-137112F4C4CC)
'     - Handler_QueryInterface / Handler_AddRef / Handler_Release / Handler_Invoke
'       本体のロジックは無変更 (kind は動的アクセスなので新 kind=8 も自動的に通る)
'
'   第9.8a/9.8b の変更点 (継続):
'     - モジュール名 Module9_9_9 → Wv2Thunks に正式固定 (9.8a)
'     - 配列型 m_handlers() As ComCallbackHandler (9.8a)
'     - 第9.8b では Wv2Thunks の変更なし (Settings 系は Wv2Pane 内で完結)
'
'   モジュールの役割:
'     WebView2 ハンドラの実体を支える基盤層。サンク領域の確保 (VirtualAlloc)、
'     スロット管理 (m_handlers / m_freeSlots)、参照カウント (HandlerAddRefInternal /
'     HandlerReleaseInternal)、サンクから飛んでくる Handler_QueryInterface /
'     Handler_AddRef / Handler_Release / Handler_Invoke、IID チェック
'     (m_iidTable + IsEqualGUIDInPlace、第9.7a)、センチネル機構による
'     VBA リセット時の領域回収 (第9.5)、Win32 API 宣言の集約 (第9.7b で
'     マウスリサイズ用 API も追加) を担う。
'
'   過去段階の主要な確立内容 (詳細は開発メモ.md 参照):
'     第9.2: サンク方式の確立 (DispCallFunc → VBA メソッドへの折り返し)
'     第9.3: 上位クラス (Wv2Environment / Wv2Pane) の導入、案 D1-α
'     第9.4: 永続ハンドラ機構 (案 P1)、HandlerKind 連動配列 (案 r)
'     第9.5: センチネル機構による領域リーク完全解消、&H リテラル罠の発見
'            (MEM_RELEASE = &H8000& 修正)
'     第9.6: NewWindowRequested 横展開、出自接頭辞による命名規約 (View_/Ctrl_)
'     第9.7a: Handler_QueryInterface の真正化 (IID チェック導入)
'     第9.7b: UserForm マウスリサイズ用 API (GetAncestor / SetWindowLongPtrW /
'             SetWindowPos) と関連定数 (GA_ROOT / WS_THICKFRAME 等)
'
'   ★ 第9.7a で確立した重要な事実 ★
'     Test_AllEvents_Real のログ観察により、以下が判明している:
'     1. WebView2 ランタイムは登録済みイベントハンドラに対して、
'        IID_IUnknown と本来の IID 以外を QueryInterface で要求しない
'     2. WebView2 は「登録時に QI して vtable をキャッシュ、Invoke 時には
'        QI せず直接 Invoke」というパターンを採用している
'     3. 9.6b 以前の「全 IID 素通し」実装で実害が出なかった真因は、
'        WebView2 が QI を最小限しか叩かないため、危険な「嘘の IMarshal
'        ポインタを返す」場面そのものが発生していなかった
'
'   ★ トラブルシューティング指針 (第9.7a 以降) ★
'     もし不可解な挙動 (= ハンドラ未登録、Invoke が呼ばれない、
'     AddRef/Release 不整合、Cleanup 時にハング、等) が出た場合:
'       (1) Handler_QueryInterface 冒頭で
'           「ppvObject = this : HandlerAddRefInternal this :
'            Handler_QueryInterface = S_OK : Exit Function」
'           を一時的に追加して 9.6b 相当に戻す
'       (2) この状態で機能が直るなら、IID チェック起因と確定
'       (3) 直らないなら IID チェックは無関係、別の原因を探す
'
'   将来の課題 (第9.9c 完了 = 第9.9 全完了時点で残るもの):
'     - TabWebView2 包括クラス (第9.10 で着手予定、最大規模)
'     - 実 SPA テスト環境 (http(s) origin での pushState/replaceState 本格検証、9.10 以降)
'     - バージョン拡張 IF (ICoreWebView2_3 以降、必要時に追加)
'     - 32bit (x86) 対応: サンク機械語の x86 化 + 型抽象化 (将来の別段階)
'     ※ SPA 対応 (HistoryChanged / DOMContentLoaded / postMessage 規約サンプル) と
'        ExecuteScript 補完群 (D + E + F + G) は本セッション (第9.9c) までで完了。
''''''''''''''''''''''''''''''''''

''
' PointerAccessor / SafeArray logic
' Copyright (c) 2025 Cristian Buse
' Licensed under the MIT License
' https://github.com/WNKLER/refTypes/discussions/3
''

Option Explicit

#If Win64 Then
    Private Const NullPtr As LongLong = 0^
    Private Const PtrSize = 8
#Else
    #Error "このモジュールは 64 ビット VBA (x64) が必要です"
#End If

' --- SafeArray / PointerAccessor 用構造体 ---
Private Enum SAFEARRAY_FEATURES
    FADF_AUTO = &H1
    FADF_FIXEDSIZE = &H10
End Enum

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As LongPtr
    rgsabound0 As SAFEARRAYBOUND
End Type

Private Type PointerAccessor
    arr() As LongPtr
    sa As SAFEARRAY_1D
End Type

' --- WebView2Loader.dll ---
'   第9.3a 段階で Public Declare に変更 (Wv2Environment.Init から呼ぶため)
Public Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" ( _
    ByVal browserExecutableFolder As LongPtr, _
    ByVal userDataFolder As LongPtr, _
    ByVal additionalBrowserArguments As LongPtr, _
    ByVal environmentCreatedHandler As LongPtr) As Long

' --- Win32 API ---
Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, _
    ByVal dwSize As LongPtr, _
    ByVal flAllocationType As Long, _
    ByVal flProtect As Long) As LongPtr

Private Declare PtrSafe Function VirtualFree Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, _
    ByVal dwSize As LongPtr, _
    ByVal dwFreeType As Long) As Long

' --- DispCallFunc (oleaut32) ---
'   第9.3a 段階で新規追加。
'   ICoreWebView2Environment::Release などを VBA 側から呼ぶための
'   最小ヘルパー (ComRelease / ComAddRef) で使用する。
'
'   ByRef ... As Any のパラメータには、引数 0 個のメソッドを呼ぶ際に
'   呼び出し側で「ByVal 0&」と書いて NULL を渡す。
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" ( _
    ByVal pvInstance As LongPtr, _
    ByVal oVft As LongPtr, _
    ByVal cc As Long, _
    ByVal vtReturn As Integer, _
    ByVal cActuals As Long, _
    ByRef prgvt As Any, _
    ByRef prgpvarg As Any, _
    ByRef pvargResult As Any) As Long

' --- user32: GetClientRect (第9.3c 段階で追加) ---
'   HWND のクライアント領域を取得する。位置は常に (0,0,width,height)。
'   put_Bounds の引数として直接そのまま使える。
'   Wv2Pane から呼ばれる前提で Public Declare。
Public Declare PtrSafe Function GetClientRect Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByRef lpRect As RECT) As Long

' --- UserForm マウスリサイズ用 API 群 (第9.7b 段階で追加) ---
'   UserForm1.frm 側で「Frame1 の HWND → UserForm 自身の HWND」を取得し、
'   UserForm 自身の HWND に WS_THICKFRAME (= 太枠 = マウスリサイズ可) を
'   付与するために使う。
'
'   Declare は Module9 に集約 (設計原則的に Win32 API は Module9 で一括管理)。
'   定数も Public Const として Module9 に置く (UserForm1 から参照可能)。
'
'   GetAncestor:
'     指定した HWND の祖先 HWND を取得する。
'     gaFlags = GA_ROOT (=2) を渡すと、その HWND を含む「ルートウィンドウ」
'     (= デスクトップから直接見えるトップレベル/ポップアップウィンドウ) が
'     得られる。Frame1 → UserForm 自身の HWND の取得に使う。
'
'   GetWindowLongPtrW / SetWindowLongPtrW:
'     ウィンドウの拡張情報 (Window Style 等) を取得/設定する。
'     第二引数 nIndex に GWL_STYLE (=-16) を渡すと、ウィンドウスタイルを
'     LongPtr で取得/設定できる (x64 では LONG_PTR 幅、x86 では LONG 幅、
'     PtrSafe + LongPtr で抽象化されている)。
'     WS_THICKFRAME (=&H40000) を OR で立てると、マウスでリサイズ可能になる。
'
'   SetWindowPos:
'     ウィンドウの位置/サイズ/Z オーダー/見た目を更新する。
'     WS_THICKFRAME 付与後、SWP_FRAMECHANGED を渡して呼ぶことで、
'     非クライアント領域の再描画 (= 太枠の表示反映) をトリガする。
'     位置/サイズ/Z オーダーは変えたくないので SWP_NOMOVE | SWP_NOSIZE |
'     SWP_NOZORDER も同時に立てる。
Public Declare PtrSafe Function GetAncestor Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal gaFlags As Long) As LongPtr

Public Declare PtrSafe Function GetWindowLongPtrW Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal nIndex As Long) As LongPtr

Public Declare PtrSafe Function SetWindowLongPtrW Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As LongPtr) As LongPtr

Public Declare PtrSafe Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal hWndInsertAfter As LongPtr, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal uFlags As Long) As Long

' --- UserForm マウスリサイズ関連の定数 (第9.7b 段階で追加) ---
'   GetAncestor / GetWindowLongPtrW / SetWindowLongPtrW / SetWindowPos で使う
'   定番の値。Win32 標準のシンボル名と一致させる。
'
'   いずれも Integer 範囲を超えないが、設計原則 29 の罠に倣って
'   念のため &H... には & サフィックスを付ける (Integer に解釈される
'   可能性を排除し、Long 確定として扱う)。
Public Const GA_ROOT          As Long = 2
Public Const GWL_STYLE        As Long = -16
Public Const WS_THICKFRAME    As Long = &H40000
Public Const SWP_NOSIZE       As Long = &H1&
Public Const SWP_NOMOVE       As Long = &H2&
Public Const SWP_NOZORDER     As Long = &H4&
Public Const SWP_FRAMECHANGED As Long = &H20&

' --- 文字列ヘルパー用 API (第9.4a 段階で追加、旧 mWebView2Helper 由来) ---
'   LPWSTR (UTF-16 ヌル終端文字列) を VBA String に変換するために使用。
'   各メソッドの用途:
'     lstrlenW       : LPWSTR の文字数 (NULL 終端を除く) を取得
'     CoTaskMemFree  : COM の out 引数 LPWSTR* (CoTaskMemAlloc で確保された
'                      文字列) を解放する。COM 規約上、呼び出し側の責任。
'     RtlMoveMemory  : メモリ間コピー (旧版踏襲のため使用)
Public Declare PtrSafe Function lstrlenW Lib "kernel32" ( _
    ByVal lpString As LongPtr) As Long

Public Declare PtrSafe Sub CoTaskMemFree Lib "ole32" ( _
    ByVal pv As LongPtr)

Public Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" ( _
    ByVal destination As LongPtr, _
    ByVal source As LongPtr, _
    ByVal length As LongPtr)

' --- 環境変数 API (第9.5 段階で追加、センチネル機構用) ---
'   プロセス環境変数に「最後に Thunks_Init で確保した領域のベースアドレス」を
'   十進文字列で保存する。MouseScroll.bas (Cristian Buse 流) と同じ流派。
'
'   W 版 (UTF-16) を採用。StrPtr() で VBA String の内部 UTF-16 バッファを
'   そのまま渡せるため、A/W 変換が挟まらない。
'
'   SetEnvironmentVariableW(lpName, NULL) は環境変数の削除を意味する。
'   VBA から NULL ポインタを渡すには第二引数を 0 で呼び出せばよい
'   (LongPtr の ByVal なので 0 が NULL として扱われる)。
Private Declare PtrSafe Function GetEnvironmentVariableW Lib "kernel32" ( _
    ByVal lpName As LongPtr, _
    ByVal lpBuffer As LongPtr, _
    ByVal nSize As Long) As Long

Private Declare PtrSafe Function SetEnvironmentVariableW Lib "kernel32" ( _
    ByVal lpName As LongPtr, _
    ByVal lpValue As LongPtr) As Long

' --- Win32 エラー取得 (第9.5 段階で追加、診断目的) ---
'   直前の Win32 API 呼び出しが失敗した際のエラーコードを取得する。
'   VirtualFree が 0 を返した原因の切り分け (ERROR_BUSY=170 /
'   ERROR_INVALID_ADDRESS=487 / ERROR_INVALID_PARAMETER=87 など) に使う。
'
'   注意: GetLastError は per-thread の状態を持つ。VBA から Win32 API を
'         呼ぶたびに更新されるので、観察したい API の直後に取得すること。
Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

' --- VirtualQuery 用構造体 (第9.5 段階で追加) ---
'   x64 における MEMORY_BASIC_INFORMATION のレイアウト (合計 48 bytes):
'     +0  BaseAddress       LongPtr (8)
'     +8  AllocationBase    LongPtr (8)
'     +16 AllocationProtect Long    (4)
'     +20 (pad1)            Long    (4) ← アラインメント用パディング
'     +24 RegionSize        LongPtr (8)
'     +32 State             Long    (4)
'     +36 Protect           Long    (4)
'     +40 Type              Long    (4)  (VBA 予約語と紛れるので Type_ と命名)
'     +44 (pad2)            Long    (4) ← 末尾アラインメント
'   VBA Type は宣言順 + 自動アラインメントで配置されるので、明示的に
'   pad1/pad2 を挟むことでドキュメント上の正確性も確保する。
Private Type MEMORY_BASIC_INFORMATION
    BaseAddress       As LongPtr
    AllocationBase    As LongPtr
    AllocationProtect As Long
    pad1              As Long
    RegionSize        As LongPtr
    state             As Long
    Protect           As Long
    Type_             As Long
    pad2              As Long
End Type

' --- VirtualQuery (第9.5 段階で追加、診断目的) ---
'   指定アドレスを含むメモリページの状態を取得する。
'   戻り値はコピーされた MEMORY_BASIC_INFORMATION のバイト数 (失敗時 0)。
'   Type 定義を先に置く必要があるため、Declare はこの直後に書く。
Private Declare PtrSafe Function VirtualQuery Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, _
    ByRef lpBuffer As MEMORY_BASIC_INFORMATION, _
    ByVal dwLength As LongPtr) As LongPtr

' MEMORY_BASIC_INFORMATION.State の値:
'   MEM_COMMIT  = &H1000  (4096)   ← 既存の MEM_COMMIT と同じ値
'   MEM_RESERVE = &H2000  (8192)   ← 既存の MEM_RESERVE と同じ値
'   MEM_FREE    = &H10000 (65536)
' MEMORY_BASIC_INFORMATION.Protect の値は FormatMemProtect 内で個別判定するため
' 個別の定数定義は省略。

' --- RECT 構造体 (第9.3c 段階で追加) ---
'   ICoreWebView2Controller::put_Bounds に渡される 16 bytes 構造体。
'   x64 ABI 的には「16 bytes 以下の構造体は値渡し」だが、実機検証では
'   旧版が VarPtr(rect) でポインタ渡しして動作していたため、本実装も
'   同じ方式 (= ポインタ渡しを dcf の LongPtr 引数として渡す) を踏襲。
'   Wv2Pane から参照されるので Public Type。
Public Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Const MEM_COMMIT             As Long = &H1000&
Private Const MEM_RESERVE            As Long = &H2000&
' ↓ 9.5g 段階で「&H8000」→「&H8000&」に修正。サフィックスなしだと
'   Integer の &H8000 (= -32768) として評価された後 Long に拡張されるため、
'   VirtualFree に -32768 を渡してしまい ERROR_INVALID_PARAMETER で失敗していた。
'   この罠は VBA 7 + x64 環境で長年踏まれ続けている古典的バグ。
Private Const MEM_RELEASE            As Long = &H8000&
' MEM_FREE (第9.10a 追加): VirtualQuery の State 判定用。解放済み領域を表す。
'   &H10000 = 65536。サフィックス & を必ず付ける (MEM_RELEASE の罠と同種の予防)。
Private Const MEM_FREE               As Long = &H10000
Private Const PAGE_EXECUTE_READWRITE As Long = &H40&
' PAGE_READWRITE (第9.10a 追加): Sentinel ガードで「読み書き可能か」を判定する用。
'   サンク領域は PAGE_EXECUTE_READWRITE (&H40) で確保されるが、健在判定では
'   念のため PAGE_READWRITE (&H4) も読み書き可として許容する。
Private Const PAGE_READWRITE         As Long = &H4&
Private Const S_OK                   As Long = 0
Private Const E_NOINTERFACE          As Long = &H80004002
Private Const CC_STDCALL             As Long = 4

' --- センチネル機構用定数 (第9.5 段階で追加) ---
'
'   SENTINEL_ENV_NAME:
'     プロセス環境変数の名前。1 マシン上の複数 Excel ワークブック間で
'     同じ名前を使うと衝突するが、その場合「片方の領域が回収できない」
'     だけでクラッシュはしない (= リーク許容に戻るだけ)。
'
'   SENTINEL_BUFFER_SIZE:
'     GetEnvironmentVariableW で受け取るバッファの文字数 (WCHAR 単位)。
'     x64 のアドレスを CStr で十進文字列化すると最大 19 桁
'     (2^63 - 1 = 9223372036854775807)、ヌル終端を含めても 20 文字あれば
'     足りるが、安全マージンを取って 32。
Private Const SENTINEL_ENV_NAME    As String = "WV2_VBA_LastRegion"
Private Const SENTINEL_BUFFER_SIZE As Long = 32

' --- スロット / 領域レイアウト定数 ---

' EntryPoint スタブ内で「次に呼ぶ先のアドレス」が即値として埋め込まれる位置
Private Const LATE_BIND_OFFSET As Long = 55

' スタブの実体長 (第8.5段階で確認、+88 の C2 00 00 まで含めて 91 bytes)
Private Const STUB_LEN As Long = 91

' サンクの長さ (フラグチェック 18 + 既存 56 = 74 bytes)
Private Const THUNK_LEN As Long = 74

' サンク領域のオフセット (スタブ 91 + パディング 5 = 96)
Private Const THUNK_OFFSET   As Long = 96

' vtableObj 領域のオフセット (= 旧 SLOT_SIZE と同じ値で覚えやすい)
'   この位置から WebView2VTable 構造体 (40 bytes) を配置:
'     +0..+7   : pVTable (= 自身の +8 を指す)
'     +8..+15  : Functions(0) = Handler_QueryInterface
'     +16..+23 : Functions(1) = Handler_AddRef
'     +24..+31 : Functions(2) = Handler_Release
'     +32..+39 : Functions(3) = pSlot (スタブクローン先頭)
'
'   第9.3a 段階で Public Const に変更 (Wv2Environment.Init から参照)
Public Const VTABLE_OBJ_OFFSET As Long = 176

' 1 スロットあたりの確保サイズ (8 の倍数に揃えて 224 bytes)
'   +0..+90      スタブクローン (91 bytes)
'   +91..+95     パディング
'   +96..+169    サンク (74 bytes)
'   +170..+175   パディング (8 byte アラインメント)
'   +176..+183   vtableObj.pVTable
'   +184..+215   vtableObj.Functions(0..3)
'   +216..+223   予備パディング (8 byte アラインメント)
Private Const SLOT_SIZE      As Long = 224

' スロット数 (固定)
Private Const SLOT_COUNT     As Long = 512

' 領域先頭のヘッダサイズ (生存フラグ + 将来予約)
'   +0          : 生存フラグ (1 byte、ただし 8 byte 単位で扱う)
'   +1..+63     : 予約 (将来のセンチネル用フラグクリアサンクなど)
Private Const HEADER_SIZE    As Long = 64

' 領域全体のサイズ (= 64 + 224×512 = 114,752 bytes ≒ 112 KB)
Private Const REGION_SIZE    As Long = HEADER_SIZE + SLOT_SIZE * SLOT_COUNT


' ============================================================
' GUID 構造体 (Win32 標準、第9.7a で新規追加)
'
'   COM の REFIID として渡される 16 byte の構造体。
'   QueryInterface の riid 引数はこの構造体へのポインタ。
'
'   メモリ上の x64 レイアウト (16 bytes、Win32 標準):
'     +0..+3   : Data1     (Long、little-endian)
'     +4..+5   : Data2     (Integer、little-endian)
'     +6..+7   : Data3     (Integer、little-endian)
'     +8..+15  : Data4(0..7) (各 Byte をそのまま、big-endian 風)
'
'   例: IID {4e8a3389-c9d8-4bd2-b6b5-124fee6cc14d} は
'     Data1 = &H4E8A3389
'     Data2 = &HC9D8 (Integer なので符号付き表現で書く際は &HC9D8 のまま)
'     Data3 = &H4BD2
'     Data4 = (&HB6, &HB5, &H12, &H4F, &HEE, &H6C, &HC1, &H4D)
'
'   GUID 比較は memcmp 相当 (16 byte 全体の二進一致) で行う。本モジュールでは
'   LongLong を 2 個読んで比較する高速版を IsEqualGUIDInPlace で実装。
'
'   Public にしているのは、上位クラスで派生 IF QueryInterface を発行する
'   将来 (第9.8 以降のバージョン拡張 IF 等) で使うかも知れないため。
' ============================================================
Public Type GUID
    data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type


' ============================================================
' HandlerKind 列挙型 (上位モジュールから参照可能)
'   WebView2 のコールバック種別を識別する。ComCallbackHandler.cls の
'   m_kind に格納され、Handler_Invoke 内で分岐に使われる。
' ============================================================
Public Enum HandlerKind
    HK_None = 0
    HK_EnvironmentCompleted = 1
    HK_ControllerCompleted = 2
    HK_NavigationStarting = 3
    HK_NavigationCompleted = 4
    HK_WebMessageReceived = 5
    HK_DocumentTitleChanged = 6
    HK_NewWindowRequested = 7
    HK_ExecuteScriptCompleted = 8   ' 第9.8c で追加、1 ショット系、同時並行可
    HK_HistoryChanged = 9           ' 第9.9b で追加、永続、基底 ICoreWebView2 (vtable 13/14)
    HK_DOMContentLoaded = 10        ' 第9.9b で追加、永続、ICoreWebView2_2 (vtable 64/65)
End Enum


' --- 起動時に 1 回取得して保持する関数アドレス ---
'   各スロットの vtableObj.Functions(0..2) に書き込む値。
'   AddressOf 演算子は実引数位置でしか書けないため、GetAddr 経由で取得する。
'   Thunks_Init で 1 度だけ取得し、Acquire 時にはこれを使い回す。
Private m_pHandler_QI       As LongPtr
Private m_pHandler_AddRef   As LongPtr
Private m_pHandler_Release  As LongPtr

' --- スロットプール状態 ---
Private m_pRegionBase As LongPtr        ' VirtualAlloc で確保した領域の先頭
Private m_freeHead    As Long           ' フリーリストの先頭インデックス (-1 で空)
Private m_freeNext()  As Long           ' freeNext(i) = i 番スロットの次の空きスロット
Private m_inUse       As Long           ' 現在使用中のスロット数 (デバッグ用)

' --- スロット index → Handler オブジェクトの対応表 ---
'   Handler_AddRef / Release / QueryInterface から
'   「this (= pSlot + VTABLE_OBJ_OFFSET) → idx (SlotIndexFromVTableObjAddr)
'    → m_handlers(idx)」の経路で対応する ComCallbackHandler インスタンスに到達する。
'   refcount = 0 の自動解放経路でもこの配列をクリアする。
Private m_handlers(0 To SLOT_COUNT - 1) As ComCallbackHandler

' --- IID テーブル (第9.7a 段階で新規追加) ---
'   HandlerKind ごとの「自分の本来の IID」を保持する。
'   Handler_QueryInterface が riid をこれと比較して、一致すれば
'   S_OK + 自身、不一致なら ppvObject = 0 + E_NOINTERFACE を返す。
'
'   Thunks_Init の末尾で InitIIDTable によって初期化される。
'   宣言のレンジは HandlerKind の全範囲 (HK_None = 0 ? HK_DOMContentLoaded = 10)。
'   HK_None のエントリは未使用 (ハンドラとして使われない、初期値 0 のまま)。
Private m_iidTable(HK_None To HK_DOMContentLoaded) As GUID

' --- IID_IUnknown (Win32 標準、Thunks_Init で初期化) ---
'   {00000000-0000-0000-C000-000000000046}。Handler_QueryInterface での
'   比較に使う。値はハードコードされた COM 標準値。
Private m_iidIUnknown As GUID

Private Const THUNK_BUF_SIZE As Long = 80     ' 74 を 8 の倍数に切り上げ


' ============================================================
' EntryPoint スタブのソース
'   このスタブのバイト列を VirtualAlloc 領域へコピーする。
'   Force-Compile のため起動時に 1 回呼び出しておく必要がある。
' ============================================================
Private Sub EntryPoint(): End Sub
' ============================================================
' AcquireHandlerFor (第9.3a 段階で新規)
'
'   上位クラス (Wv2Environment など) が「自分専用のハンドラを 1 個欲しい」
'   ときに呼ぶ統合エントリ。以下を一括で実行する:
'     1. Thunks_Init() を呼ぶ (まだなら)
'     2. ComCallbackHandler を New
'     3. ComCallbackHandler.Handler_Invoke のアドレスを vTable から取得
'     4. Thunks_AcquireSlot を呼んでスロット確保 + サンク書き込み + m_handlers 登録
'     5. ComCallbackHandler.Init(kind, owner, pSlot) を呼ぶ
'     6. 初期化済みの ComCallbackHandler を返す
'
'   呼び出し側は戻り値を強参照で受け取り、自分のフィールドに格納する。
'   その後、上位クラスは ComCallbackHandler.Slot + VTABLE_OBJ_OFFSET を WebView2 に
'   渡す形でコールバック登録を行う。
'
'   失敗時 (プール枯渇等) は Nothing を返す。呼び出し側は Nothing チェック必須。
' ============================================================
Public Function AcquireHandlerFor( _
    ByVal kind As HandlerKind, _
    ByVal owner As Object) As ComCallbackHandler

    ' --- 1. プール初期化 ---
    If m_pRegionBase = 0 Then
        If Not Thunks_Init() Then Exit Function
    End If

    ' --- 2. ComCallbackHandler を New ---
    Dim h As ComCallbackHandler
    Set h = New ComCallbackHandler

    ' --- 3. ComCallbackHandler.Handler_Invoke のアドレスを取得 ---
    Dim pHandlerInvoke As LongPtr
    pHandlerInvoke = GetClassMethodAddrAtFixedSlot(h, 7)
    If pHandlerInvoke = 0 Then Exit Function

    ' --- 4. スロット確保 + サンク書き込み + m_handlers 登録 ---
    Dim pSlot As LongPtr
    pSlot = Thunks_AcquireSlot(h, ObjPtr(h), pHandlerInvoke)
    If pSlot = 0 Then Exit Function

    ' --- 5. ComCallbackHandler 自身を初期化 ---
    h.Init kind, owner, pSlot

    ' --- 6. 初期化済み ComCallbackHandler を返す ---
    Set AcquireHandlerFor = h
End Function


' ============================================================
' EnsureFolder (第9.3a 段階で新規)
'
'   ユーザデータフォルダなどの存在確認 + 自動作成。
'   Wv2Environment.Init から呼ばれる。
'
'   "C:\Temp\VBA_WebView2" のように複数階層の場合でも 1 階層だけ親を作る。
'   より深い階層が必要なら呼び出し側で工夫する (現状は 1 階層で足りる前提)。
' ============================================================
Public Sub EnsureFolder(ByVal path As String)
    If LenB(path) = 0 Then Exit Sub

    ' 親ディレクトリ
    Dim parentDir As String
    Dim slashPos As Long
    slashPos = InStrRev(path, "\")
    If slashPos > 0 Then
        parentDir = Left$(path, slashPos - 1)
        If LenB(parentDir) > 0 Then
            If Dir(parentDir, vbDirectory) = "" Then MkDir parentDir
        End If
    End If

    ' 自身
    If Dir(path, vbDirectory) = "" Then MkDir path
End Sub


' ============================================================
' FinalizeWebView2Environment (第9.3a 段階で簡素化)
'
'   旧版で gHandler を Nothing にする処理が含まれていたが、
'   gHandler が撤廃されたため、Thunks_Shutdown を呼ぶだけになった。
'
'   呼び出すのは「全てのハンドラ・上位クラスを片付けた後の最終クリーンアップ」
'   としてのみ。通常のテストフローでは Test 内で個別に Thunks_Shutdown を
'   呼ぶ方針 (Test_TwoHandlers_Mock がそうしている)。
' ============================================================
Public Sub FinalizeWebView2Environment()
    Thunks_Shutdown
End Sub


' ============================================================
' dcf (第9.3b 段階で新規)
'
'   汎用 DispCallFunc ラッパー。可変長引数 (ParamArray) で
'   引数 0?N 個の COM メソッドを叩ける。
'
'   旧版の dcf からの整理点:
'     - 「CopyMemory による pVTable / pMethod の二重取得」を削除。
'       これは取得した値が一切使われないデッドコードだった (DispCallFunc
'       が内部で同じことをやってくれる)。
'     - 「Case Else 内の CLngPtr(CLng(...))」を削除。直前の
'       Case vbLong: 分岐で既に処理済みのため到達不能だった。
'
'   引数:
'     pInterface : COM インターフェースポインタ (this)
'     vtblIndex  : vtable のスロット番号 (IUnknown は 0=QI, 1=AddRef, 2=Release、
'                  以降のメソッドは派生インターフェースの宣言順に 3 から並ぶ)
'     funcName   : デバッグ用 (空文字なら出力しない)
'     args       : メソッドに渡す可変長引数 (LongPtr/Long/Double/その他)
'
'   引数の型推論ルール (VarType ベース):
'     vbLong       → 4 bytes 整数として渡す
'     vbLongLong   → 8 bytes ポインタ/整数として渡す (Win64 LongPtr 互換)
'     vbDouble     → 8 bytes 倍精度浮動小数として渡す
'     その他全部   → vbLongLong 扱い (LongPtr フォールバック)
'
'   戻り値: 呼び出した COM メソッドの戻り値 (Long にキャスト)
'           ・DispCallFunc 自体が失敗したらその hr を返し、Debug.Print する
'           ・成功時は COM メソッドの戻り値をそのまま返す
'             (HRESULT のときも、AddRef/Release のような ULONG 戻り値の
'              ときも、呼び出し側がそのまま受け取る)
'           ・「戻り値 = HRESULT」を仮定した自動エラー判定はしない:
'             AddRef/Release は新しい参照カウントを返すため、
'             値 ≠ 0 を一律でエラー扱いにすると誤検知が出る (実機で確認済み)。
'             失敗判定は呼び出し側で「これは HRESULT を返すメソッドだ」と
'             分かっている場合に行う。
'
'   注意: out 引数 (例: pp* 系の二重ポインタ) を含むメソッドを呼ぶ場合は、
'         呼び出し側で Variant に格納した LongPtr 変数の VarPtr を渡す前提。
'         Variant に格納した値が変更されて返るので、呼び出し後に CLngPtr 等で
'         取り出す。第9.3c 以降の Controller::get_CoreWebView2 などで使う。
' ============================================================
Public Function dcf( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    ByVal funcName As String, _
    ParamArray args() As Variant) As Long

    If pInterface = 0 Then
        Debug.Print "dcf: null interface - " & funcName
        dcf = &H80004003   ' E_POINTER
        Exit Function
    End If

    Dim argc As Long
    argc = UBound(args) - LBound(args) + 1
    If argc < 0 Then argc = 0

    Dim res As Variant
    Dim hr As Long

    If argc = 0 Then
        ' 引数なし: vt/vp に NULL を渡す
        hr = DispCallFunc(pInterface, vtblIndex * PtrSize, _
                          CC_STDCALL, vbLong, _
                          0, ByVal 0&, ByVal 0&, res)
    Else
        ' 引数あり: vt / vp 配列を構築
        Dim vt() As Integer
        Dim vp() As LongPtr
        Dim vals() As Variant
        ReDim vt(0 To argc - 1)
        ReDim vp(0 To argc - 1)
        ReDim vals(0 To argc - 1)

        Dim i As Long
        For i = 0 To argc - 1
            vals(i) = args(LBound(args) + i)
            Select Case VarType(vals(i))
                Case vbLong:     vt(i) = vbLong
                Case vbLongLong: vt(i) = vbLongLong   ' = vbLongPtr on Win64
                Case vbDouble:   vt(i) = vbDouble
                Case Else:       vt(i) = vbLongLong   ' フォールバック (LongPtr 扱い)
            End Select
            vp(i) = VarPtr(vals(i))
        Next i

        hr = DispCallFunc(pInterface, vtblIndex * PtrSize, _
                          CC_STDCALL, vbLong, _
                          argc, vt(0), vp(0), res)
    End If

    If hr <> 0 Then
        If LenB(funcName) > 0 Then _
            Debug.Print "dcf CALL failed: " & funcName & " hr=&H" & Hex(hr)
        dcf = hr
    Else
        ' 成功時は COM メソッドの戻り値をそのまま返す。
        ' HRESULT を返すメソッドの場合、呼び出し側が <> 0 をチェックしてエラー処理する。
        ' AddRef/Release のような ULONG 戻り値の場合、新しい参照カウントが返る
        ' (ここで一律に「<> 0 ならエラー」とは判定しない)。
        dcf = CLng(res)
    End If
End Function


' ============================================================
' ComRelease / ComAddRef
'
'   IUnknown::Release / AddRef を呼ぶ薄いラッパ。
'   第9.3b で dcf に統一されたが、可読性のため Public Function は残す
'   (Wv2Environment.Class_Terminate などからの呼び出しで「Release を撃っている」
'   という意図を明示できるため)。
' ============================================================
Public Function ComRelease(ByVal pInterface As LongPtr) As Long
    If pInterface <> 0 Then ComRelease = dcf(pInterface, 2, "Release")
End Function

Public Function ComAddRef(ByVal pInterface As LongPtr) As Long
    If pInterface <> 0 Then ComAddRef = dcf(pInterface, 1, "AddRef")
End Function


' ============================================================
' PtrToString (第9.4a 段階で追加、旧 mWebView2Helper 由来)
'
'   LPWSTR (UTF-16 ヌル終端文字列のポインタ) を VBA の String に変換する。
'   COM の out 引数で受け取った LPWSTR を扱うために使う。
'
'   実装メモ:
'     RtlMoveMemory を使った旧版実装をそのまま流用 (案 β)。
'     文字列変換だけは我々の「CopyMemory 不使用」方針から外れるが、
'     PointerAccessor で 1 文字ずつ ChrW で組み立てる方式より高速で、
'     URL 程度の長さでも十分早い。第9.4a 段階での妥協点。
'
'   注意:
'     入力ポインタ p は CoTaskMemAlloc で確保されている前提。
'     呼び出し側は PtrToString で String を取得した後、必ず
'     CoTaskMemFree(p) で解放する責任を持つ。
'     GetStringProperty を使えば取得 → 変換 → 解放まで一括で行える。
' ============================================================
Public Function PtrToString(ByVal p As LongPtr) As String
    If p = 0 Then Exit Function
    Dim cch As Long
    cch = lstrlenW(p)
    If cch = 0 Then Exit Function
    PtrToString = String$(cch, vbNullChar)
    RtlMoveMemory StrPtr(PtrToString), p, CLngPtr(cch * 2)
End Function


' ============================================================
' GetStringProperty (第9.4a 段階で追加、旧 mWebView2Helper 由来)
'
'   COM インターフェースの get_*String 系プロパティを呼び出して String を返す。
'   汎用ヘルパー: vtable レイアウトが HRESULT get_Xxx([out, retval] LPWSTR *value)
'   であるすべてのメソッドで使える。
'
'   流れ:
'     1. dcf で out 引数 LPWSTR* を取得 (pStr に値が書かれる)
'     2. PtrToString で VBA String に変換
'     3. CoTaskMemFree で LPWSTR を解放 (COM 規約上、呼び出し側の責任)
'
'   引数:
'     pInterface : COM ポインタ (例: ICoreWebView2NavigationStartingEventArgs*)
'     vtblIndex  : 取得対象メソッドの vtable index (例: get_Uri は 3)
'     funcName   : デバッグ用 (空文字なら出力しない、HRESULT 失敗時のみ意味あり)
'
'   戻り値: 取得した文字列 (空文字なら取得失敗 or 空文字列)
'
'   第9.4a での使用例:
'     Dim url As String
'     url = GetStringProperty(pArgs, 3, "NavStartingArgs.get_Uri")
' ============================================================
Public Function GetStringProperty( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    Optional ByVal funcName As String = "") As String

    If pInterface = 0 Then Exit Function

    Dim pStr As LongPtr
    Dim hr As Long
    hr = dcf(pInterface, vtblIndex, funcName, VarPtr(pStr))
    If hr = 0 And pStr <> 0 Then
        GetStringProperty = PtrToString(pStr)
        CoTaskMemFree pStr
    End If
End Function


' ============================================================
' GetBoolProperty (第9.4b 段階で追加、旧 mWebView2Helper 由来)
'
'   COM インターフェースの get_*Bool 系プロパティを呼び出して Boolean を返す。
'   vtable レイアウト: HRESULT get_Xxx([out, retval] BOOL *value)
'
'   BOOL は Win32 の型で 4 bytes 整数 (0 = FALSE、それ以外 = TRUE)。
'   VBA の Boolean に変換するには 0 比較が必要。
'
'   第9.4b での使用例:
'     Dim isOk As Boolean
'     isOk = GetBoolProperty(pArgs, 3, "NavCompletedArgs.get_IsSuccess")
' ============================================================
Public Function GetBoolProperty( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    Optional ByVal funcName As String = "") As Boolean

    If pInterface = 0 Then Exit Function

    Dim value As Long
    Dim hr As Long
    hr = dcf(pInterface, vtblIndex, funcName, VarPtr(value))
    If hr = 0 Then GetBoolProperty = (value <> 0)
End Function


' ============================================================
' GetLongProperty (第9.4b 段階で追加、旧 mWebView2Helper 由来)
'
'   COM インターフェースの get_*Long 系プロパティを呼び出して Long を返す。
'   vtable レイアウト: HRESULT get_Xxx([out, retval] long *value)
'
'   列挙型 (例: COREWEBVIEW2_WEB_ERROR_STATUS) を取得するときにも使える。
' ============================================================
Public Function GetLongProperty( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    Optional ByVal funcName As String = "") As Long

    If pInterface = 0 Then Exit Function

    Dim value As Long
    Dim hr As Long
    hr = dcf(pInterface, vtblIndex, funcName, VarPtr(value))
    If hr = 0 Then GetLongProperty = value
End Function


' ============================================================
' GetLongLongProperty (第9.9b 段階で追加)
'
'   COM インターフェースの get_*UINT64 / get_*INT64 系プロパティを呼び出して
'   LongLong (符号付き 64 bit) を返す。
'   vtable レイアウト: HRESULT get_Xxx([out, retval] UINT64 *value)
'
'   用途:
'     ICoreWebView2DOMContentLoadedEventArgs.get_NavigationId (UINT64) の取得。
'     NavigationId は単調増加する識別子で、実用上 LongLong の正の範囲に収まる。
'
'   ★ 32bit 化ポイント (方針 X、案 a) ★
'     32bit VBA には LongLong 型が存在しないため、本関数は #If Win64 で
'     囲んで 64bit 専用としている。32bit 対応段階では、UINT64 を受ける
'     別実装が必要になる。定石は以下のいずれか:
'       (a) Currency 型 (8 byte 固定小数) で受けて 10000 倍を補正
'       (b) Long 2 個 (lo/hi) で受けて手動合成
'       (c) NavigationId は識別用途なので下位 32bit だけ Long で受けて妥協
'     どれを採るかは 32bit 対応段階でまとめて設計する (今は先回りしない)。
' ============================================================
#If Win64 Then
Public Function GetLongLongProperty( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    Optional ByVal funcName As String = "") As LongLong

    If pInterface = 0 Then Exit Function

    Dim value As LongLong
    Dim hr As Long
    hr = dcf(pInterface, vtblIndex, funcName, VarPtr(value))
    If hr = 0 Then GetLongLongProperty = value
End Function
#End If


' ============================================================
' Test_TwoHandlers_Mock (第9.2段階のものをそのまま据え置き)
'
'   第9.2段階の中核機能テスト: 2 個のハンドラを Acquire し、
'   それぞれの vtableObj アドレスから別々に AddRef/Release を撃って、
'   idx の逆引きが正しく動くことを確認する。WebView2 は呼ばない
'   (純粋に逆引きロジックの単体試験)。
'
'   第9.3a 段階での確認ポイント:
'     - gHandler 撤廃後も従来通り pass すること (リグレッション)
'     - owner = Nothing で ComCallbackHandler を Acquire してもエラー無く動くこと
'     - HandlerReleaseInternal の自動解放経路で「gHandler Is h」分岐を
'       削除した影響が無いこと
'
'   このテストは Wv2Environment を一切使わない。AcquireHandlerFor も使わず、
'   従来通り Thunks_AcquireSlot を直接叩く方式を維持する
'   (基盤の単体試験としての性質を保つため)。
' ============================================================
Public Sub Test_TwoHandlers_Mock()

    Debug.Print String(60, "=")
    Debug.Print "Test_TwoHandlers_Mock 開始"
    Debug.Print String(60, "=")

    ' --- 1. プール初期化 ---
    If Not Thunks_Init() Then
        Debug.Print "Thunks_Init 失敗"
        Exit Sub
    End If

    Debug.Print "Thunks_Init OK、m_inUse = " & m_inUse

    ' --- 2. 2 個のハンドラを作成 + Acquire ---
    Dim h1 As ComCallbackHandler, h2 As ComCallbackHandler
    Set h1 = New ComCallbackHandler
    Set h2 = New ComCallbackHandler

    Dim pInvoke1 As LongPtr, pInvoke2 As LongPtr
    pInvoke1 = GetClassMethodAddrAtFixedSlot(h1, 7)
    pInvoke2 = GetClassMethodAddrAtFixedSlot(h2, 7)

    Dim pSlot1 As LongPtr, pSlot2 As LongPtr
    pSlot1 = Thunks_AcquireSlot(h1, ObjPtr(h1), pInvoke1)
    pSlot2 = Thunks_AcquireSlot(h2, ObjPtr(h2), pInvoke2)

    If pSlot1 = 0 Or pSlot2 = 0 Then
        Debug.Print "Acquire 失敗"
        Exit Sub
    End If

    h1.Init HK_EnvironmentCompleted, Nothing, pSlot1
    h2.Init HK_ControllerCompleted, Nothing, pSlot2

    Dim idx1 As Long, idx2 As Long
    idx1 = SlotIndexFromAddr(pSlot1)
    idx2 = SlotIndexFromAddr(pSlot2)
    Debug.Print "h1 -> idx " & idx1 & ", pSlot " & pSlot1
    Debug.Print "h2 -> idx " & idx2 & ", pSlot " & pSlot2
    Debug.Print "m_inUse = " & m_inUse & " (期待値 2)"

    If idx1 = idx2 Then
        Debug.Print "[NG] 異なるハンドラが同じ idx を取得した"
        Exit Sub
    End If

    ' --- 3. 各ハンドラの vtableObj アドレスを直接計算 ---
    Dim pObj1 As LongPtr, pObj2 As LongPtr
    pObj1 = pSlot1 + VTABLE_OBJ_OFFSET
    pObj2 = pSlot2 + VTABLE_OBJ_OFFSET

    ' --- 4. 各々に AddRef を撃って、独立にカウントが上がることを確認 ---
    Debug.Print "--- AddRef sequence ---"
    Dim r As Long
    r = Handler_AddRef(pObj1)
    Debug.Print "  Handler_AddRef(pObj1) -> " & r
    r = Handler_AddRef(pObj1)
    Debug.Print "  Handler_AddRef(pObj1) -> " & r
    r = Handler_AddRef(pObj2)
    Debug.Print "  Handler_AddRef(pObj2) -> " & r
    Debug.Print "h1.RefCount = " & h1.RefCount & " (期待値 2)"
    Debug.Print "h2.RefCount = " & h2.RefCount & " (期待値 1)"

    If h1.RefCount <> 2 Or h2.RefCount <> 1 Then
        Debug.Print "[NG] AddRef の振り分けが期待値と異なる"
    Else
        Debug.Print "[OK] AddRef の振り分けが期待通り"
    End If

    ' --- 5. Release を撃って、各々独立に refcount=0 へ到達するか確認 ---
    Debug.Print "--- Release sequence ---"
    r = Handler_Release(pObj1)
    Debug.Print "  Handler_Release(pObj1) -> " & r
    Debug.Print "h1.RefCount = " & h1.RefCount & " (期待値 1)"
    Debug.Print "m_inUse = " & m_inUse & " (期待値 2、まだ両方使用中)"

    r = Handler_Release(pObj1)
    Debug.Print "  Handler_Release(pObj1) -> " & r
    Debug.Print "(h1 解放後) m_inUse = " & m_inUse & " (期待値 1)"

    Debug.Print "h1.RefCount (解放後) = " & h1.RefCount & " (期待値 0)"
    Debug.Print "h2.RefCount (h1解放と独立) = " & h2.RefCount & " (期待値 1)"

    r = Handler_Release(pObj2)
    Debug.Print "  Handler_Release(pObj2) -> " & r
    Debug.Print "(h2 解放後) m_inUse = " & m_inUse & " (期待値 0)"
    Debug.Print "h2.RefCount (解放後) = " & h2.RefCount & " (期待値 0)"

    ' --- 6. 解放済みスロットへの操作が安全に no-op で通ることを確認 ---
    Debug.Print "--- Post-release safety check ---"
    r = HandlerAddRefInternal(pObj1)
    Debug.Print "解放済み pObj1 への AddRef 戻り値 = " & r & " (期待値 0)"
    r = HandlerReleaseInternal(pObj2)
    Debug.Print "解放済み pObj2 への Release 戻り値 = " & r & " (期待値 0)"

    ' --- 7. 後片付け ---
    Set h1 = Nothing
    Set h2 = Nothing

    Debug.Print "--- 最終状態 ---"
    Debug.Print "m_inUse = " & m_inUse & " (期待値 0)"
    Debug.Print "m_freeHead = " & m_freeHead

    Thunks_Shutdown

    Debug.Print "Test_TwoHandlers_Mock 完了"
    Debug.Print String(60, "=")
End Sub


' ============================================================
' Test_Environment_Real (第9.3a 段階で新規)
'
'   Wv2Environment クラス を使った統合テスト。
'   実 WebView2Loader.dll を経由して Environment 作成 → 完了通知受信 →
'   AddRef → Class_Terminate での Release までの一連の流れを検証する。
'
'   検証項目:
'     1. Wv2Environment.Init で ComCallbackHandler が 1 個 Acquire され、API が呼ばれること
'     2. WebView2 から Invoke が来て ComCallbackHandler.Handler_Invoke 経由で
'        Wv2Environment.OnEnvironmentCompleted が呼ばれること (動的バインディング)
'     3. errorCode = 0、pEnvironment != 0 を受け取り IsReady になること
'     4. m_handler が Nothing に切られること (基盤の自動解放経路が走る)
'     5. m_pEnvironment への AddRef が成功すること (= 1 を返す等の妥当性)
'     6. End Sub で Class_Terminate が走り、ComRelease で参照解放されること
'
'   タイムアウトは 5 秒。WebView2 起動に時間がかかる場合 (初回起動など)
'   は伸ばす必要があるかも。
' ============================================================
Public Sub Test_Environment_Real()

    Debug.Print String(60, "=")
    Debug.Print "Test_Environment_Real 開始"
    Debug.Print String(60, "=")

    Dim env As Wv2Environment
    Set env = New Wv2Environment

    Debug.Print "env.Init を呼びます ..."
    env.Init "C:\Temp\VBA_WebView2"

    If env.IsFailed Then
        Debug.Print "[NG] env.Init が即座に失敗 (LastError = &H" & Hex(env.LastError) & ")"
        GoTo CleanUp
    End If

    Debug.Print "Init 直後の State = " & env.state & _
                " (1=Es_Waiting / 2=Es_Ready 同期コールバック / 3=Es_Failed)"
    Debug.Print "完了通知を待ちます (上限 5 秒、同期コールバック済みなら即抜け) ..."

    ' 待機ループは IsReady / IsFailed のみで判定する。
    ' (Wv2Environment の Public Enum を標準モジュール側から参照することは
    '  処理系依存なので、Bool プロパティ経由の方が安全)
    Dim t As Single: t = Timer
    Do While (Not env.IsReady) And (Not env.IsFailed) And ((Timer - t) < 5#)
        DoEvents
    Loop

    If env.IsReady Then
        Debug.Print "[OK] env.IsReady = True"
        Debug.Print "  m_pEnvironment   = " & env.EnvironmentPtr
        Debug.Print "  経過時間          = " & Format$(Timer - t, "0.000") & " 秒"
    ElseIf env.IsFailed Then
        Debug.Print "[NG] env.IsFailed = True (LastError = &H" & Hex(env.LastError) & ")"
    Else
        Debug.Print "[NG] タイムアウト (state = " & env.state & ")"
    End If

CleanUp:
    Debug.Print "--- env を解放 (Class_Terminate で ComRelease が走る予定) ---"
    Set env = Nothing
    Debug.Print "Set env = Nothing 完了"

    ' 最終的に Thunks_Shutdown でプール領域も解放
    Debug.Print "Thunks_Shutdown 実行"
    Thunks_Shutdown

    Debug.Print "Test_Environment_Real 完了"
    Debug.Print String(60, "=")
End Sub


' ============================================================
' Test_Controller_Real (第9.3b 段階で新規)
'
'   UserForm1 を Show vbModeless で表示し、UserForm1.StartWebView2 を
'   呼んで Wv2Environment → Wv2Pane の生成チェーンを
'   起動する統合テスト。
'
'   第9.3b 段階の検証目標:
'     1. Wv2Environment.Init が同期コールバックで Es_Ready になること (リグレッション)
'     2. Wv2Pane.Init で Environment->CreateCoreWebView2Controller (vtable 3) が
'        dcf 経由で呼ばれること
'     3. WebView2 から HK_ControllerCompleted の Invoke が来て、ComCallbackHandler 経由で
'        Wv2Pane.OnControllerCompleted が呼ばれること
'     4. 偵察ログで this / arg1 / arg2 の値が予測通り (this=obj, arg1=errorCode,
'        arg2=pController) であること
'     5. Controller の参照カウント機構が動くこと (ComAddRef / ComRelease)
'
'   注意:
'     UserForm1 は Show vbModeless で表示されるため、テスト終了後に
'     ユーザが手動で閉じる必要がある。閉じると UserForm_Terminate で
'     m_pane と m_env が解放される (Class_Terminate チェーンが走る)。
' ============================================================
Public Sub Test_Controller_Real()

    Debug.Print String(60, "=")
    Debug.Print "Test_Controller_Real 開始"
    Debug.Print String(60, "=")

    Debug.Print "UserForm1 を Show vbModeless で表示します ..."
    UserForm1.Show vbModeless

    Debug.Print "UserForm1.StartWebView2 を呼び出します ..."
    UserForm1.StartWebView2

    Debug.Print "Test_Controller_Real のメイン処理は完了。"
    Debug.Print "ユーザは UserForm を閉じるとクリーンアップが走ります。"
    Debug.Print String(60, "=")
End Sub


' ============================================================
' Test_Navigate_Real (第9.3c 段階で新規)
'
'   UserForm1 を Show vbModeless で表示し、UserForm1.StartWebView2 を
'   呼び出す。UserForm1.StartWebView2 は (第9.3c 改修版で) 内部で:
'     Environment 生成
'     → Controller 生成
'     → Ctrl_PutBounds (Frame1 のクライアント領域に合わせる)
'     → Ctrl_GetCoreWebView2 (m_pView を取得)
'     → View_Navigate("https://www.bing.com")
'   まで一気に行うため、本テストは Show + StartWebView2 を呼ぶだけで
'   画面表示までのフルパスが検証される。
'
'   第9.3c 段階の検証目標:
'     1. 第9.3b までのリグレッションが通ること
'     2. Ctrl_PutBounds が dcf 経由で成功すること (戻り値が 0)
'     3. Ctrl_GetCoreWebView2 で out 引数が機能すること (m_pView != 0)
'     4. View_Navigate で実際に Bing が表示されること (← 視覚確認)
'     5. UserForm を閉じたときに Wv2Pane.Class_Terminate で
'        m_pView と m_pCtrl の両方が Release されること
'
'   注意: WebView2 ランタイムが「ナビゲーション完了」を返すコールバック
'     は HK_NavigationCompleted で受けるが、これは第9.4 段階で扱う。
'     第9.3c では Navigate を撃つだけで、表示が出るのを目視確認する。
' ============================================================
Public Sub Test_Navigate_Real()

    Debug.Print String(60, "=")
    Debug.Print "Test_Navigate_Real 開始"
    Debug.Print String(60, "=")

    Debug.Print "UserForm1 を Show vbModeless で表示します ..."
    UserForm1.Show vbModeless

    Debug.Print "UserForm1.StartWebView2 を呼び出します ..."
    UserForm1.StartWebView2

    Debug.Print "Test_Navigate_Real のメイン処理は完了。"
    Debug.Print "Bing が表示されたら成功。閉じるボタンでクリーンアップが走ります。"
    Debug.Print String(60, "=")
End Sub


' ============================================================
' Test_AllEvents_Real (第9.4b 段階で新規)
'
'   UserForm1 を Show vbModeless で表示し、UserForm1.StartWebView2 を
'   呼ぶことで Bing 表示 + 3 種類のイベント永続ハンドラ登録を一括実施する
'   (StartWebView2 は第9.4b 段階で改修済み)。
'
'   検証目標:
'     1. リグレッション: Environment / Controller / Navigate が引き続き動作
'     2. 第9.4a の HK_NavigationStarting イベントが発火 (URL ログ出力)
'     3. **新規 HK_NavigationCompleted** イベントが発火 (IsSuccess ログ出力)
'     4. **新規 HK_DocumentTitleChanged** イベントが発火 (title ログ出力)
'     5. 検索ボックスで検索 → 検索結果ページへ遷移 →
'        Starting → Completed → TitleChanged の連鎖を観察
'     6. UserForm を閉じたとき、3 種類すべての永続ハンドラが remove_ される
'        (Wv2Pane.CleanupAllPersistentHandlers の HandlerKind 連動ループが
'         自動的に新種別もカバーする)
'
'   想定ログの流れ (理想):
'     UserForm1.StartWebView2: AddNavigationStarting OK
'     UserForm1.StartWebView2: AddNavigationCompleted OK     ← 新規
'     UserForm1.StartWebView2: AddDocumentTitleChanged OK    ← 新規
'     [Bing 初期ロード時]
'     Wv2Pane.View_OnNavigationStarting: → https://www.bing.com/
'     Wv2Pane.View_OnNavigationCompleted: isSuccess=True            ← 新規
'     Wv2Pane.View_OnDocumentTitleChanged: title=Bing               ← 新規
'     [ユーザが Bing で検索]
'     Wv2Pane.View_OnNavigationStarting: → https://www.bing.com/search?q=...
'     Wv2Pane.View_OnNavigationCompleted: isSuccess=True
'     Wv2Pane.View_OnDocumentTitleChanged: title=検索 - Bing 等
'
'   ナビゲーション制御メソッドの追加検証:
'     Test_AllEvents_Real 完了後、別のテスト (UserForm に View_NavigateToString
'     や View_GoBack を撃つテスト) を実行することで動作確認できる。
'     これは第9.4b の Wv2Pane に Public メソッドとして用意されている。
' ============================================================
Public Sub Test_AllEvents_Real()

    Debug.Print String(60, "=")
    Debug.Print "Test_AllEvents_Real 開始"
    Debug.Print String(60, "=")

    Debug.Print "UserForm1 を Show vbModeless で表示します ..."
    UserForm1.Show vbModeless

    Debug.Print "UserForm1.StartWebView2 を呼び出します ..."
    UserForm1.StartWebView2

    Debug.Print "Test_AllEvents_Real のメイン処理は完了。"
    Debug.Print "Bing が表示されたら、検索ボックスで検索してイベントの連鎖を確認してください。"
    Debug.Print "閉じるボタンでクリーンアップが走ります。"
    Debug.Print String(60, "=")
End Sub


' ============================================================
' Test_MessageEcho_Real (第9.4c 段階で新規)
'
'   WebMessage 双方向通信のフル検証。Bing ではなくテスト用 HTML を
'   NavigateToString で表示し、VBA ←→ JavaScript の往復通信を確認する。
'
'   流れ:
'     1. UserForm1 を Show vbModeless で表示
'     2. UserForm1.StartWebView2_LocalHtml でテスト HTML を表示
'        (Environment / Controller 生成 + 4 種類のイベントハンドラを登録 +
'         BuildTestHtml の HTML を NavigateToString)
'     3. NavigationCompleted を UserForm1.IsPageLoaded で待つ (DoEvents ループ)
'     4. VBA → JS の最初のメッセージを送信:
'        UserForm1.SendMessageToJS "Hello from VBA (auto-sent after page load)"
'        → JS 側の log 領域に表示される (目視確認)
'     5. ユーザが画面の入力欄に文字列を入れて送信ボタンをクリック
'        → JS → VBA の経路で OnWebMessageReceived が発火
'        → Debug.Print で source URL と message body がログに出る
'     6. UserForm を閉じてクリーンアップ
'
'   検証目標:
'     ? View_PostWebMessageAsString で VBA → JS が届く
'     ? Bing 以外のページ (NavigateToString で表示した HTML) も正常表示できる
'     ? JS の window.chrome.webview.postMessage で JS → VBA が届く
'     ? OnWebMessageReceived で args.get_Source と
'       args.TryGetWebMessageAsString が取れる
'     ? 4 種類目のハンドラ追加でも Class_Terminate の一括クリーンアップが動く
'
'   注意:
'     初回ロード前 (IsPageLoaded = False の段階) では PostMessage を撃っても
'     届かない (JS の addEventListener がまだ登録されていない)。
'     必ず IsPageLoaded を待ってから VBA → JS を撃つこと。
' ============================================================
Public Sub Test_MessageEcho_Real()

    Debug.Print String(60, "=")
    Debug.Print "Test_MessageEcho_Real 開始"
    Debug.Print String(60, "=")

    Debug.Print "UserForm1 を Show vbModeless で表示します ..."
    UserForm1.Show vbModeless

    Debug.Print "UserForm1.StartWebView2_LocalHtml を呼び出します (テスト HTML を表示) ..."
    UserForm1.StartWebView2_LocalHtml

    Debug.Print "ページの読み込み完了を待ちます (上限 10 秒) ..."
    Dim t As Single: t = Timer
    Do While (Not UserForm1.IsPageLoaded) And ((Timer - t) < 10#)
        DoEvents
    Loop

    If Not UserForm1.IsPageLoaded Then
        Debug.Print "[NG] ページ読み込みタイムアウト"
        Exit Sub
    End If

    Debug.Print "[OK] ページ読み込み完了"

    Debug.Print "VBA → JS の最初のメッセージを送信します ..."
    UserForm1.SendMessageToJS "Hello from VBA (auto-sent after page load)"

    Debug.Print "Test_MessageEcho_Real のメイン処理は完了。"
    Debug.Print "画面の入力欄から JS → VBA のメッセージを送ってみてください。"
    Debug.Print "Debug.Print に受信内容が出力されます。"
    Debug.Print "閉じるボタンでクリーンアップが走ります。"
    Debug.Print String(60, "=")
End Sub


' ============================================================
' BuildTestHtml (第9.4c 段階で新規)
'
'   双方向通信テスト用の HTML を文字列で返す。
'   UserForm1.StartWebView2_LocalHtml から呼ばれて、
'   View_NavigateToString でブラウザに表示される。
'
'   含まれる機能:
'     - 入力欄 + 送信ボタン: クリックすると window.chrome.webview.postMessage
'       で JS → VBA に送信
'     - log 領域: VBA から受信したメッセージを履歴表示
'     - <script>: WebView2 が注入する window.chrome.webview API を使って
'       双方向通信を実装
'
'   注意:
'     VBA の String リテラルでは " のエスケープが面倒なので、HTML 内の
'     ダブルクォートはシングルクォートで代用したり、必要な部分だけ
'     Chr(34) で挿入する。
' ============================================================
Public Function BuildTestHtml() As String
    Dim s As String
    s = "<!DOCTYPE html>" & vbCrLf
    s = s & "<html><head>" & vbCrLf
    s = s & "<meta charset='utf-8'>" & vbCrLf
    s = s & "<title>WebView2 双方向通信テスト</title>" & vbCrLf
    s = s & "<style>" & vbCrLf
    s = s & "  body { font-family: 'Segoe UI', sans-serif; padding: 20px; }" & vbCrLf
    s = s & "  h1 { color: #0078D4; }" & vbCrLf
    s = s & "  #msg { width: 300px; padding: 4px; }" & vbCrLf
    s = s & "  button { padding: 4px 16px; margin-left: 8px; }" & vbCrLf
    s = s & "  #log { background: #f3f3f3; padding: 10px; min-height: 100px; " & _
                    "white-space: pre-wrap; border: 1px solid #ccc; }" & vbCrLf
    s = s & "</style>" & vbCrLf
    s = s & "</head><body>" & vbCrLf
    s = s & "<h1>WebView2 &lt;-&gt; VBA 通信テスト</h1>" & vbCrLf
    s = s & "<div>" & vbCrLf
    s = s & "  <input type='text' id='msg' placeholder='JS → VBA に送るメッセージ'>" & vbCrLf
    s = s & "  <button onclick='sendToVBA()'>送信</button>" & vbCrLf
    s = s & "</div>" & vbCrLf
    s = s & "<h2>VBA から受信:</h2>" & vbCrLf
    s = s & "<div id='log'></div>" & vbCrLf
    s = s & "<script>" & vbCrLf
    s = s & "  function sendToVBA() {" & vbCrLf
    s = s & "    const text = document.getElementById('msg').value;" & vbCrLf
    s = s & "    window.chrome.webview.postMessage(text);" & vbCrLf
    s = s & "    document.getElementById('msg').value = '';" & vbCrLf
    s = s & "  }" & vbCrLf
    s = s & "  window.chrome.webview.addEventListener('message', event => {" & vbCrLf
    s = s & "    const log = document.getElementById('log');" & vbCrLf
    s = s & "    log.innerText += event.data + '\n';" & vbCrLf
    s = s & "  });" & vbCrLf
    s = s & "</script>" & vbCrLf
    s = s & "</body></html>"

    BuildTestHtml = s
End Function


' ============================================================
' BuildSpaTestHtml (第9.9b 段階で新規、SPA 対応テスト用)
'
'   SPA (Single Page Application) の挙動を模した HTML を返す。
'   UserForm1.StartWebView2_SpaTest から NavigateToString で表示される。
'
'   検証対象イベント:
'     - HistoryChanged   : NavigateToString 直後にランタイムが自動発火する
'                          (実機検証で 2 回連続発火を観測)。これにより
'                          add_HistoryChanged の登録経路が機能していることを確認可。
'                          (基底 ICoreWebView2、vtable 13)
'     - DOMContentLoaded : 初回ロード時に 1 回だけ発火 (navigationId 取得可)
'                          (ICoreWebView2_2、vtable 64)
'     - WebMessageReceived : SPA 規約サンプル (postMessage で JSON 通知)
'                          → 既存の View_OnWebMessageReceived がそのまま受ける
'
'   ★ 実機検証で判明した制約 (第9.9b の知見) ★
'     NavigateToString で表示したページは about:blank origin で動作する。
'     about:blank origin では history.pushState / replaceState が
'     SecurityError で弾かれるのが Web 標準仕様。
'       Uncaught SecurityError: Failed to execute "pushState" on "History":
'         A history state object with URL "" cannot be created in a
'         document with origin "null" and URL "about:blank".
'     ボタン側で try-catch して画面に「about:blank origin では使えない」旨を
'     表示するようにした (案 R')。9.10 以降で http(s) origin に移行する際に
'     pushState/replaceState の本格テストを行う。
'
'   ボタン構成:
'     [pushState]    : try-catch 入り (SecurityError 表示)
'     [replaceState] : 同上
'     [back]         : history.back()  ※履歴に push されていないので無反応
'     [forward]      : history.forward() ※同上
'     [notify route] : window.chrome.webview.postMessage(JSON) → WebMessageReceived
'                      (SPA 規約サンプル: {type:'spa-route', path:'/xxx'}、機能確認済み)
'
'   ★ SPA 規約サンプル (postMessage) について ★
'     VBA 側は既存の View_OnWebMessageReceived がそのまま受信する
'     (9.9b で新規 VBA コードは不要)。実 SPA に組み込むときは
'     「ルート変更時に JSON で type フィールドを持たせて postMessage する」
'     という規約を参考にできる。type による分岐パーサを VBA 側に作るかは
'     案件次第なので、ここでは規約サンプルを示すのみ (案 a)。
' ============================================================
Public Function BuildSpaTestHtml() As String
    Dim s As String
    s = "<!DOCTYPE html>" & vbCrLf
    s = s & "<html><head>" & vbCrLf
    s = s & "<meta charset='utf-8'>" & vbCrLf
    s = s & "<title>WebView2 SPA テスト</title>" & vbCrLf
    s = s & "<style>" & vbCrLf
    s = s & "  body { font-family: 'Segoe UI', sans-serif; padding: 20px; }" & vbCrLf
    s = s & "  h1 { color: #0078D4; }" & vbCrLf
    s = s & "  button { padding: 6px 14px; margin: 4px; }" & vbCrLf
    s = s & "  #url { font-weight: bold; color: #107C10; }" & vbCrLf
    s = s & "  #log { background: #f3f3f3; padding: 10px; min-height: 80px; " & _
                    "white-space: pre-wrap; border: 1px solid #ccc; margin-top: 10px; }" & vbCrLf
    s = s & "</style>" & vbCrLf
    s = s & "</head><body>" & vbCrLf
    s = s & "<h1>SPA イベントテスト</h1>" & vbCrLf
    s = s & "<div>現在のパス: <span id='url'>/</span></div>" & vbCrLf
    s = s & "<div>" & vbCrLf
    s = s & "  <button onclick='doPush()'>pushState</button>" & vbCrLf
    s = s & "  <button onclick='doReplace()'>replaceState</button>" & vbCrLf
    s = s & "  <button onclick='history.back()'>back</button>" & vbCrLf
    s = s & "  <button onclick='history.forward()'>forward</button>" & vbCrLf
    s = s & "  <button onclick='notifyRoute()'>notify route (postMessage)</button>" & vbCrLf
    s = s & "</div>" & vbCrLf
    s = s & "<div id='log'></div>" & vbCrLf
    s = s & "<script>" & vbCrLf
    s = s & "  let n = 0;" & vbCrLf
    s = s & "  function logMsg(t) {" & vbCrLf
    s = s & "    document.getElementById('log').innerText += t + '\n';" & vbCrLf
    s = s & "  }" & vbCrLf
    s = s & "  function showUrl() {" & vbCrLf
    s = s & "    document.getElementById('url').innerText = location.pathname;" & vbCrLf
    s = s & "  }" & vbCrLf
    s = s & "  function doPush() {" & vbCrLf
    s = s & "    // 第9.9b 検証で判明: NavigateToString は about:blank origin で表示されるため、" & vbCrLf
    s = s & "    // pushState は SecurityError で弾かれる (Web 標準仕様)。" & vbCrLf
    s = s & "    // 9.9b の検証目的 (HistoryChanged ハンドラ基盤) は NavigateToString 直後の" & vbCrLf
    s = s & "    // ランタイム自動発火で達成済み。本ボタンは UI 上のエラー表示確認用。" & vbCrLf
    s = s & "    try {" & vbCrLf
    s = s & "      n++;" & vbCrLf
    s = s & "      history.pushState({i:n}, '', '/page' + n);" & vbCrLf
    s = s & "      showUrl(); logMsg('pushState -> ' + location.pathname);" & vbCrLf
    s = s & "    } catch (e) {" & vbCrLf
    s = s & "      logMsg('pushState ERROR: ' + e.name + ' (' + e.message + ')');" & vbCrLf
    s = s & "      logMsg('  ※ about:blank origin では pushState 禁止 (Web 標準)。9.10 で http(s) origin に移行予定。');" & vbCrLf
    s = s & "    }" & vbCrLf
    s = s & "  }" & vbCrLf
    s = s & "  function doReplace() {" & vbCrLf
    s = s & "    // doPush と同じ理由で about:blank origin では SecurityError" & vbCrLf
    s = s & "    try {" & vbCrLf
    s = s & "      history.replaceState({}, '', '/replaced');" & vbCrLf
    s = s & "      showUrl(); logMsg('replaceState -> ' + location.pathname);" & vbCrLf
    s = s & "    } catch (e) {" & vbCrLf
    s = s & "      logMsg('replaceState ERROR: ' + e.name + ' (' + e.message + ')');" & vbCrLf
    s = s & "      logMsg('  ※ about:blank origin では replaceState 禁止 (Web 標準)。9.10 で http(s) origin に移行予定。');" & vbCrLf
    s = s & "    }" & vbCrLf
    s = s & "  }" & vbCrLf
    s = s & "  window.addEventListener('popstate', () => {" & vbCrLf
    s = s & "    showUrl(); logMsg('popstate -> ' + location.pathname);" & vbCrLf
    s = s & "  });" & vbCrLf
    s = s & "  function notifyRoute() {" & vbCrLf
    s = s & "    const payload = JSON.stringify({type:'spa-route', path: location.pathname});" & vbCrLf
    s = s & "    window.chrome.webview.postMessage(payload);" & vbCrLf
    s = s & "    logMsg('postMessage -> ' + payload);" & vbCrLf
    s = s & "  }" & vbCrLf
    s = s & "  showUrl();" & vbCrLf
    s = s & "  logMsg('SPA test page loaded (about:blank origin)');" & vbCrLf
    s = s & "</script>" & vbCrLf
    s = s & "</body></html>"

    BuildSpaTestHtml = s
End Function


' ============================================================
' Test_SpaTest_Real (第9.9b 段階で新規、SPA 対応の実機検証)
'
'   SPA テストページを表示し、HistoryChanged / DOMContentLoaded の
'   イベント発火を確認する。Test_AllEvents_Real と同形だが、
'   StartWebView2 ではなく StartWebView2_SpaTest を呼ぶ点が異なる。
'
'   検証手順 (実機検証で判明した実際の挙動に基づく):
'     1. 本プロシージャを実行 → SPA テストページが表示される
'     2. ページ表示直後、以下のイベントが順に発火するはず:
'        - NavigationStarting (data:text/html;... の巨大 URL)
'        - HistoryChanged x 2 (NavigateToString 完了時にランタイムが自動発火)
'                              → URL=about:blank、CanGoBack/Forward=False
'        - DocumentTitleChanged → タイトル="WebView2 SPA テスト"
'        - DOMContentLoaded → navigationId=N (UINT64)
'        - NavigationCompleted → isSuccess=True
'     3. [notify route] ボタン → WebMessageReceived 発火
'        → message: {"type":"spa-route","path":"blank"}
'        (SPA 規約サンプルの動作確認、画面にも postMessage の行が追記される)
'     4. [pushState] / [replaceState] ボタン
'        → about:blank origin の制約で SecurityError、画面にエラー文言を表示
'        (try-catch で捕捉、第9.9b 案 R' の対処)
'     5. [back] / [forward] ボタン → 履歴に push されていないので無反応
'     6. フォームを閉じてクリーンアップ
'        → HistoryChanged (kind=9) / DOMContentLoaded (kind=10) も
'           CleanupAllPersistentHandlers で一括解除されることを確認
'        → 特に kind=10 は RemoveInterfaceFor が m_pView2 を返して
'           remove_DOMContentLoaded (vtable 65) を撃つことを確認 (設計原則 45)
'
'   検証目標 (実機で達成済み):
'     ? HistoryChanged ハンドラの登録経路が機能 (自動発火 2 回で確認)
'     ? DOMContentLoaded が初回ロードで発火、navigationId が UINT64 で取れる
'     ? SPA 規約サンプル (postMessage JSON) が WebMessageReceived で受信できる
'     ? 7 種類目・8 種類目のハンドラ追加でもクリーンアップが一括で動く
'     ? DOMContentLoaded ハンドラ登録で View2 (m_pView2) が初投入される
'     ? RemoveInterfaceFor が DOMContentLoaded には m_pView2 を返す (設計原則 45 厳守)
'
'   将来の課題 (第9.10 以降):
'     - about:blank origin の制約により pushState/replaceState の本格テストは
'       未達。9.10 で http(s) origin (実 SPA サーバ or file:// + 工夫) に
'       移行する際に再検証する
' ============================================================
Public Sub Test_SpaTest_Real()

    Debug.Print String(60, "=")
    Debug.Print "Test_SpaTest_Real 開始 (SPA 対応検証)"
    Debug.Print String(60, "=")

    Debug.Print "UserForm1 を Show vbModeless で表示します ..."
    UserForm1.Show vbModeless

    Debug.Print "UserForm1.StartWebView2_SpaTest を呼び出します ..."
    UserForm1.StartWebView2_SpaTest

    Debug.Print "Test_SpaTest_Real のメイン処理は完了。"
    Debug.Print "[期待] ページ表示直後に HistoryChanged が 2 回自動発火 (NavigateToString)"
    Debug.Print "[期待] DOMContentLoaded が 1 回発火 (navigationId 取得)"
    Debug.Print "[手順] [notify route] ボタンで WebMessageReceived (SPA 規約) を確認可能"
    Debug.Print "[既知] [pushState] [replaceState] は about:blank origin の制約により"
    Debug.Print "       SecurityError、画面にエラー文言を表示 (try-catch で捕捉)"
    Debug.Print "[既知] [back] [forward] は履歴空のため無反応"
    Debug.Print "閉じるボタンでクリーンアップが走ります (kind=10 は m_pView2 で remove)。"
    Debug.Print String(60, "=")
End Sub


' ============================================================
' Test_ExecuteScript_Parallel (第9.9c 段階で新規、補完群 項目 D)
'
'   ExecuteScript の同時並行性 (案 a、9.8c で確立) が実機で正しく動作することを
'   検証する。5 個の ExecuteScript を連続発行し、5 個の callback ID が同時に
'   進行することを確認する。
'
'   設計判断 (たーぼーさん合意済み):
'     ・D1: 5 個並行 (採番が綺麗、ログ過剰でない)
'     ・D2: 連番計算 (callback ID と結果値の対応で混線がないことを確認、案 ii)
'
'   実行する JS と期待される結果:
'     ExecuteScript("1 + 1")  → resultJson = "2"
'     ExecuteScript("2 + 1")  → resultJson = "3"
'     ExecuteScript("3 + 1")  → resultJson = "4"
'     ExecuteScript("4 + 1")  → resultJson = "5"
'     ExecuteScript("5 + 1")  → resultJson = "6"
'
'   View_ExecuteScript の連番採番により callback ID は順に N, N+1, ..., N+4。
'   完了通知は WebView2 ランタイムから順次飛んでくるが、callback ID と
'   結果値が正しくペアリングされていれば「混線していない」と判定できる
'   (例: callback ID = N → resultJson = "2" のペアが崩れないこと)。
'
'   検証手順:
'     1. 本プロシージャを実行 → UserForm + Bing が表示される
'     2. 自動的に 5 個の ExecuteScript が連続発行される
'     3. Debug.Print で 5 個の callback ID 採番ログを確認
'     4. WebView2 が順次完了通知を飛ばすので、5 個の OnExecuteScriptCompleted
'        ログを確認 (callback ID と resultJson の対応が正しいこと)
'     5. UserForm1.OnExecuteScriptResult も 5 回呼ばれるはず (9.9c E)
'     6. フォームを閉じる
'
'   注意:
'     ExecuteScript はページ読み込み完了前に撃つと正しく実行されないことが
'     ある。本テストでは Bing が完全に表示されるまで少し待つ。
'     Test_AllEvents_Real と同じく vbModeless でフォームを出すので、
'     Bing 表示完了は IsPageLoaded で待つ。
' ============================================================
Public Sub Test_ExecuteScript_Parallel()

    Debug.Print String(60, "=")
    Debug.Print "Test_ExecuteScript_Parallel 開始 (補完群 項目 D)"
    Debug.Print String(60, "=")

    Debug.Print "UserForm1 を Show vbModeless で表示します ..."
    UserForm1.Show vbModeless

    Debug.Print "UserForm1.StartWebView2 を呼び出します ..."
    UserForm1.StartWebView2

    Debug.Print "ページの読み込み完了を待ちます (上限 15 秒) ..."
    Dim t As Single: t = Timer
    Do While (Not UserForm1.IsPageLoaded) And ((Timer - t) < 15#)
        DoEvents
    Loop

    If Not UserForm1.IsPageLoaded Then
        Debug.Print "[NG] ページ読み込みタイムアウト (Bing にアクセスできず?)"
        Debug.Print "閉じるボタンでクリーンアップしてください。"
        Debug.Print String(60, "=")
        Exit Sub
    End If

    Debug.Print "[OK] ページ読み込み完了。5 個の ExecuteScript を連続発行します ..."
    Debug.Print ""

    ' --- 5 個の ExecuteScript を連続発行 ---
    '   各呼び出しは callback ID を即座に返す (非同期)。
    '   完了通知は WebView2 ランタイムから後ほど飛んでくる。
    Dim pane As Wv2Pane
    Set pane = UserForm1.GetActivePane
    If pane Is Nothing Then
        Debug.Print "[NG] GetActivePane が Nothing"
        Exit Sub
    End If

    Dim ids(1 To 5) As Long
    Dim i As Long
    For i = 1 To 5
        ids(i) = pane.View_ExecuteScript(CStr(i) & " + 1")
        Debug.Print "[発行 #" & i & "] callback ID=" & ids(i) & _
                    " (期待結果=" & (i + 1) & ")"
    Next i

    Debug.Print ""
    Debug.Print "[手順] 5 個の OnExecuteScriptCompleted ログを順に確認:"
    Debug.Print "       callback ID と resultJson の対応が正しいことを目視確認 (混線なし)"
    Debug.Print "       期待ペア: (" & ids(1) & ", ""2"") (" & ids(2) & ", ""3"") " & _
                "(" & ids(3) & ", ""4"") (" & ids(4) & ", ""5"") (" & ids(5) & ", ""6"")"
    Debug.Print "[手順] UserForm1.OnExecuteScriptResult も 5 回ログされる (9.9c E)"
    Debug.Print "閉じるボタンでクリーンアップが走ります。"
    Debug.Print String(60, "=")
End Sub


' ============================================================
' Test_ResetSurvival_Manual (第9.4a 段階で新規)
'
'   リセットボタン耐性 (案 R4) の実機検証ガイド。
'   このプロシージャ自体は手順をログ出力するだけで、実際の検証は
'   ユーザが VBE で手作業で行う。
'
'   検証手順 (ユーザが実機で行う):
'     1. このプロシージャ Test_ResetSurvival_Manual を実行 (案内表示)
'     2. Test_Navigate_Real を実行して Bing 表示 + NavigationStarting 登録
'     3. Bing が表示されたら、ページ内のリンクを 1?2 回クリックして
'        NavigationStarting イベントが正しく発火することを Debug.Print で確認
'     4. **その状態で VBE の停止ボタン (リセット) を押す**
'        → m_pRegionBase などのモジュール変数が初期値 0 にリセットされる
'        → ただし VirtualAlloc 領域は OS が回収しないので残る
'        → 生存フラグも残ったまま (Thunks_Shutdown が呼ばれていない)
'     5. Excel が落ちないことを確認
'     6. しばらく待って (もしくは別のリンクをクリックしようとして) 、
'        WebView2 が ComCallbackHandler サンクを叩いても Excel がクラッシュしないことを確認
'        (※生存フラグの仕組みでは、リセット後でも領域は残っているので
'         サンクは即 return せず本来の処理を試みる可能性あり。これが
'         クラッシュにつながるかどうかが本検証の目的)
'     7. 同じセッションで Test_Navigate_Real を再実行できることを確認
'        (Thunks_Init が「m_pRegionBase = 0 なら新規確保」なので、
'         過去のリーク領域とは別に新規確保される)
'
'   期待される結果:
'     ? リセット後も Excel がクラッシュしない
'     ? 同セッション内で再実行が可能
'     ? 第9.5 段階以降は、再実行時にセンチネル機構が旧領域を
'       自動的に VirtualFree する (リーク完全解消)
'
'   センチネル機構の動作観察には Test_SentinelStatus を併用すること。
' ============================================================
Public Sub Test_ResetSurvival_Manual()
    Debug.Print String(60, "=")
    Debug.Print "Test_ResetSurvival_Manual (リセット試験ガイド)"
    Debug.Print String(60, "=")
    Debug.Print "リセットボタン耐性の検証手順:"
    Debug.Print ""
    Debug.Print "1. 別途 Test_Navigate_Real を実行して Bing を表示"
    Debug.Print "2. Bing のリンクを 1?2 回クリックして NavigationStarting イベントが"
    Debug.Print "   Debug.Print されることを確認"
    Debug.Print "3. その状態で VBE の停止ボタン (リセット) を押す"
    Debug.Print "4. Excel が落ちないことを確認"
    Debug.Print "5. しばらく待って、リンクをクリックしようとしても"
    Debug.Print "   Excel がクラッシュしないことを確認"
    Debug.Print "6. 同セッションで Test_Navigate_Real を再実行できることを確認"
    Debug.Print ""
    Debug.Print "期待結果:"
    Debug.Print "  ? リセット後も Excel がクラッシュしない (生存フラグの効果)"
    Debug.Print "  ? 同セッション内で再実行が可能"
    Debug.Print "  ? 第9.5 段階以降は、再実行時にセンチネル機構が"
    Debug.Print "    旧領域を自動的に VirtualFree する (リーク完全解消)"
    Debug.Print ""
    Debug.Print "センチネル機構の動作確認には Test_SentinelStatus を併用:"
    Debug.Print "  - リセット直後に呼ぶと環境変数に旧領域アドレスが残っているはず"
    Debug.Print "  - 再 Init 後に呼ぶと新領域アドレスに置き換わっているはず"
    Debug.Print String(60, "=")
End Sub


' ============================================================
' Sentinel_LoadPrevRegion (第9.5 段階で追加)
'
'   環境変数 SENTINEL_ENV_NAME から「前回 Thunks_Init で確保した領域の
'   ベースアドレス」を読み出す。痕跡が無い (= 通常終了済み、または
'   初回起動) 場合は 0 を返す。
'
'   ※ プロセス起動直後にこの環境変数が存在することは通常無い。
'      存在するのは「直前にリセットで Shutdown が走らなかった」場合のみ。
'      (環境変数はプロセス固有なので、別 Excel プロセスの値が混入する
'       ことは無い)
'
'   戻り値: 前回領域のベースアドレス (なければ 0)
' ============================================================
Private Function Sentinel_LoadPrevRegion() As LongPtr
    Sentinel_LoadPrevRegion = 0^

    Dim buff As String
    buff = String$(SENTINEL_BUFFER_SIZE, vbNullChar)

    Dim n As Long
    n = GetEnvironmentVariableW(StrPtr(SENTINEL_ENV_NAME), StrPtr(buff), SENTINEL_BUFFER_SIZE)

    ' n = 0 : 変数が存在しない or 取得失敗
    If n = 0 Then Exit Function

    ' n >= SENTINEL_BUFFER_SIZE : バッファ不足 (= 異常値、無視)
    '   GetEnvironmentVariableW は成功時 nSize 未満の値、不足時は必要な
    '   バッファサイズを返す。アドレス文字列が 19 桁を超えることは無いので
    '   この経路に来た場合は環境変数が破損していると判断。
    If n >= SENTINEL_BUFFER_SIZE Then Exit Function

    Dim s As String
    s = Left$(buff, n)

    ' 数値変換に失敗 (空文字、数値以外の混入など) した場合は 0 を返す
    On Error Resume Next
    Sentinel_LoadPrevRegion = CLngLng(s)
    On Error GoTo 0
End Function


' ============================================================
' Sentinel_StorePrevRegion (第9.5 段階で追加)
'
'   現在の領域ベースアドレスを環境変数に書き込む。
'   形式は十進文字列 (MouseScroll.bas 踏襲)。
'
'   呼び出しタイミング:
'     Thunks_Init の VirtualAlloc 成功直後。これにより、Init の途中で
'     リセットされても次回 Init で確実に領域が回収される。
' ============================================================
Private Sub Sentinel_StorePrevRegion(ByVal addr As LongPtr)
    Dim s As String
    s = CStr(addr)
    SetEnvironmentVariableW StrPtr(SENTINEL_ENV_NAME), StrPtr(s)
End Sub


' ============================================================
' Sentinel_ClearPrevRegion (第9.5 段階で追加)
'
'   環境変数をクリアする。
'   Win32 仕様: SetEnvironmentVariableW(lpName, NULL) は変数を削除する。
'   VBA からは第二引数を ByVal 0 で呼べばよい (LongPtr の ByVal で 0 を
'   渡せばそのまま NULL ポインタとして扱われる)。
'
'   呼び出しタイミング:
'     1. Thunks_Shutdown の末尾 (通常終了パス)
'     2. Sentinel_RecoverIfNeeded の末尾 (リセット復帰後の新 Init 開始前)
' ============================================================
Private Sub Sentinel_ClearPrevRegion()
    SetEnvironmentVariableW StrPtr(SENTINEL_ENV_NAME), 0
End Sub


' ============================================================
' Sentinel_RecoverIfNeeded (第9.5 段階で追加)
'
'   Thunks_Init の冒頭で呼ばれる。前回の痕跡があれば旧領域を回収する。
'
'   回収手順 (案 β、生存フラグ + VirtualFree の二重防御):
'     1. 旧領域の生存フラグを 0 に倒す
'        - 万一 WebView2 がまだスタブを叩いてもサンク先頭で S_OK 帰投
'        - VirtualFree が完了するまでの一瞬の安全弁
'        - 領域が既に解放/再マップされている場合はアクセス違反の可能性が
'          あるため On Error で防御
'     2. VirtualFree で旧領域を解放
'        - 失敗時は GetLastError で原因切り分け
'     3. 環境変数をクリア
'        - この直後に Thunks_Init が新領域アドレスで上書きするが、万一
'          VirtualAlloc が失敗した場合に古い値が残らないよう先にクリア
'
'   回収対象が「同一プロセス内で放置された領域のみ」であることの保証:
'     - 環境変数はプロセス固有 (親→子へのコピーはあるが他プロセスから
'       読まれることはない)
'     - Excel が VBA リセットで再起動するわけではない (同一プロセス継続)
'     - したがって痕跡がある = 同一プロセス内の放置確定
'
'   診断ログ (将来の事故時の切り分けに役立つので残す):
'     - "VirtualQuery result ...": OS から見た領域状態
'         State = MEM_COMMIT (4096)  → 領域は健在 (正常)
'         State = MEM_FREE   (65536) → 既に解放済み
'         State = MEM_RESERVE(8192)  → 中間状態
'     - "header LongPtr before clear: X" : 旧領域先頭の生存フラグ
'         X = 1   → 領域は健在、フラグも立ったまま (正常)
'         X = 0   → 既に誰かがフラグを倒した
'         アクセス違反 → 領域は既に解放/再マップされている
'     - "VirtualFree ... LastError: Y" : VirtualFree の結果
'         Y = 0   → 通常はここに来る (succeeded)
'         Y = 87  → ERROR_INVALID_PARAMETER (第9.5 段階で真因解消済み、
'                  もし再発したら定数の & サフィックス忘れを疑え)
'         Y = 170 → ERROR_BUSY (他スレッドが領域内コードを実行中)
'         Y = 487 → ERROR_INVALID_ADDRESS (アドレス不正、既に解放済み等)
' ============================================================
Private Sub Sentinel_RecoverIfNeeded()
    Dim prevBase As LongPtr
    prevBase = Sentinel_LoadPrevRegion()
    If prevBase = 0 Then Exit Sub

    Debug.Print "Sentinel: detected previous region = " & prevBase & _
                " (decimal), recovering ..."

    ' --- 診断0 + 健在判定 (第9.5 段階追加 / 第9.10a で健在判定に格上げ) ---
    '   VirtualQuery で OS から見た領域状態を観察し、その結果で「領域が健在か」を
    '   判定する。これにより以下が切り分けられる:
    '     State = MEM_COMMIT (4096)  → 領域は健在 (VirtualFree の引数を疑え)
    '     State = MEM_FREE   (65536) → 既に解放済み (誰かが先に Free した)
    '     State = MEM_RESERVE(8192)  → Reserve のみ残存 (中間状態)
    '   AllocationBase が prevBase と一致するかも見る (= 領域の真の先頭か)。
    '
    '   ★ 第9.10a の最重要修正 (仕様事実 10 の対処) ★
    '     旧実装は「VirtualQuery で State を観察」しておきながら、その結果を
    '     分岐に使わず ReadLongPtr / MemLongPtr / VirtualFree を無条件に実行して
    '     いた。State=MEM_FREE の解放済みアドレスに ReadLongPtr / MemLongPtr を
    '     撃つと SAFEARRAY 経由のネイティブ配列アクセスで AV (アクセス違反) が
    '     発生する。この AV は CPU の SEH 例外であり VBA の On Error Resume Next
    '     では捕捉できない (= Excel プロセス即落ち、仕様事実 10)。
    '     よって VirtualQuery の結果で「健在」と確認できた場合に限り、
    '     ReadLongPtr / MemLongPtr / VirtualFree を実行する (設計原則 55)。
    Dim mbi As MEMORY_BASIC_INFORMATION
    Dim qSize As LongPtr
    Dim regionAlive As Boolean
    regionAlive = False

    qSize = VirtualQuery(prevBase, mbi, LenB(mbi))
    If qSize = 0 Then
        Debug.Print "Sentinel: VirtualQuery FAILED, LastError = " & GetLastError()
    Else
        Debug.Print "Sentinel: VirtualQuery result (" & qSize & " bytes copied):"
        Debug.Print "  BaseAddress       = " & mbi.BaseAddress
        Debug.Print "  AllocationBase    = " & mbi.AllocationBase & _
                    IIf(mbi.AllocationBase = prevBase, " (= prevBase, OK)", " (≠ prevBase, ズレあり)")
        Debug.Print "  AllocationProtect = &H" & Hex$(mbi.AllocationProtect) & _
                    " (" & FormatMemProtect(mbi.AllocationProtect) & ")"
        Debug.Print "  RegionSize        = " & mbi.RegionSize
        Debug.Print "  State             = " & mbi.state & " (" & FormatMemState(mbi.state) & ")"
        Debug.Print "  Protect           = &H" & Hex$(mbi.Protect) & _
                    " (" & FormatMemProtect(mbi.Protect) & ")"
        Debug.Print "  Type              = &H" & Hex$(mbi.Type_)

        ' 健在判定 (堅め): 3 条件を全て満たすときのみ「健在」とみなす。
        '   (1) State = MEM_COMMIT       : ページが commit 済み (解放/予約のみでない)
        '   (2) AllocationBase = prevBase : 領域の真の先頭であり別領域に再マップ
        '                                   されていない
        '   (3) Protect が読み書き可能    : PAGE_GUARD / PAGE_NOACCESS でない。
        '       サンク領域は PAGE_EXECUTE_READWRITE で確保されるが、念のため
        '       PAGE_READWRITE も許容する。PAGE_GUARD ビットが立っていると初回
        '       アクセスで STATUS_GUARD_PAGE_VIOLATION になるため除外する。
        regionAlive = (mbi.state = MEM_COMMIT) _
                      And (mbi.AllocationBase = prevBase) _
                      And IsProtectReadWritable(mbi.Protect)
    End If

    If Not regionAlive Then
        ' 領域は健在でない (= 既に解放済み / 再マップ / 予約のみ / 保護不可)。
        ' ReadLongPtr / MemLongPtr / VirtualFree は一切撃たず、環境変数クリアのみ
        ' 行って安全に抜ける。VirtualFree も撃たない理由: 解放済み領域への
        ' MEM_RELEASE は ERROR_INVALID_PARAMETER になるだけで実害はないが、
        ' 「健在でないなら触らない」を徹底することでパスを単純化する。
        Debug.Print "Sentinel: 旧領域は健在でないと判定 → 触れずにスキップ (仕様事実 10 対処)"
        Sentinel_ClearPrevRegion
        Exit Sub
    End If

    ' --- ここから先は領域が健在であることが VirtualQuery で確認済み ---

    ' 診断1: 領域先頭の生存フラグを読む (健在確認済みなので AV は起きない)
    '   注: MemLongPtr は Property Let のみ (書き込み専用)。読み出しには
    '       ReadLongPtr 関数を使う。
    Dim headerByte As LongPtr
    headerByte = ReadLongPtr(prevBase)
    Debug.Print "Sentinel: header LongPtr before clear = " & headerByte & _
                " (生存フラグ = " & (headerByte And &HFF&) & ")"

    ' 1. 生存フラグを 0 に倒す (二重防御)
    MemLongPtr(prevBase) = 0^

    ' 2. VirtualFree で旧領域を解放 (GetLastError でエラーコードも取得)
    '    Win32 仕様: VirtualFree は失敗時 0 を返し、原因は GetLastError で取得する。
    Dim freeResult As Long
    Dim lastErr As Long
    freeResult = VirtualFree(prevBase, 0, MEM_RELEASE)
    lastErr = GetLastError()
    If freeResult <> 0 Then
        Debug.Print "Sentinel: VirtualFree succeeded for " & prevBase
    Else
        Debug.Print "Sentinel: VirtualFree returned 0, LastError = " & lastErr & _
                    " (" & VirtualFreeErrorName(lastErr) & ")"
    End If

    ' 3. 環境変数を先にクリア (新領域アドレス書き込みの前)
    Sentinel_ClearPrevRegion
End Sub

' ============================================================
' IsProtectReadWritable (第9.10a 追加)
' ============================================================
'   MEMORY_BASIC_INFORMATION.Protect が「VBA から読み書きしても安全な保護属性か」
'   を判定する。Sentinel_RecoverIfNeeded の健在判定の一部 (設計原則 55)。
'
'   読み書き可能とみなす: PAGE_READWRITE (&H4) / PAGE_EXECUTE_READWRITE (&H40)。
'   ただし PAGE_GUARD (&H100) ビットが立っているとガードページ例外が起きるので、
'   そのときは読み書き不可とする。
Private Function IsProtectReadWritable(ByVal prot As Long) As Boolean
    ' PAGE_GUARD ビットが立っていたら不可
    If (prot And &H100&) <> 0 Then
        IsProtectReadWritable = False
        Exit Function
    End If
    ' ガードビットを除いた基本保護属性で判定
    Dim base As Long
    base = prot And Not &H100& And Not &H200& And Not &H400&
    Select Case base
        Case PAGE_READWRITE, PAGE_EXECUTE_READWRITE
            IsProtectReadWritable = True
        Case Else
            IsProtectReadWritable = False
    End Select
End Function


' ============================================================
' VirtualFreeErrorName (第9.5 段階で追加、診断用)
'
'   GetLastError の値を Win32 エラー名に変換する。代表的なものだけ
'   個別判定し、未知の値は "UNKNOWN" を返す。
' ============================================================
Private Function VirtualFreeErrorName(ByVal errCode As Long) As String
    Select Case errCode
        Case 0: VirtualFreeErrorName = "ERROR_SUCCESS (= 実は成功してる?)"
        Case 5: VirtualFreeErrorName = "ERROR_ACCESS_DENIED"
        Case 6: VirtualFreeErrorName = "ERROR_INVALID_HANDLE"
        Case 8: VirtualFreeErrorName = "ERROR_NOT_ENOUGH_MEMORY"
        Case 87: VirtualFreeErrorName = "ERROR_INVALID_PARAMETER"
        Case 170: VirtualFreeErrorName = "ERROR_BUSY (他スレッド使用中)"
        Case 487: VirtualFreeErrorName = "ERROR_INVALID_ADDRESS"
        Case Else: VirtualFreeErrorName = "UNKNOWN"
    End Select
End Function


' ============================================================
' FormatMemState (第9.5 段階で追加、診断用)
'
'   MEMORY_BASIC_INFORMATION.State の値を文字列化する。
' ============================================================
Private Function FormatMemState(ByVal state As Long) As String
    Select Case state
        Case &H1000: FormatMemState = "MEM_COMMIT (領域は健在)"
        Case &H2000: FormatMemState = "MEM_RESERVE (Reserve のみ、Commit 未)"
        Case &H10000: FormatMemState = "MEM_FREE (既に解放済み)"
        Case Else: FormatMemState = "UNKNOWN(&H" & Hex$(state) & ")"
    End Select
End Function


' ============================================================
' FormatMemProtect (第9.5 段階で追加、診断用)
'
'   MEMORY_BASIC_INFORMATION.Protect / AllocationProtect の値を文字列化する。
'   PAGE_GUARD 等のフラグはビット OR されることがあるので個別に検出する。
' ============================================================
Private Function FormatMemProtect(ByVal prot As Long) As String
    If prot = 0 Then
        FormatMemProtect = "NONE (= 解放済み領域では Protect=0)"
        Exit Function
    End If

    Dim s As String

    Select Case prot And &HFF&
        Case &H1: s = "PAGE_NOACCESS"
        Case &H2: s = "PAGE_READONLY"
        Case &H4: s = "PAGE_READWRITE"
        Case &H8: s = "PAGE_WRITECOPY"
        Case &H10: s = "PAGE_EXECUTE"
        Case &H20: s = "PAGE_EXECUTE_READ"
        Case &H40: s = "PAGE_EXECUTE_READWRITE"
        Case &H80: s = "PAGE_EXECUTE_WRITECOPY"
        Case Else: s = "UNKNOWN(&H" & Hex$(prot And &HFF&) & ")"
    End Select

    If (prot And &H100&) <> 0 Then s = s & "|PAGE_GUARD"
    If (prot And &H200&) <> 0 Then s = s & "|PAGE_NOCACHE"
    If (prot And &H400&) <> 0 Then s = s & "|PAGE_WRITECOMBINE"

    FormatMemProtect = s
End Function


' ============================================================
' Test_VirtualAllocFree_Roundtrip (第9.5 段階で追加)
'
'   VBA リセットを経由しない通常フローで VirtualAlloc → VirtualFree が
'   成功するかを 5 種類のシナリオで観察する。
'   9.5c で判明した「リセット経由領域だけ ERROR_INVALID_PARAMETER で
'   失敗する」現象の真因を切り分けるのが目的。
'
'   使い方:
'     1. Excel + VBE を再起動して綺麗な状態にする (任意)
'     2. イミディエイトで Test_VirtualAllocFree_Roundtrip を実行
'     3. α?δ の結果を観察
'     4. (オプション) その後 Test_AllEvents_Real などで Thunks_Init を走らせ、
'        m_pRegionBase が確保された状態で Test_VirtualAllocFree_Eps を実行する
'        ことでシナリオ ε も観察できる
'
'   各シナリオが全成功 → リセット経由領域だけが特異的に失敗する
'                       (VBA ランタイムが何らかのマークを付けている説が有力)
'   どれかが失敗   → そのシナリオに含まれる要素が真因
' ============================================================
Public Sub Test_VirtualAllocFree_Roundtrip()
    Debug.Print String(60, "=")
    Debug.Print "Test_VirtualAllocFree_Roundtrip (リセット非経由での確認)"
    Debug.Print String(60, "=")

    Dim p As LongPtr, p2 As LongPtr
    Dim r As Long
    Dim errCode As Long

    ' --- シナリオ α: alloc → 即 free ---
    Debug.Print "--- シナリオ α: alloc → 即 free ---"
    p = VirtualAlloc(0, REGION_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    Debug.Print "  VirtualAlloc -> " & p
    If p = 0 Then
        Debug.Print "  [SKIP] alloc が失敗、以降のシナリオもスキップ"
        Exit Sub
    End If
    r = VirtualFree(p, 0, MEM_RELEASE)
    errCode = GetLastError()
    Debug.Print "  VirtualFree returned " & r & ", LastError = " & errCode & _
                IIf(r <> 0, " [OK]", " [NG] (" & VirtualFreeErrorName(errCode) & ")")

    ' --- シナリオ β: alloc → 中身書き込み → free ---
    Debug.Print "--- シナリオ β: alloc → 中身書き込み → free ---"
    p = VirtualAlloc(0, REGION_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    Debug.Print "  VirtualAlloc -> " & p
    If p <> 0 Then
        MemLongPtr(p) = 1^                       ' 生存フラグ風の書き込み
        MemLongPtr(p + 8) = &H42^                ' 適当な値 (Long 範囲内で安全)
        Debug.Print "  生存フラグ書き込み済み (header LongPtr = " & ReadLongPtr(p) & ")"
        r = VirtualFree(p, 0, MEM_RELEASE)
        errCode = GetLastError()
        Debug.Print "  VirtualFree returned " & r & ", LastError = " & errCode & _
                    IIf(r <> 0, " [OK]", " [NG] (" & VirtualFreeErrorName(errCode) & ")")
    End If

    ' --- シナリオ γ: alloc → 十進文字列で往復変換 → free ---
    Debug.Print "--- シナリオ γ: alloc → アドレス十進往復変換 → free ---"
    p = VirtualAlloc(0, REGION_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    Debug.Print "  VirtualAlloc -> " & p
    If p <> 0 Then
        Dim s As String
        s = CStr(p)
        Dim pRestored As LongPtr
        pRestored = CLngLng(s)
        Debug.Print "  CStr/CLngLng round-trip: " & p & " -> '" & s & "' -> " & pRestored & _
                    IIf(p = pRestored, " [一致]", " [不一致!]")
        r = VirtualFree(pRestored, 0, MEM_RELEASE)
        errCode = GetLastError()
        Debug.Print "  VirtualFree(pRestored) returned " & r & ", LastError = " & errCode & _
                    IIf(r <> 0, " [OK]", " [NG] (" & VirtualFreeErrorName(errCode) & ")")
    End If

    ' --- シナリオ δ: alloc(A) → alloc(B) → free(A) → free(B) ---
    '   現在の Sentinel_RecoverIfNeeded と最も近い状況。ただしリセットは
    '   挟まない。
    Debug.Print "--- シナリオ δ: 2 領域確保 → 古い方を先に free → 新しい方を free ---"
    p = VirtualAlloc(0, REGION_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    Debug.Print "  領域 A alloc -> " & p
    p2 = VirtualAlloc(0, REGION_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    Debug.Print "  領域 B alloc -> " & p2
    If p <> 0 And p2 <> 0 Then
        MemLongPtr(p) = 1^                     ' 領域 A に生存フラグ風書き込み
        r = VirtualFree(p, 0, MEM_RELEASE)
        errCode = GetLastError()
        Debug.Print "  VirtualFree(A) returned " & r & ", LastError = " & errCode & _
                    IIf(r <> 0, " [OK]", " [NG] (" & VirtualFreeErrorName(errCode) & ")")
        r = VirtualFree(p2, 0, MEM_RELEASE)
        errCode = GetLastError()
        Debug.Print "  VirtualFree(B) returned " & r & ", LastError = " & errCode & _
                    IIf(r <> 0, " [OK]", " [NG] (" & VirtualFreeErrorName(errCode) & ")")
    End If

    ' --- シナリオ ε: 本物の Thunks_Init 領域に対する VirtualFree ---
    '   既に m_pRegionBase が確保されている (= Thunks_Init 済み) なら、
    '   そのアドレスに対して VirtualQuery → VirtualFree を試みる。
    '   注意: これを実行すると Thunks_Shutdown を経由しないので
    '         m_pRegionBase は 0 にならない (= 二重解放の危険)。
    '         本シナリオは Thunks_Shutdown 後の状態でのみ意味があるので、
    '         m_pRegionBase = 0 のときだけ実行する。
    Debug.Print "--- シナリオ ε: 本物の Thunks_Init を経た領域 ---"
    If m_pRegionBase <> 0 Then
        Debug.Print "  [SKIP] m_pRegionBase = " & m_pRegionBase & " (= 既に Init 済み)"
        Debug.Print "         このシナリオは Shutdown 後 (= m_pRegionBase = 0) で再実行してください"
    Else
        Debug.Print "  Thunks_Init を実行 ..."
        If Not Thunks_Init() Then
            Debug.Print "  [SKIP] Thunks_Init 失敗"
        Else
            Dim baseBak As LongPtr
            baseBak = m_pRegionBase
            Debug.Print "  m_pRegionBase = " & baseBak
            ' Thunks_Shutdown を呼ばずに直接 VirtualFree を試みる。
            ' その後の状態整合のため、m_pRegionBase などのモジュール変数は手動で
            ' クリアして Sentinel_ClearPrevRegion も呼ぶ。
            r = VirtualFree(baseBak, 0, MEM_RELEASE)
            errCode = GetLastError()
            Debug.Print "  VirtualFree(本物) returned " & r & ", LastError = " & errCode & _
                        IIf(r <> 0, " [OK]", " [NG] (" & VirtualFreeErrorName(errCode) & ")")
            ' 後始末
            m_pRegionBase = 0
            m_freeHead = -1
            m_inUse = 0
            Erase m_freeNext
            Sentinel_ClearPrevRegion
            Debug.Print "  後始末完了 (m_pRegionBase = 0、環境変数クリア済み)"
        End If
    End If

    Debug.Print String(60, "=")
    Debug.Print "Test_VirtualAllocFree_Roundtrip 完了"
    Debug.Print String(60, "=")
End Sub


' ============================================================
' Test_SentinelStatus (第9.5 段階で追加、デバッグ用)
'
'   現在の環境変数の中身と m_pRegionBase を Debug.Print するだけのテスト関数。
'   イミディエイトウィンドウから「Test_SentinelStatus」で実行できる。
'
'   観察ポイント:
'     - 初回起動直後: 環境変数なし、m_pRegionBase = 0
'     - StartWebView2 後: 環境変数 = m_pRegionBase = 新領域アドレス
'     - 通常終了後: 環境変数なし、m_pRegionBase = 0
'     - リセット直後: 環境変数 = 旧領域アドレス、m_pRegionBase = 0
'         (VBA 変数はリセットで消えるが環境変数はプロセス内で生存)
'     - リセット後の再 Init 後: 環境変数 = m_pRegionBase = 新領域アドレス
'         (旧領域は既に VirtualFree 済み)
' ============================================================
Public Sub Test_SentinelStatus()
    Dim prev As LongPtr
    prev = Sentinel_LoadPrevRegion()

    Debug.Print String(60, "=")
    Debug.Print "Test_SentinelStatus"
    Debug.Print "  Env var name : " & SENTINEL_ENV_NAME
    If prev = 0 Then
        Debug.Print "  Env var value: (empty or unset) → no leftover region"
    Else
        Debug.Print "  Env var value: " & prev & " (decimal pointer)"
        If prev = m_pRegionBase Then
            Debug.Print "  → 現在使用中の領域と一致 (正常状態)"
        ElseIf m_pRegionBase = 0 Then
            Debug.Print "  → 次回 Thunks_Init で回収される予定 (リセット後)"
        Else
            Debug.Print "  → 環境変数と m_pRegionBase が食い違っている (異常)"
        End If
    End If
    Debug.Print "  m_pRegionBase: " & m_pRegionBase
    Debug.Print String(60, "=")
End Sub


' ============================================================
' Thunks_Init
'   領域を VirtualAlloc で確保し、ヘッダの生存フラグを立て、
'   全スロットにスタブをコピーし、vtableObj を初期化し、
'   フリーリストを初期化する。
'
'   第9.2段階での処理 (第9.3a でも変更なし):
'     - Handler_QI / AddRef / Release のアドレスを 1 回だけ取得
'     - 全 512 スロットの vtableObj を一括初期化
'         pVTable      = pFunctions
'         Functions(0) = m_pHandler_QI
'         Functions(1) = m_pHandler_AddRef
'         Functions(2) = m_pHandler_Release
'         Functions(3) = pSlot (= スタブクローン先頭)
'
'   第9.5 段階での追加処理:
'     - 冒頭で Sentinel_RecoverIfNeeded を呼ぶ
'         (前回リセット時に放置された領域があれば VirtualFree して回収)
'     - VirtualAlloc 成功直後に Sentinel_StorePrevRegion を呼ぶ
'         (新領域アドレスを環境変数に記録、次回リセット復帰時に使う)
'
'   戻り値: 成功 = True、失敗 = False
' ============================================================
Public Function Thunks_Init() As Boolean

    If m_pRegionBase <> 0 Then
        Thunks_Init = True
        Exit Function
    End If

    ' --- センチネル機構 (第9.5 段階追加) ---
    '   前回プロセス内でリセットが押されて Shutdown が走らなかった場合、
    '   環境変数に旧領域アドレスが残っている。これを回収してから
    '   新領域確保に進む。
    Sentinel_RecoverIfNeeded

    ' Force-Compile EntryPoint (スタブのソースを実体化)
    EntryPoint

    ' EntryPoint スタブのソースアドレス
    Dim pStubSrc As LongPtr
    pStubSrc = VBA.Int(AddressOf EntryPoint)
    If pStubSrc = 0 Then Exit Function

    ' Handler_QI / AddRef / Release のアドレスを起動時に 1 回だけ取得して保存
    m_pHandler_QI = GetAddr(AddressOf Handler_QueryInterface)
    m_pHandler_AddRef = GetAddr(AddressOf Handler_AddRef)
    m_pHandler_Release = GetAddr(AddressOf Handler_Release)
    If m_pHandler_QI = 0 Or m_pHandler_AddRef = 0 Or m_pHandler_Release = 0 Then
        Exit Function
    End If

    ' 領域を一括確保
    m_pRegionBase = VirtualAlloc(0, REGION_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If m_pRegionBase = 0 Then Exit Function

    ' --- センチネル機構 (第9.5 段階追加) ---
    '   今確保したアドレスを環境変数に記録。
    '   これにより、この後 Init が完了する前にリセットが押されても
    '   次回 Init で確実に回収できる。
    Sentinel_StorePrevRegion m_pRegionBase

    ' ヘッダのゼロ初期化 (VirtualAlloc は本来ゼロクリア済みだが念のため)
    Dim k As Long
    For k = 0 To HEADER_SIZE - 1 Step 8
        MemLongPtr(m_pRegionBase + k) = 0^
    Next k

    ' 生存フラグを立てる (領域先頭の 1 byte = 1)
    MemLongPtr(m_pRegionBase) = 1^

    ' 全スロット初期化 (スタブコピー + vtableObj 初期化)
    Dim i As Long, pSlot As LongPtr, pVTableObj As LongPtr, pFunctions As LongPtr
    For i = 0 To SLOT_COUNT - 1
        pSlot = SlotAddrAt(i)

        ' --- スタブコピー (各 96 bytes、+91..+95 のパディングは無害) ---
        For k = 0 To 95 Step 8
            MemLongPtr(pSlot + k) = ReadLongPtr(pStubSrc + k)
        Next k

        ' --- vtableObj 初期化 (40 bytes、各 8 byte 単位で書き込み) ---
        pVTableObj = pSlot + VTABLE_OBJ_OFFSET
        pFunctions = pVTableObj + PtrSize        ' +8、pVTable の直後

        MemLongPtr(pVTableObj) = pFunctions
        MemLongPtr(pFunctions + 0 * PtrSize) = m_pHandler_QI
        MemLongPtr(pFunctions + 1 * PtrSize) = m_pHandler_AddRef
        MemLongPtr(pFunctions + 2 * PtrSize) = m_pHandler_Release
        MemLongPtr(pFunctions + 3 * PtrSize) = pSlot
    Next i

    ' フリーリスト初期化
    ReDim m_freeNext(0 To SLOT_COUNT - 1)
    For i = 0 To SLOT_COUNT - 2
        m_freeNext(i) = i + 1
    Next i
    m_freeNext(SLOT_COUNT - 1) = -1
    m_freeHead = 0
    m_inUse = 0

    ' Handler 対応表を全クリア
    For i = 0 To SLOT_COUNT - 1
        Set m_handlers(i) = Nothing
    Next i

    ' --- IID テーブル初期化 (第9.7a 段階で追加) ---
    '   IID_IUnknown と HandlerKind ごとの IID をテーブルに埋め込む。
    InitIIDTable

    Thunks_Init = True
End Function


' ============================================================
' Thunks_AcquireSlot
'   フリーリストから空きスロットを 1 個取得し、そのスロットの
'   サンク領域に「pSelfObj を Me に注入して pTargetFunc を呼ぶ」
'   マシンコードを書き込み、+55 にサンクのアドレスを書き込む。
'
'   第9.3a 段階では AcquireHandlerFor から呼ばれることが基本だが、
'   Test_TwoHandlers_Mock のように直接叩くケースも引き続き有効。
'
'   戻り値: スロット先頭アドレス、失敗時 0
' ============================================================
Public Function Thunks_AcquireSlot( _
    ByVal handler As ComCallbackHandler, _
    ByVal pSelfObj As LongPtr, _
    ByVal pTargetFunc As LongPtr) As LongPtr

    If m_pRegionBase = 0 Then Exit Function
    If m_freeHead < 0 Then Exit Function     ' プール枯渇
    If handler Is Nothing Then Exit Function

    ' フリーリストから 1 個取り出す
    Dim idx As Long
    idx = m_freeHead
    m_freeHead = m_freeNext(idx)
    m_freeNext(idx) = -1                      ' -1 = 使用中

    Dim pSlot As LongPtr
    pSlot = SlotAddrAt(idx)

    ' サンクを書き込む (offset +96 から 74 bytes)
    WriteThunkMachineCode pSlot + THUNK_OFFSET, pSelfObj, pTargetFunc, m_pRegionBase

    ' スタブクローンの +55 にサンクのアドレスを書き込む
    MemLongPtr(pSlot + LATE_BIND_OFFSET) = pSlot + THUNK_OFFSET

    ' Handler オブジェクトを idx で対応付け
    Set m_handlers(idx) = handler

    m_inUse = m_inUse + 1
    Thunks_AcquireSlot = pSlot
End Function


' ============================================================
' Thunks_ReleaseSlot
'   指定されたスロットをフリーリストに返却する。
'   m_handlers(idx) もクリアする。
' ============================================================
Public Sub Thunks_ReleaseSlot(ByVal pSlot As LongPtr)
    If m_pRegionBase = 0 Then Exit Sub
    If pSlot = 0 Then Exit Sub

    Dim idx As Long
    idx = SlotIndexFromAddr(pSlot)
    If idx < 0 Then Exit Sub
    If m_freeNext(idx) <> -1 Then Exit Sub           ' 既に空き = 二重解放

    ' Handler 対応表もクリア
    Set m_handlers(idx) = Nothing

    ' フリーリスト先頭に戻す
    m_freeNext(idx) = m_freeHead
    m_freeHead = idx

    m_inUse = m_inUse - 1
End Sub


' ============================================================
' SlotIndexFromAddr (内部ヘルパ)
'   サンクのアドレス (= スロット先頭アドレス) から
'   スロット index を逆算する。
' ============================================================
Private Function SlotIndexFromAddr(ByVal pSlot As LongPtr) As Long
    SlotIndexFromAddr = -1
    If m_pRegionBase = 0 Then Exit Function

    Dim offset As LongPtr
    offset = pSlot - (m_pRegionBase + HEADER_SIZE)
    If offset < 0 Then Exit Function
    If (offset Mod SLOT_SIZE) <> 0 Then Exit Function

    Dim idx As Long
    idx = CLng(offset \ SLOT_SIZE)
    If idx < 0 Or idx >= SLOT_COUNT Then Exit Function

    SlotIndexFromAddr = idx
End Function


' ============================================================
' SlotIndexFromVTableObjAddr (内部ヘルパ)
'   vtableObj のアドレス (= WebView2 が AddRef/Release で this として渡してくる値、
'                     = pSlot + VTABLE_OBJ_OFFSET) から、スロット index を逆算する。
' ============================================================
Private Function SlotIndexFromVTableObjAddr(ByVal pVTableObj As LongPtr) As Long
    SlotIndexFromVTableObjAddr = -1
    If m_pRegionBase = 0 Then Exit Function

    Dim offset As LongPtr
    offset = pVTableObj - (m_pRegionBase + HEADER_SIZE + VTABLE_OBJ_OFFSET)
    If offset < 0 Then Exit Function
    If (offset Mod SLOT_SIZE) <> 0 Then Exit Function

    Dim idx As Long
    idx = CLng(offset \ SLOT_SIZE)
    If idx < 0 Or idx >= SLOT_COUNT Then Exit Function

    SlotIndexFromVTableObjAddr = idx
End Function


' ============================================================
' Thunks_Shutdown
'   生存フラグを 0 に倒してから VirtualFree で領域を解放する。
'
'   第9.5 段階での追加処理:
'     - 末尾で Sentinel_ClearPrevRegion を呼んで環境変数の痕跡を消す。
'       消し忘れると次回 Init で「既に解放済みのアドレス」を再度
'       VirtualFree しようとして二重解放になる。VirtualFree は害なく
'       0 を返すだけだが、痕跡を残さないのが正解。
' ============================================================
Public Sub Thunks_Shutdown()
    If m_pRegionBase = 0 Then Exit Sub

    ' まず生存フラグを倒す
    MemLongPtr(m_pRegionBase) = 0^

    ' 全 Handler の owner を切って参照を解放 (循環参照防止)
    Dim i As Long
    For i = 0 To SLOT_COUNT - 1
        If Not (m_handlers(i) Is Nothing) Then
            m_handlers(i).ClearOwner
            Set m_handlers(i) = Nothing
        End If
    Next i

    ' 領域を解放 (第9.5 段階で診断ログ追加: 戻り値と LastError を観察)
    Dim shutFreeResult As Long
    Dim shutFreeErr As Long
    shutFreeResult = VirtualFree(m_pRegionBase, 0, MEM_RELEASE)
    shutFreeErr = GetLastError()
    Debug.Print "Thunks_Shutdown: VirtualFree(" & m_pRegionBase & ") returned " & _
                shutFreeResult & ", LastError = " & shutFreeErr & _
                IIf(shutFreeResult <> 0, " [OK]", " [NG] (" & VirtualFreeErrorName(shutFreeErr) & ")")

    m_pRegionBase = 0
    m_freeHead = -1
    m_inUse = 0
    Erase m_freeNext

    ' --- センチネル機構 (第9.5 段階追加) ---
    '   通常終了パスでは痕跡を残さない。これを忘れると次回 Init で
    '   既に解放済みのアドレスを VirtualFree しようとして無駄な処理
    '   (実害は無いが Debug.Print 上の "VirtualFree returned 0" が
    '    紛らわしくなる) が発生する。
    Sentinel_ClearPrevRegion
End Sub


' ============================================================
' スロットインデックス → スロット先頭アドレス
' ============================================================
Private Function SlotAddrAt(ByVal idx As Long) As LongPtr
    SlotAddrAt = m_pRegionBase + HEADER_SIZE + CLngLng(idx) * SLOT_SIZE
End Function


' ============================================================
' サンクのマシンコードを指定アドレスに書き込む (有効長 74 bytes)
'
'   レイアウト:
'     +0..+17  : 生存フラグチェック (18 bytes)
'     +18..+73 : 既存サンク本体 (56 bytes、第六?八段階と同一)
'
'   実装上の注意:
'     バイト配列 b は 80 bytes (= 8 の倍数) で宣言し、末尾 6 bytes は
'     ゼロパディング扱い。書き込み先のスロット +96..+175 (= 80 bytes ぶん)
'     はもともとパディング/未使用領域なので、80 bytes 書き込んでも無害。
' ============================================================
Private Sub WriteThunkMachineCode( _
    ByVal addr As LongPtr, _
    ByVal pSelfObj As LongPtr, _
    ByVal pTargetFunc As LongPtr, _
    ByVal pRegionBase As LongPtr)

    Dim b(0 To THUNK_BUF_SIZE - 1) As Byte, i As Long
    i = 0

    ' --- 生存フラグチェック (18 bytes) ---
    ' mov rax, imm64   (= pAliveFlag = pRegionBase + 0)
    b(i) = &H48: b(i + 1) = &HB8: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pRegionBase: i = i + 8
    ' cmp byte ptr [rax], 1
    b(i) = &H80: b(i + 1) = &H38: b(i + 2) = &H1: i = i + 3
    ' je +3 (= alive ラベル、xor + ret の 3 bytes をスキップ)
    b(i) = &H74: b(i + 1) = &H3: i = i + 2
    ' xor eax, eax
    b(i) = &H33: b(i + 1) = &HC0: i = i + 2
    ' ret
    b(i) = &HC3: i = i + 1

    ' --- 以降、第六?八段階と同一の 56 bytes サンク本体 ---

    ' sub rsp, 0x38
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HEC: b(i + 3) = &H38: i = i + 4
    ' mov r9, r8
    b(i) = &H4D: b(i + 1) = &H89: b(i + 2) = &HC1: i = i + 3
    ' mov r8, rdx
    b(i) = &H49: b(i + 1) = &H89: b(i + 2) = &HD0: i = i + 3
    ' mov rdx, rcx
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &HCA: i = i + 3
    ' mov rcx, imm64   (pSelfObj)
    b(i) = &H48: b(i + 1) = &HB9: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pSelfObj: i = i + 8
    ' lea rax, [rsp+0x28]
    b(i) = &H48: b(i + 1) = &H8D: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H28: i = i + 5
    ' mov [rsp+0x20], rax
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H20: i = i + 5
    ' mov rax, imm64   (pTargetFunc)
    b(i) = &H48: b(i + 1) = &HB8: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pTargetFunc: i = i + 8
    ' call rax
    b(i) = &HFF: b(i + 1) = &HD0: i = i + 2
    ' mov eax, [rsp+0x28]
    b(i) = &H8B: b(i + 1) = &H44: b(i + 2) = &H24: b(i + 3) = &H28: i = i + 4
    ' add rsp, 0x38
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HC4: b(i + 3) = &H38: i = i + 4
    ' ret
    b(i) = &HC3: i = i + 1
    ' int3 / int3 (パディング)
    b(i) = &HCC: i = i + 1
    b(i) = &HCC: i = i + 1

    ' --- 80 bytes を 8 byte 単位で書き込む (末尾 6 bytes は 0 パディング、無害) ---
    Dim k As Long
    For k = 0 To THUNK_BUF_SIZE - 1 Step 8
        MemLongPtr(addr + k) = ReadLongPtrFromBytes(b, k)
    Next k
End Sub


' ============================================================
' バイト配列の指定オフセットから 8 bytes (LongPtr) を取り出す
' ============================================================
Private Function ReadLongPtrFromBytes(ByRef b() As Byte, ByVal offset As Long) As LongPtr
    ReadLongPtrFromBytes = ReadLongPtr(VarPtr(b(offset)))
End Function


' ============================================================
' クラスのvTableから固定スロットの関数アドレスを取得
' ============================================================
Private Function GetClassMethodAddrAtFixedSlot( _
    ByVal cls As Object, ByVal slotIndex As Long) As LongPtr

    If cls Is Nothing Then Exit Function
    Dim pObj As LongPtr: pObj = ObjPtr(cls)
    If pObj = 0 Then Exit Function

    Dim pVTable As LongPtr
    pVTable = ReadLongPtr(pObj)
    If pVTable = 0 Then Exit Function

    GetClassMethodAddrAtFixedSlot = ReadLongPtr(pVTable + slotIndex * PtrSize)
End Function


' ============================================================
' IUnknown スタブ群（標準モジュール）
'
'   第9.2段階での仕様:
'     - this から SlotIndexFromVTableObjAddr で idx を逆引き
'     - m_handlers(idx) を直接操作
'     - これにより複数ハンドラを並行管理できる
'
'   第9.7a 段階で導入された IID チェック:
'     - Handler_QueryInterface が riid を本物の IID と比較する
'     - IID_IUnknown または「自分の HandlerKind に対応する本来の IID」のみ S_OK
'     - それ以外は ppvObject = 0 + E_NOINTERFACE
'     - 失敗時は riid.Data1 を Hex でログ出力 (成功時はノイズ削減のためログなし)
'
'   COM 規約遵守ポイント:
'     - 成功時 (S_OK 返却時) は AddRef する
'     - 失敗時 (E_NOINTERFACE 返却時) は AddRef しない、ppvObject に NULL を書く
'
'   トラブル時の切り戻し手順はモジュール冒頭コメントを参照。
' ============================================================
Private Function GetAddr(ByVal addr As LongPtr) As LongPtr
    GetAddr = addr
End Function

Private Function Handler_QueryInterface( _
    ByVal this As LongPtr, _
    ByVal riid As LongPtr, _
    ByRef ppvObject As LongPtr) As Long

    ' --- 防御: riid が NULL なら E_POINTER ---
    If riid = 0 Then
        ppvObject = 0
        Handler_QueryInterface = &H80004003   ' E_POINTER
        Exit Function
    End If

    ' --- 防御: this が無効なら E_NOINTERFACE ---
    Dim idx As Long
    idx = SlotIndexFromVTableObjAddr(this)
    If idx < 0 Then
        ppvObject = 0
        Handler_QueryInterface = E_NOINTERFACE
        Exit Function
    End If
    If m_handlers(idx) Is Nothing Then
        ppvObject = 0
        Handler_QueryInterface = E_NOINTERFACE
        Exit Function
    End If

    ' --- riid と IID_IUnknown を比較 ---
    If IsEqualGUIDInPlace(riid, m_iidIUnknown) Then
        ppvObject = this
        HandlerAddRefInternal this
        Handler_QueryInterface = S_OK
        Exit Function
    End If

    ' --- riid と「自分の本来の IID」を比較 ---
    Dim kind As HandlerKind
    kind = m_handlers(idx).kind
    If IsEqualGUIDInPlace(riid, m_iidTable(kind)) Then
        ppvObject = this
        HandlerAddRefInternal this
        Handler_QueryInterface = S_OK
        Exit Function
    End If

    ' --- どの IID とも一致しない: E_NOINTERFACE ---
    '   riid.Data1 (先頭 4 byte、little-endian の Long) をログに出す。
    '   QueryInterface が呼ばれた瞬間に WebView2 がどの IID を試したかが
    '   観察できる (お行儀 IID = IMarshal / IAgileObject 等の検出にも有用)。
    Dim data1 As Long
    data1 = LongPtrLowDword(ReadLongPtr(riid))
    Debug.Print "  QI rejected [idx=" & idx & " kind=" & kind & _
                "] riid.Data1=&H" & Hex(data1) & " → E_NOINTERFACE"
    ppvObject = 0
    Handler_QueryInterface = E_NOINTERFACE
End Function

Private Function Handler_AddRef(ByVal this As LongPtr) As Long
    Handler_AddRef = HandlerAddRefInternal(this)
End Function

Private Function Handler_Release(ByVal this As LongPtr) As Long
    Handler_Release = HandlerReleaseInternal(this)
End Function


' ============================================================
' IsEqualGUIDInPlace (第9.7a 段階で新規)
'
'   2 つの GUID (16 byte) を比較する。
'   片方は riid (ポインタ、WebView2 から渡される)、
'   もう片方は m_iidTable / m_iidIUnknown (VBA 側の GUID 型変数)。
'
'   実装方式: LongPtr (= x64 で 8 byte) を 2 個読んで比較する高速版。
'     - 前半 8 byte: Data1 (4) + Data2 (2) + Data3 (2)
'     - 後半 8 byte: Data4(0..7)
'
'   VBA 側の GUID 型変数のアドレスは VarPtr で取得する。
'   x64 では GUID 型は 16 byte 連続配置 (パディングなし、データ並びは
'   Win32 標準の little-endian と一致)。これは Win32 GUID 構造体と
'   VBA の Type 宣言が偶然 (というか自然に) 一致しているため。
'
'   検証ポイント:
'     - VBA の Type が本当にパディングなしで 16 byte 連続配置になるか
'     - Long / Integer が x64 メモリ上で little-endian で並ぶか
'   いずれも x64 VBA の標準動作だが、もし不一致なら IID チェック自体が
'   失敗する (riid と m_iidTable が一致しない) ため、検証で気付ける。
' ============================================================
Private Function IsEqualGUIDInPlace(ByVal pRiid As LongPtr, ByRef refGuid As GUID) As Boolean
    Dim pRef As LongPtr
    pRef = VarPtr(refGuid)

    If ReadLongPtr(pRiid) <> ReadLongPtr(pRef) Then Exit Function
    If ReadLongPtr(pRiid + 8) <> ReadLongPtr(pRef + 8) Then Exit Function

    IsEqualGUIDInPlace = True
End Function


' ============================================================
' LongPtrLowDword (補助ヘルパー、第9.7a 段階で新規)
'
'   LongPtr の下位 32 bit を Long として取り出す。riid.Data1 (最初の
'   4 byte) を Handler_QueryInterface のログ出力で使う。
'   ComCallbackHandler.cls にも同名の Private 関数があるが、こちらは Module9 内専用。
'   実装は同じ。
' ============================================================
Private Function LongPtrLowDword(ByVal v As LongPtr) As Long
    Dim u As LongLong
    u = CLngLng(v) And &HFFFFFFFF^
    If u > &H7FFFFFFF^ Then
        LongPtrLowDword = CLng(u - &H100000000^)
    Else
        LongPtrLowDword = CLng(u)
    End If
End Function


' ============================================================
' InitIIDTable (第9.7a 段階で新規)
'
'   m_iidIUnknown と m_iidTable(HK_xxx) を初期化する。
'   Thunks_Init の末尾から 1 回だけ呼ばれる。
'
'   IID 値の出典: WebView2.h (Microsoft Edge WebView2 SDK)。
'   各 IID は WebView2_IID_vtable_reference.md に記載済み。
'
'   GUID の各フィールドへの書き込み方:
'     Data1 (Long、4 byte little-endian): 整数リテラル
'     Data2 (Integer、2 byte little-endian): 整数リテラル (符号付き範囲のため、
'       &H8000 以上は & サフィックスで Long 化してから CInt で Integer に
'       戻す手間がある。本実装では FillGUID で String パース方式を採用し、
'       コード上で整数値を直接書かないことで &H リテラル罠 (設計原則 29) を回避)
' ============================================================
Private Sub InitIIDTable()
    ' --- IID_IUnknown (Win32 標準、{00000000-0000-0000-C000-000000000046}) ---
    FillGUID m_iidIUnknown, "00000000-0000-0000-C000-000000000046"

    ' --- HandlerKind ごとの本来の IID ---

    ' ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler
    FillGUID m_iidTable(HK_EnvironmentCompleted), _
             "4e8a3389-c9d8-4bd2-b6b5-124fee6cc14d"

    ' ICoreWebView2CreateCoreWebView2ControllerCompletedHandler
    FillGUID m_iidTable(HK_ControllerCompleted), _
             "6c4819f3-c9b7-4260-8127-c9f5bde7f68c"

    ' ICoreWebView2NavigationStartingEventHandler
    FillGUID m_iidTable(HK_NavigationStarting), _
             "9adbe429-f36d-432b-9ddc-f8881fbd76e3"

    ' ICoreWebView2NavigationCompletedEventHandler
    FillGUID m_iidTable(HK_NavigationCompleted), _
             "d33a35bf-1c49-4f98-93ab-006e0533fe1c"

    ' ICoreWebView2WebMessageReceivedEventHandler
    FillGUID m_iidTable(HK_WebMessageReceived), _
             "57213f19-00e6-49fa-8e07-898ea01ecbd2"

    ' ICoreWebView2DocumentTitleChangedEventHandler
    FillGUID m_iidTable(HK_DocumentTitleChanged), _
             "f5f2b923-953e-4042-9f95-f3a118e1afd4"

    ' ICoreWebView2NewWindowRequestedEventHandler
    FillGUID m_iidTable(HK_NewWindowRequested), _
             "d4c185fe-c81c-4989-97af-2d3fa7ab5651"

    ' ICoreWebView2ExecuteScriptCompletedHandler (第9.8c で追加)
    FillGUID m_iidTable(HK_ExecuteScriptCompleted), _
             "49511172-cc67-4bca-9923-137112f4c4cc"

    ' ICoreWebView2HistoryChangedEventHandler (第9.9b で追加)
    '   出典: WebView2.h L3790
    FillGUID m_iidTable(HK_HistoryChanged), _
             "c79a420c-efd9-4058-9295-3e8b4bcab645"

    ' ICoreWebView2DOMContentLoadedEventHandler (第9.9b で追加)
    '   出典: WebView2.h L5718
    FillGUID m_iidTable(HK_DOMContentLoaded), _
             "4bac7e9c-199e-49ed-87ed-249303acf019"
End Sub


' ============================================================
' FillGUID (第9.7a 段階で新規、第9.9a で Public 昇格)
'
'   "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" 形式の文字列から GUID 構造体を埋める。
'   コード上で &H リテラルを直接書くと罠 (設計原則 29) に嵌まる場合があるため、
'   文字列経由でパースする方式を採用。
'
'   入力例: "4e8a3389-c9d8-4bd2-b6b5-124fee6cc14d"
'   分割:
'     g(0..7)   = "4e8a3389"  → Data1 (Long)
'     g(9..12)  = "c9d8"      → Data2 (Integer)
'     g(14..17) = "4bd2"      → Data3 (Integer)
'     g(19..22) = "b6b5"      → Data4(0..1)
'     g(24..35) = "124fee6cc14d" → Data4(2..7)
'
'   各 16 進文字列は HexStrToLong / HexStrToInt で数値化。
'   Data2 / Data3 は Integer (符号付き 16 bit)。&HC9D8 のような値は
'   素直に CInt() すると Overflow するので、一旦 Long で読んでから
'   符号反転処理を入れる必要がある。本実装ではそのための専用ヘルパー
'   HexStrToInt を用意。
'
'   第9.9a で Public 昇格:
'     Wv2Pane.EnsureView2 から、QueryInterface 用のローカル GUID 構造体を
'     初期化するために本関数を呼ぶ。Wv2Thunks の InitIIDTable 専用ユーティリティ
'     だったものを、共通ユーティリティに格上げ。GUID 型 (Public Type) と
'     ペアで使われる基本機能なので、公開しても自然 (= 設計原則 50 として
'     新設、9.9a 検証後に確定予定)。
' ============================================================
Public Sub FillGUID(ByRef g As GUID, ByVal s As String)
    g.data1 = HexStrToLong(Mid$(s, 1, 8))
    g.Data2 = HexStrToInt(Mid$(s, 10, 4))
    g.Data3 = HexStrToInt(Mid$(s, 15, 4))

    ' Data4(0..7): 残り 4 バイト分 (Mid 20..23) + 後半 12 桁 (Mid 25..36)
    g.Data4(0) = CByte("&H" & Mid$(s, 20, 2))
    g.Data4(1) = CByte("&H" & Mid$(s, 22, 2))
    g.Data4(2) = CByte("&H" & Mid$(s, 25, 2))
    g.Data4(3) = CByte("&H" & Mid$(s, 27, 2))
    g.Data4(4) = CByte("&H" & Mid$(s, 29, 2))
    g.Data4(5) = CByte("&H" & Mid$(s, 31, 2))
    g.Data4(6) = CByte("&H" & Mid$(s, 33, 2))
    g.Data4(7) = CByte("&H" & Mid$(s, 35, 2))
End Sub


' ============================================================
' HexStrToLong (第9.7a 段階で新規)
'
'   8 桁の 16 進文字列を Long に変換する。値が &H80000000 以上の場合は
'   負の値になる (Long は符号付き 32 bit、二補表現)。
'
'   VBA の CLng("&H...") は値が Long の範囲を超えると Overflow する。
'   そこで LongLong 経由で読んで、最上位ビットが立つときは
'   2 の補数表現に変換してから CLng で Long に落とす。
'
'   入力例: "4e8a3389" → 1317797257
'           "d33a35bf" → -752318529 (符号付き Long)
' ============================================================
Private Function HexStrToLong(ByVal s As String) As Long
    Dim v As LongLong
    v = CLngLng("&H" & s)                     ' 0..&HFFFFFFFF を LongLong (符号付き 64 bit) で受け取る
    If v >= &H80000000^ Then
        HexStrToLong = CLng(v - &H100000000^) ' 32 bit の符号反転処理 (2 の補数)
    Else
        HexStrToLong = CLng(v)
    End If
End Function


' ============================================================
' HexStrToInt (第9.7a 段階で新規)
'
'   4 桁の 16 進文字列を Integer (符号付き 16 bit) に変換する。
'   値が &H8000 以上の場合は負の値になる。
'
'   VBA の CInt("&H...") は &H8000 以上で Overflow する可能性がある
'   (= 「32768 を Integer の範囲に収めようとして例外」、これは設計原則 29
'   の罠と同根の問題)。
'
'   安全策として CLng で一旦 Long に取り、&H8000 以上なら手動で
'   2 の補数表現に変換して CInt する。
' ============================================================
Private Function HexStrToInt(ByVal s As String) As Integer
    Dim v As Long
    v = CLng("&H" & s)                        ' 0..&HFFFF の Long として取得
    If v >= &H8000& Then
        HexStrToInt = CInt(v - &H10000)       ' 負の Integer に変換
    Else
        HexStrToInt = CInt(v)
    End If
End Function


' ============================================================
' HandlerAddRefInternal / HandlerReleaseInternal
'
'   第9.2段階での実装:
'     - this (= pSlot + VTABLE_OBJ_OFFSET) から idx を求める
'     - m_handlers(idx) のオブジェクトを操作する
'     - 不正な this は安全に 0 を返す
'
'   第9.3a 段階での変更:
'     - HandlerReleaseInternal の自動解放経路から、末尾の
'       「gHandler Is h」分岐を削除 (gHandler 自体が撤廃されたため)
'
'   戻り値: 操作後の refcount
' ============================================================
Private Function HandlerAddRefInternal(ByVal this As LongPtr) As Long
    Dim idx As Long
    idx = SlotIndexFromVTableObjAddr(this)
    If idx < 0 Then Exit Function
    If m_handlers(idx) Is Nothing Then Exit Function

    Dim n As Long
    n = m_handlers(idx).RefCount + 1
    m_handlers(idx).RefCount = n
    Debug.Print "  AddRef [idx=" & idx & "] → refCount = " & n
    HandlerAddRefInternal = n
End Function

Private Function HandlerReleaseInternal(ByVal this As LongPtr) As Long
    Dim idx As Long
    idx = SlotIndexFromVTableObjAddr(this)
    If idx < 0 Then Exit Function
    If m_handlers(idx) Is Nothing Then Exit Function

    Dim n As Long
    n = m_handlers(idx).RefCount - 1
    If n < 0 Then n = 0                ' 防御 (本来あり得ない)
    m_handlers(idx).RefCount = n
    Debug.Print "  Release [idx=" & idx & "] → refCount = " & n

    If n = 0 Then
        ' refcount が 0 になったので、循環参照を切ってスロットを解放する。
        ' 順序が重要 (第9.1段階以降同じ):
        '   1. ClearOwner で循環参照を切る (ComCallbackHandler → 上位クラスの参照を切る)
        '   2. ローカル参照を保持してから m_handlers(idx) を Nothing に
        '   3. Thunks_ReleaseSlot でスロットをフリーリストに戻す
        '   4. ローカル参照も Nothing に (ここで ComCallbackHandler が完全解放される)
        '
        ' 第9.3a での変更: 旧 4. の「gHandler Is h なら gHandler も切る」
        ' という分岐は不要になった (gHandler が撤廃されたため)。

        Dim h As ComCallbackHandler
        Set h = m_handlers(idx)

        Dim pSlot As LongPtr
        pSlot = h.Slot

        ' 1. 循環参照を切る (ComCallbackHandler.m_owner = Nothing)
        '    これにより上位クラス (Wv2Environment 等) への参照が切れ、
        '    上位クラス側でも ComCallbackHandler への参照が切られていれば
        '    両方が GC で解放される流れに乗る。
        h.ClearOwner

        ' 2. m_handlers(idx) を Nothing に
        Set m_handlers(idx) = Nothing

        ' 3. スロットをフリーリストに戻す
        Thunks_ReleaseSlot pSlot

        ' 4. h はこのプロシージャを抜ける時点で参照が切れて解放される
        Set h = Nothing

        Debug.Print "  Release [idx=" & idx & "] → refCount = 0 → Slot " & idx & " auto-released"
    End If

    HandlerReleaseInternal = n
End Function


' ============================================================
' PointerAccessor によるメモリ操作プリミティブ
' ============================================================
Private Property Let MemLongPtr(ByVal addr As LongPtr, ByVal newValue As LongPtr)
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.cbElements = PtrSize
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        .arr(0) = newValue
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Property

Private Function ReadLongPtr(ByVal addr As LongPtr) As LongPtr
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.cbElements = PtrSize
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        ReadLongPtr = .arr(0)
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Function

Private Sub WritePtrNatively(ByRef ptrs() As LONG_PTR, ByVal ptr As LongPtr)
    ptrs(0) = ptr
End Sub



