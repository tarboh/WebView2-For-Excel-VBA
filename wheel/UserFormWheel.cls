VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserFormWheel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' UserFormWheel.cls  (x64専用)
'
' UserForm とその上のスクロール可能コントロール (TextBox/ListBox/Frame/
' UserForm 自身) にマウスホイール対応を提供する。
'
' 使い方:
'   Private m_Wheel As UserFormWheel
'
'   Private Sub UserForm_Activate()
'       Set m_Wheel = New UserFormWheel
'       m_Wheel.Attach Me
'   End Sub
'
'   Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'       If Not m_Wheel Is Nothing Then m_Wheel.Detach
'       Set m_Wheel = Nothing
'   End Sub
'==============================================================================
Option Explicit

'==============================================================================
' Win32 API
'==============================================================================
Private Declare PtrSafe Function SetWindowLongPtrA Lib "user32" ( _
    ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowLongPtrA Lib "user32" ( _
    ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function CallWindowProcA Lib "user32" ( _
    ByVal lpPrevWndFunc As LongPtr, ByVal hwnd As LongPtr, _
    ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function DefWindowProcA Lib "user32" ( _
    ByVal hwnd As LongPtr, ByVal Msg As Long, _
    ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function FindWindowA Lib "user32" ( _
    ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function ScreenToClient Lib "user32" ( _
    ByVal hwnd As LongPtr, ByRef lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function GetDpiForWindow Lib "user32" ( _
    ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function SendMessageA Lib "user32" ( _
    ByVal hwnd As LongPtr, ByVal Msg As Long, _
    ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" ( _
    ByVal vKey As Long) As Integer

Private Const GWLP_WNDPROC As Long = -4
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_MOUSEHWHEEL As Long = &H20E
Private Const MK_SHIFT As Long = &H4
Private Const MK_CONTROL As Long = &H8

' TextBox への EM メッセージ
Private Const EM_LINESCROLL As Long = &HB6
Private Const EM_GETFIRSTVISIBLELINE As Long = &HCE
Private Const EM_LINEINDEX As Long = &HBB

' VK コード
Private Const VK_SHIFT As Long = &H10

Private Type POINTAPI
    x As Long
    y As Long
End Type

'==============================================================================
' VTable インデックス
'   MyWndProc を本クラスの最初の Public メソッドとして宣言する必要がある。
'   IUnknown(3) + IDispatch(4) = 7 → vtable[7]
'==============================================================================
Private Enum VTableIndex
    MyWndProc_Slot = 7
End Enum

'==============================================================================
' インスタンス状態
'==============================================================================
Private m_Form As Object              ' 対象 UserForm (MSForms.UserForm 遅延バインド)
Private m_hwnd As LongPtr             ' UserForm の hwnd
Private m_pThunk As LongPtr           ' サブクラス化プロシージャ
Private m_ThunkSize As LongPtr
Private m_OrigWndProc As LongPtr      ' 元の WndProc
Private m_Attached As Boolean

' スクロール量 (既定: 3行/tick)
Public LinesPerNotch As Long

'==============================================================================
' MyWndProc
'   ***このメソッドは必ず最初の Public メソッドとして宣言すること***
'   ディスパッチャから vtable[7] 経由で呼ばれる。
'==============================================================================
Public Function MyWndProc(ByVal hwnd As LongPtr, ByVal Msg As Long, _
                          ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    On Error Resume Next

    Select Case Msg
    Case WM_MOUSEWHEEL, WM_MOUSEHWHEEL
        Dim handled As Boolean
        handled = HandleWheel(Msg, wParam, lParam)
        If handled Then
            MyWndProc = 0
            Exit Function
        End If
    End Select

    ' 未処理 → 元の WndProc に委譲
    MyWndProc = CallWindowProcA(m_OrigWndProc, hwnd, Msg, wParam, lParam)
End Function

'==============================================================================
' 公開 API
'==============================================================================
Public Sub Attach(ByVal targetForm As Object)
    If m_Attached Then Exit Sub

    LinesPerNotch = 3
    Set m_Form = targetForm
    m_hwnd = FindUserFormHwnd(targetForm)
    If m_hwnd = 0 Then
        Err.Raise 5, "Attach", "UserForm の hwnd を特定できませんでした"
    End If

    ' ディスパッチャとスタブパッチは初回のみ
    ModHookSupport.EnsureDispatcherInstalled VTableIndex.MyWndProc_Slot

    ' サブクラス化する前に現在の (オリジナル) WndProc を取得
    ' → サンクにフィルタ不一致時のジャンプ先として焼き込むため
    m_OrigWndProc = GetWindowLongPtrA(m_hwnd, GWLP_WNDPROC)
    If m_OrigWndProc = 0 Then
        Err.Raise 7, "Attach", "GetWindowLongPtr failed"
    End If

    ' サンク生成 (WM_MOUSEWHEEL と WM_MOUSEHWHEEL のみ捕捉、それ以外は
    '             m_OrigWndProc へ末尾ジャンプで即委譲)
    m_pThunk = ModHookSupport.BuildInstanceThunk( _
        ObjPtr(Me), WM_MOUSEWHEEL, WM_MOUSEHWHEEL, m_OrigWndProc, m_ThunkSize)

    ' サブクラス化登録
    Dim prev As LongPtr
    prev = SetWindowLongPtrA(m_hwnd, GWLP_WNDPROC, m_pThunk)
    If prev = 0 Then
        ModHookSupport.FreeThunk m_pThunk
        m_pThunk = 0
        m_OrigWndProc = 0
        Err.Raise 7, "Attach", "SetWindowLongPtr failed"
    End If
    ' prev == m_OrigWndProc であるはず (念のため保持)
    m_Attached = True
    
    Debug.Print "hwnd=" & Hex(m_hwnd) & " pThunk=" & Hex(m_pThunk) & " origWndProc=" & Hex(m_OrigWndProc)
    
End Sub

Public Sub Detach()
    If Not m_Attached Then Exit Sub

    ' サブクラス化解除
    If m_hwnd <> 0 And m_OrigWndProc <> 0 Then
        SetWindowLongPtrA m_hwnd, GWLP_WNDPROC, m_OrigWndProc
    End If

    ' サンク解放は WndProc 呼出が確実に終わってからでないと危険なので、
    ' 少し待機 (Detach 直後に OS がまだ呼び出し中の可能性)
    DoEvents

    If m_pThunk <> 0 Then
        ModHookSupport.FreeThunk m_pThunk
        m_pThunk = 0
    End If

    m_OrigWndProc = 0
    m_hwnd = 0
    Set m_Form = Nothing
    m_Attached = False
End Sub

Private Sub Class_Terminate()
    Detach
End Sub

'==============================================================================
' UserForm の hwnd 取得
'   MSForms.UserForm の hwnd は直接取得できないので、以下の方法を使う:
'     1. FindWindowA で "ThunderDFrame" (UserForm のクラス名) を探す
'        ただしこれだと複数 UserForm が開いている時に誤検出する
'     2. Caption をキャプションでマッチング
'==============================================================================
Private Function FindUserFormHwnd(ByVal targetForm As Object) As LongPtr
    On Error Resume Next
    Dim cap As String
    cap = targetForm.Caption
    FindUserFormHwnd = FindWindowA("ThunderDFrame", cap)

    ' 見つからなかった場合、VBA7 の UserForm は "ThunderXFrame" の可能性
    If FindUserFormHwnd = 0 Then
        FindUserFormHwnd = FindWindowA("F3 Server 80x80", cap)
    End If
End Function

'==============================================================================
' ホイール処理本体
'==============================================================================
Private Function HandleWheel(ByVal Msg As Long, ByVal wParam As LongPtr, _
                              ByVal lParam As LongPtr) As Boolean
    ' wParam 上位16bit = delta, 下位16bit = modifier keys
    Dim delta As Long
    delta = HiWord(wParam)

    Dim shiftHeld As Boolean
    shiftHeld = (LoWord(wParam) And MK_SHIFT) <> 0

    ' スクロール方向決定
    Dim isHorizontal As Boolean
    isHorizontal = (Msg = WM_MOUSEHWHEEL) Or shiftHeld

    ' lParam はスクリーン座標
    Dim sx As Long, sy As Long
    sx = LoWord(lParam)
    sy = HiWord(lParam)
    ' HiWord は符号付き扱いが必要
    If sy >= &H8000& Then sy = sy - &H10000
    If sx >= &H8000& Then sx = sx - &H10000

    ' カーソル下のコントロールを特定
    Dim ctrl As Object
    Set ctrl = FindControlAtScreenPoint(sx, sy)
    If ctrl Is Nothing Then
        ' どのコントロールにも該当しない → UserForm 自体 (スクロール可能なら)
        HandleWheel = TryScrollUserForm(delta, isHorizontal)
        Exit Function
    End If

    ' コントロール種別で分岐
    HandleWheel = ScrollControl(ctrl, delta, isHorizontal)
End Function

'==============================================================================
' カーソル下のコントロールを特定
'   MSForms コントロールは windowless なので UserForm の Controls を走査し、
'   MSForms 座標系 (ポイント) でヒットテストする。
'==============================================================================
Private Function FindControlAtScreenPoint(ByVal sx As Long, ByVal sy As Long) As Object
    On Error Resume Next

    ' スクリーン座標 → UserForm クライアント座標 (ピクセル)
    Dim pt As POINTAPI
    pt.x = sx
    pt.y = sy
    ScreenToClient m_hwnd, pt

    ' ピクセル → ポイント変換 (MSForms 座標)
    Dim dpi As Long
    dpi = GetDpiForWindow(m_hwnd)
    If dpi = 0 Then dpi = 96
    Dim fx As Double, fy As Double
    fx = pt.x * 72# / dpi
    fy = pt.y * 72# / dpi

    ' Controls コレクションを走査 (Frame 内再帰対応)
    Set FindControlAtScreenPoint = HitTestContainer(m_Form, fx, fy, 0#, 0#)
End Function

'==============================================================================
' コンテナ (UserForm or Frame) のコントロール配列をヒットテスト
'   offsetX/Y はコンテナの左上がフォーム座標系でどこにあるか
'==============================================================================
Private Function HitTestContainer(ByVal container As Object, _
                                   ByVal fx As Double, ByVal fy As Double, _
                                   ByVal offsetX As Double, ByVal offsetY As Double) As Object
    On Error Resume Next
    Dim i As Long
    Dim ctrl As Object
    Dim hit As Object

    ' 後ろから走査 (Z-order で手前にあるものを優先)
    Dim cnt As Long
    cnt = container.Controls.Count
    For i = cnt - 1 To 0 Step -1
        Set ctrl = container.Controls(i)
        If IsControlVisible(ctrl) Then
            Dim left_ As Double, top_ As Double, w As Double, h As Double
            left_ = ctrl.Left + offsetX
            top_ = ctrl.Top + offsetY
            w = ctrl.Width
            h = ctrl.Height

            If fx >= left_ And fx < left_ + w And fy >= top_ And fy < top_ + h Then
                ' Frame なら内部を再帰探索
                If TypeName(ctrl) = "Frame" Then
                    ' Frame 内座標はフレームの ScrollLeft/Top オフセットも考慮
                    Dim innerOffsetX As Double, innerOffsetY As Double
                    innerOffsetX = left_ - ctrl.ScrollLeft
                    innerOffsetY = top_ - ctrl.ScrollTop
                    Set hit = HitTestContainer(ctrl, fx, fy, innerOffsetX, innerOffsetY)
                    If Not hit Is Nothing Then
                        Set HitTestContainer = hit
                        Exit Function
                    End If
                    ' Frame 自身をヒットとする
                    Set HitTestContainer = ctrl
                    Exit Function
                Else
                    Set HitTestContainer = ctrl
                    Exit Function
                End If
            End If
        End If
    Next i

    Set HitTestContainer = Nothing
End Function

Private Function IsControlVisible(ByVal ctrl As Object) As Boolean
    On Error Resume Next
    IsControlVisible = ctrl.Visible
End Function

'==============================================================================
' コントロール種別に応じたスクロール処理
'==============================================================================
Private Function ScrollControl(ByVal ctrl As Object, ByVal delta As Long, _
                                ByVal isHorizontal As Boolean) As Boolean
    On Error Resume Next

    ' delta > 0 で上 / 左スクロール、delta < 0 で下 / 右スクロール
    Dim notches As Long
    notches = -Sgn(delta) * LinesPerNotch  ' Sgn = +1/-1
    If delta = 0 Then Exit Function

    Select Case TypeName(ctrl)
    Case "TextBox"
        If ctrl.Multiline Then
            ScrollControl = ScrollTextBox(ctrl, notches, isHorizontal)
        End If

    Case "ListBox"
        ScrollControl = ScrollListBox(ctrl, notches)

    Case "ComboBox"
        ' ドロップダウン中は MSForms のリスト部分が別ウィンドウになっており、
        ' UserForm のサブクラス化では捕捉できない。ここでは非対応。
        ScrollControl = False

    Case "Frame"
        ScrollControl = ScrollFrame(ctrl, notches, isHorizontal)

    Case Else
        ' その他: UserForm 自身のスクロールにフォールバック
        ScrollControl = TryScrollUserForm(delta, isHorizontal)
    End Select
End Function

'==============================================================================
' TextBox (MultiLine) スクロール
'   EM_LINESCROLL メッセージで行単位スクロール。
'   ただし MSForms の TextBox の hwnd は取得できない (windowless) ため、
'   CurLine/SelStart ベースで擬似的に行う。
'==============================================================================
Private Function ScrollTextBox(ByVal tb As Object, ByVal lines As Long, _
                                ByVal isHorizontal As Boolean) As Boolean
    On Error Resume Next
    If isHorizontal Then Exit Function  ' 横スクロールは未対応

    ' MSForms.TextBox には TopIndex が無いので CurLine を使う
    Dim curLine As Long
    curLine = tb.curLine
    Dim newLine As Long
    newLine = curLine + lines
    If newLine < 0 Then newLine = 0

    ' 目標行の先頭 SelStart を計算するには全テキストを改行で分割する
    Dim txt As String
    txt = tb.Text
    Dim arr() As String
    arr = Split(txt, vbCrLf)
    If newLine > UBound(arr) Then newLine = UBound(arr)

    Dim pos As Long, j As Long
    pos = 0
    For j = 0 To newLine - 1
        pos = pos + Len(arr(j)) + 2  ' +2 for vbCrLf
    Next j

    ' キャレット移動 → これで自動的にスクロールが追従
    tb.SelStart = pos
    tb.SelLength = 0
    ScrollTextBox = True
End Function

'==============================================================================
' ListBox スクロール
'   TopIndex プロパティ直接操作
'==============================================================================
Private Function ScrollListBox(ByVal lb As Object, ByVal lines As Long) As Boolean
    On Error Resume Next
    Dim newTop As Long
    newTop = lb.TopIndex + lines
    If newTop < 0 Then newTop = 0
    If newTop >= lb.ListCount Then newTop = lb.ListCount - 1
    If newTop < 0 Then newTop = 0
    lb.TopIndex = newTop
    ScrollListBox = True
End Function

'==============================================================================
' Frame スクロール
'==============================================================================
Private Function ScrollFrame(ByVal fr As Object, ByVal lines As Long, _
                              ByVal isHorizontal As Boolean) As Boolean
    On Error Resume Next

    ' 1 notch あたりの px 量を推定 (1 行 = 約 15pt)
    Dim step_ As Double
    step_ = lines * 15#

    If isHorizontal Then
        Dim nl As Double
        nl = fr.ScrollLeft + step_
        If nl < 0 Then nl = 0
        fr.ScrollLeft = nl
    Else
        Dim nt As Double
        nt = fr.ScrollTop + step_
        If nt < 0 Then nt = 0
        fr.ScrollTop = nt
    End If
    ScrollFrame = True
End Function

'==============================================================================
' UserForm 自身のスクロール (ScrollBars プロパティで有効化されている場合)
'==============================================================================
Private Function TryScrollUserForm(ByVal delta As Long, _
                                    ByVal isHorizontal As Boolean) As Boolean
    On Error Resume Next
    Dim notches As Long
    notches = -Sgn(delta) * LinesPerNotch
    Dim step_ As Double
    step_ = notches * 15#

    If isHorizontal Then
        Dim nl As Double
        nl = m_Form.ScrollLeft + step_
        If nl < 0 Then nl = 0
        m_Form.ScrollLeft = nl
    Else
        Dim nt As Double
        nt = m_Form.ScrollTop + step_
        If nt < 0 Then nt = 0
        m_Form.ScrollTop = nt
    End If
    TryScrollUserForm = True
End Function

'==============================================================================
' ヘルパー: wParam/lParam の 上位/下位 16bit 取得
'==============================================================================
Private Function LoWord(ByVal x As LongPtr) As Long
    LoWord = CLng(x And &HFFFF&)
End Function

Private Function HiWord(ByVal x As LongPtr) As Long
    Dim v As Long
    v = CLng((x \ &H10000) And &HFFFF&)
    If v >= &H8000& Then v = v - &H10000
    HiWord = v
End Function


