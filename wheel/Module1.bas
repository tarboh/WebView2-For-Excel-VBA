Attribute VB_Name = "Module1"
'==============================================================================
' UserFormWheel 使用例
'
' 最小構成:
'   1. UserForm (例: UserForm1) を用意
'   2. UserForm1 のコードモジュールに以下を記述:
'
'      Option Explicit
'      Private m_Wheel As UserFormWheel
'
'      Private Sub UserForm_Activate()
'          Set m_Wheel = New UserFormWheel
'          m_Wheel.Attach Me
'      End Sub
'
'      Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'          If Not m_Wheel Is Nothing Then m_Wheel.Detach
'          Set m_Wheel = Nothing
'      End Sub
'
'   3. 動作確認手順:
'      a) UserForm にスクロール可能な TextBox (MultiLine=True) を配置
'      b) ListBox を配置して AddItem で複数項目追加
'      c) 十分に長い内容を入れてホイール操作
'      d) Shift+ホイールで横スクロール (Frame 使用時)
'
'   4. クラッシュ耐性確認:
'      a) UserForm を開いた状態で別モジュールのマクロを Ctrl+Break で停止
'      b) その状態で UserForm 上でホイール操作
'      c) Break 中は何も起きないが、クラッシュしないことを確認
'==============================================================================
Option Explicit

' テスト用 UserForm 表示
Public Sub ShowTestForm()
    UserForm1.Show
End Sub
