Attribute VB_Name = "ModTestMouseHook"
'==============================================================================
' MouseWheelHook 使用例
'==============================================================================
Option Explicit

' WithEvents でイベントを受け取る場合は ThisWorkbook やクラスモジュールに書く
' ここでは最小の動作確認として、モジュール変数で保持する
Private m_Hook As MouseWheelHook

Public Sub StartMouseHook()
    Set m_Hook = New MouseWheelHook
    m_Hook.StartHook
    Debug.Print "Mouse wheel hook started."
End Sub

Public Sub StopMouseHook()
    If Not m_Hook Is Nothing Then
        m_Hook.StopHook
        Set m_Hook = Nothing
    End If
    Debug.Print "Mouse wheel hook stopped."
End Sub

'------------------------------------------------------------------------------
' WithEvents でイベントを受けたい場合は、以下のようなクラスモジュールを作る:
'
' ' クラスモジュール: HookOwner.cls
' Option Explicit
' Private WithEvents m_Hook As MouseWheelHook
'
' Public Sub Init()
'     Set m_Hook = New MouseWheelHook
'     m_Hook.StartHook
' End Sub
'
' Private Sub m_Hook_WheelScroll(ByVal delta As Long, _
'                                 ByVal x As Long, ByVal y As Long, _
'                                 ByRef bHandled As Boolean)
'     Debug.Print "Wheel delta=" & delta & " at (" & x & "," & y & ")"
'     ' bHandled = True で後続フックにイベントを流さない
' End Sub
'------------------------------------------------------------------------------


