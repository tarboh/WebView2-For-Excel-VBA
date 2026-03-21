Attribute VB_Name = "mHandleManager"
' --- 標準モジュール：HandlerManager ---
Option Explicit
Public m_InstanceMap As Dictionary

' 登録：ハンドラの住所をキーにして、クラスインスタンスを紐付ける
Public Sub RegisterInstance(ByVal pHandler As LongPtr, ByRef obj As Object)
    If m_InstanceMap Is Nothing Then Set m_InstanceMap = New Dictionary 'CreateObject("Scripting.Dictionary")
    m_InstanceMap.Add CStr(pHandler), obj
End Sub

' 削除：解放時に呼び出す
Public Sub UnregisterInstance(ByVal pHandler As LongPtr)
    If Not m_InstanceMap Is Nothing Then
        If m_InstanceMap.Exists(CStr(pHandler)) Then m_InstanceMap.Remove CStr(pHandler)
    End If
End Sub

' 逆引き：Invoke 内で使用
Public Function GetInstance(ByVal pHandler As LongPtr) As Object
    If Not m_InstanceMap Is Nothing Then
        If m_InstanceMap.Exists(CStr(pHandler)) Then
            Set GetInstance = m_InstanceMap(CStr(pHandler))
        End If
    End If
End Function
