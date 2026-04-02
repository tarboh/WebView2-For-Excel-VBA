Attribute VB_Name = "mHandleManager"
' --- Standard Module: HandlerManager ---
Option Explicit
Public m_InstanceMap As Dictionary

' Register: Map a class instance using the handler's memory address as the key
Public Sub RegisterInstance(ByVal pHandler As LongPtr, ByRef obj As Object)
    If m_InstanceMap Is Nothing Then Set m_InstanceMap = New Dictionary
    m_InstanceMap.Add CStr(pHandler), obj
End Sub

' Unregister: Call this when releasing the handler to prevent memory leaks
Public Sub UnregisterInstance(ByVal pHandler As LongPtr)
    If Not m_InstanceMap Is Nothing Then
        If m_InstanceMap.Exists(CStr(pHandler)) Then m_InstanceMap.Remove CStr(pHandler)
    End If
End Sub

' Lookup: Used inside the Invoke method to retrieve the instance from the pointer
Public Function GetInstance(ByVal pHandler As LongPtr) As Object
    If Not m_InstanceMap Is Nothing Then
        If m_InstanceMap.Exists(CStr(pHandler)) Then
            Set GetInstance = m_InstanceMap(CStr(pHandler))
        End If
    End If
End Function
