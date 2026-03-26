VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c7_DTPEventReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Class c7_DTPEventReceiver

Option Explicit

Public m_Pointer As LongPtr
Private m_eventName As String

Public Property Get eventName() As String
    eventName = m_Name
End Property

'GetDevToolsProtocolEventReceiver
Public Function SetupEventReceiver(ppWebView2 As LongPtr, ByVal eventName As String) As Long
    m_Pointer = 0
    Dim hr As Long
    hr = dcf(ppWebView2, _
            42, _
            "GetDevToolsProtocolEventReceiver(" & eventName & ")", _
            StrPtr(eventName), _
            VarPtr(m_Pointer))
    If hr = 0 Then
        m_eventName = eventName
    End If
    SetupEventReceiver = hr
End Function

'public HRESULT add_DevToolsProtocolEventReceived(
'   ICoreWebView2DevToolsProtocolEventReceivedEventHandler * handler,
'   EventRegistrationToken * token)
Public Function add_DevToolsProtocolEventReceived()
'The Invoke method of the handler runs whenever the corresponding DevToolsProtocol event runs.
'Invoke runs with an event args object containing the parameter object of the DevTools Protocol event as a JSON string.
End Function

'public HRESULT remove_DevToolsProtocolEventReceived(
'   EventRegistrationToken token)
Public Function remove_DevToolsProtocolEventReceived()
'Remove an event handler previously added with add_DevToolsProtocolEventReceived.
End Function

