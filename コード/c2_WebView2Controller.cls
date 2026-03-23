VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c2_WebView2Controller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: c2_WebView2Controller
Option Explicit

Public pController As LongPtr

Public WithEvents WebView2 As c3_WebView2
Attribute WebView2.VB_VarHelpID = -1
Public Event NavigationCompleted()
Public Event ScriptResultReceived(ByVal result As String)
Public Event WebView2ReadyCompleted()

' ---------------------------------------------------------
'  ICoreWebView2Controller::get_IsVisible (VTable Index 3)
' ---------------------------------------------------------
Public Property Get IsVisible() As Boolean
    Dim Value As Long ' BOOL variable to receive output
    Dim hr As Long    ' DispCallFunc HRESULT
    Dim res As Variant ' Method HRESULT
    
    Dim vArgs(0) As Variant: vArgs(0) = VarPtr(Value)
    Dim vTypes(0) As Integer: vTypes(0) = vbLongPtr
    Dim vPtrs(0) As LongPtr: vPtrs(0) = VarPtr(vArgs(0))
    
    hr = DispCallFunc(pController, 3 * LenB(pController), CC_STDCALL, vbLong, 1, vTypes(0), vPtrs(0), res)
    
    If hr = 0 Then
        If res = 0 Then
            IsVisible = (Value <> 0)
        Else
            Debug.Print "Failed to get IsVisible property. res: " & res
        End If
    Else
        Debug.Print "DispCallFunc failed during get_IsVisible. hr: " & hr
    End If
End Property

' ---------------------------------------------------------
'  ICoreWebView2Controller::put_IsVisible (VTable Index 4)
' ---------------------------------------------------------
Public Property Let IsVisible(ByVal Value As Boolean)
    Dim hr As Long
    Dim res As Variant

    Dim vArgs(0) As Variant
    vArgs(0) = CLng(IIf(Value, 1, 0))
    Dim vTypes(0) As Integer: vTypes(0) = vbLong
    Dim vPtrs(0) As LongPtr: vPtrs(0) = VarPtr(vArgs(0))
    
    hr = DispCallFunc(pController, 4 * LenB(pController), CC_STDCALL, vbLong, 1, vTypes(0), vPtrs(0), res)
    
    If hr <> 0 Or res <> 0 Then
        Debug.Print "Failed put_IsVisible. hr: " & hr & " res: " & res
    End If
End Property

' --- Architectural Notes for put_IsVisible ---
' This method acts as a drawing toggle between WebView2 and the parent window (e.g., UserForm).
'
' 1. Rendering Process Ignition
'    Even when initialized via CreateCoreWebView2Controller, the internal rendering process
'    doesn't run at full capacity. Setting it to True (1) kicks off pixel transfer to the parent HWND.
'
' 2. Parent-Child Relationship Enforcement
'    Setting it to True finalizes the Win32 behavior as a proper child window.
'    - True: Redraws its client area automatically as the parent redraws.
'    - False: Suspends visual rendering to save CPU/GPU (JavaScript keeps running in background).
'
' 3. Implementation Order
'    Typically, you call put_Bounds (sizing) first, then put_IsVisible(True).
'    If size is 0x0, making it Visible yields a blank screen.
'
' 4. Focus & Input
'    Keyboard and mouse events are ignored unless IsVisible is True.

' ---------------------------------------------------------
'  ICoreWebView2Controller::get_CoreWebView2 (VTable Index 25)
' ---------------------------------------------------------
Public Function GetWebView2() As c3_WebView2
    Dim ppWebView2 As LongPtr
    Dim args(0) As Variant
    Dim argTypes(0) As Integer
    Dim argPtrs(0) As LongPtr
    
    Dim hr As Long
    Dim result As Variant

    args(0) = VarPtr(ppWebView2)
    argTypes(0) = vbLongPtr
    argPtrs(0) = VarPtr(args(0))
 
    hr = DispCallFunc(pController, 25 * LenB(pController), CC_STDCALL, vbLong, 1, argTypes(0), argPtrs(0), result)
    
    If hr = 0 And result = 0 Then
        Dim WV2 As New c3_WebView2
        WV2.ppWebView2 = ppWebView2
        Set WebView2 = WV2
    Else
        Debug.Print "get_CoreWebView2 failed. hr: " & hr & " result: " & result
    End If
End Function

Public Sub ReadyCompleted()
    RaiseEvent WebView2ReadyCompleted
End Sub

' ---------------------------------------------------------
'  ICoreWebView2Controller::Close (VTable Index 24)
' ---------------------------------------------------------
Public Sub CloseWebView2()
    Dim hr As Long
    Dim result As Variant

    If pController = 0 Then Exit Sub

    hr = DispCallFunc(pController, 24 * LenB(pController), CC_STDCALL, vbLong, 0, 0, 0, result)

    If hr = 0 Then
        Debug.Print "WebView2 Controller Closed."
    Else
        Debug.Print "Close Failed. hr: " & hr
    End If

    Set WebView2 = Nothing
    pController = 0
End Sub

Private Sub WebView2_NavigationCompleted()
    RaiseEvent NavigationCompleted
End Sub

Private Sub WebView2_ReceiveScriptResult(ByVal result As String)
    RaiseEvent ScriptResultReceived(result)
End Sub
