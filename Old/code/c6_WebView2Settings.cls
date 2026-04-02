VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c6_WebView2Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Class c6_WebView2Settings
Option Explicit

Private m_pSettings As LongPtr

' Initialize the class by passing the object pointer
Public Sub Init(ByVal ptr As LongPtr)
    m_pSettings = ptr
End Sub

Public Property Get Pointer() As LongPtr
    Pointer = m_pSettings
End Property

'MIDL_INTERFACE ("e562e4f0-d7fa-43ac-8d71-c05150499f00")
'ICoreWebView2Settings:      Public IUnknown
'{
'public:
'3
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_IsScriptEnabled(
'    /* [retval][out] */ BOOL *isScriptEnabled) = 0;
Public Property Get IsScriptEnabled() As Boolean
    IsScriptEnabled = GetBoolProperty(3)
End Property
'4
'virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_IsScriptEnabled(
'    /* [in] */ BOOL isScriptEnabled) = 0;
Public Property Let IsScriptEnabled(ByVal Value As Boolean)
    LetBoolProperty 4, Value
End Property
'5
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_IsWebMessageEnabled(
'    /* [retval][out] */ BOOL *isWebMessageEnabled) = 0;
Public Property Get IsWebMessageEnabled() As Boolean
    IsWebMessageEnabled = GetBoolProperty(5)
End Property
'6
'virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_IsWebMessageEnabled(
'    /* [in] */ BOOL isWebMessageEnabled) = 0;
Public Property Let IsWebMessageEnabled(ByVal Value As Boolean)
    LetBoolProperty 6, Value
End Property
'7
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_AreDefaultScriptDialogsEnabled(
'    /* [retval][out] */ BOOL *areDefaultScriptDialogsEnabled) = 0;
Public Property Get AreDefaultScriptDialogsEnabled() As Boolean
    AreDefaultScriptDialogsEnabled = GetBoolProperty(7)
End Property
'8
'virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_AreDefaultScriptDialogsEnabled(
'    /* [in] */ BOOL areDefaultScriptDialogsEnabled) = 0;
Public Property Let AreDefaultScriptDialogsEnabled(ByVal Value As Boolean)
    LetBoolProperty 8, Value
End Property
'9
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_IsStatusBarEnabled(
'    /* [retval][out] */ BOOL *isStatusBarEnabled) = 0;
Public Property Get IsStatusBarEnabled() As Boolean
    IsStatusBarEnabled = GetBoolProperty(9)
End Property
'10
'virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_IsStatusBarEnabled(
'    /* [in] */ BOOL isStatusBarEnabled) = 0;
Public Property Let IsStatusBarEnabled(ByVal Value As Boolean)
    LetBoolProperty 10, Value
End Property
'11
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_AreDevToolsEnabled(
'    /* [retval][out] */ BOOL *areDevToolsEnabled) = 0;
Public Property Get AreDevToolsEnabled() As Boolean
    AreDevToolsEnabled = GetBoolProperty(11)
End Property
'12
'virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_AreDevToolsEnabled(
'    /* [in] */ BOOL areDevToolsEnabled) = 0;
Public Property Let AreDevToolsEnabled(ByVal Value As Boolean)
    LetBoolProperty 12, Value
End Property
'13
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_AreDefaultContextMenusEnabled(
'    /* [retval][out] */ BOOL *enabled) = 0;
'
'14
'virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_AreDefaultContextMenusEnabled(
'    /* [in] */ BOOL enabled) = 0;
'
'15
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_AreHostObjectsAllowed(
'    /* [retval][out] */ BOOL *allowed) = 0;
'
'16
'virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_AreHostObjectsAllowed(
'    /* [in] */ BOOL allowed) = 0;
'
'17
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_IsZoomControlEnabled(
'    /* [retval][out] */ BOOL *enabled) = 0;
'
'18
'virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_IsZoomControlEnabled(
'    /* [in] */ BOOL enabled) = 0;
'
'19
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_IsBuiltInErrorPageEnabled(
'    /* [retval][out] */ BOOL *enabled) = 0;
'
'20
'virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_IsBuiltInErrorPageEnabled(
'    /* [in] */ BOOL enabled) = 0;
'
'};



' ==========================================================
'  Internal utility for calling DispCallFunc
' ==========================================================

Private Function GetBoolProperty(ByVal Index As Long) As Boolean
    If m_pSettings = 0 Then Exit Function
    
    Dim hr As Long, res As Variant
    Dim outVal As Long ' C++ BOOL is 4 bytes
    
    Dim types(0) As Integer: types(0) = vbLongPtr
    Dim args(0) As Variant:  args(0) = VarPtr(outVal)
    Dim ptrs(0) As LongPtr: ptrs(0) = VarPtr(args(0))
    
    hr = DispCallFunc(m_pSettings, Index * LenB(m_pSettings), CC_STDCALL, vbLong, 1, types(0), ptrs(0), res)
    
    If hr = 0 Then
        GetBoolProperty = (outVal <> 0)
    End If
End Function

Private Sub LetBoolProperty(ByVal Index As Long, ByVal Value As Boolean)
    If m_pSettings = 0 Then Exit Sub
    
    Dim hr As Long, res As Variant
    Dim inVal As Long
    inVal = IIf(Value, 1, 0) ' C++ TRUE is 1 (VBA True is -1)
    
    Dim types(0) As Integer: types(0) = vbLong
    Dim args(0) As Variant:  args(0) = inVal
    Dim ptrs(0) As LongPtr: ptrs(0) = VarPtr(args(0))
    
    hr = DispCallFunc(m_pSettings, Index * LenB(m_pSettings), CC_STDCALL, vbLong, 1, types(0), ptrs(0), res)
    If hr = 0 And res = 0 Then
        Debug.Print "LetBoolProperty Success."
    End If
End Sub
