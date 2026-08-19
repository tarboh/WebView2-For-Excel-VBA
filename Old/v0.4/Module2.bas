Attribute VB_Name = "Module2"
''''''''''''''''''''''''''''''''''
' --- Module2.bas 第二段階 ---
''''''''''''''''''''''''''''''''''
Option Explicit

#If x64 Then
    Private Const NullPtr As LongLong = 0^
    Private Const PtrSize = 8
#Else
    Private Const NullPtr As Long = 0&
    Private Const PtrSize = 4
#End If

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

Private Const S_OK As Long = 0

' WebView2Loader.dll entry point (Called from c0.CreateWebView2Environment)
Private Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" ( _
    ByVal browserExecutableFolder As LongPtr, _
    ByVal userDataFolder As LongPtr, _
    ByVal additionalBrowserArguments As LongPtr, _
    ByVal environmentCreatedHandler As LongPtr) As Long

' --- vTable保持用変数 ---
' PointerAccessorによって、これら自体がCOMオブジェクトのメモリ実体として機能する
Private Type WebView2Handler
    pVTable As LongPtr     ' オブジェクトの先頭（vTableへのポインタ）
    Functions(3) As LongPtr ' vTableの実体（関数ポインタの配列）
End Type
Private m_Handler As WebView2Handler

' Helper function to receive AddressOf as LongPtr
Private Function GetAddr(ByVal addr As LongPtr) As LongPtr
    GetAddr = addr
End Function

' IUnknown::QueryInterface
Private Function Handler_QueryInterface(ByVal this As LongPtr, ByVal riid As LongPtr, ByRef ppvObject As LongPtr) As Long
    ' Normally used to check GUID, but for now it returns itself
    ppvObject = this
    Handler_QueryInterface = S_OK
End Function

' IUnknown::AddRef / Release (Returns 1 as a stub/dummy)
Private Function Handler_AddRef(ByVal this As LongPtr) As Long:
    Handler_AddRef = 1
End Function
Private Function Handler_Release(ByVal this As LongPtr) As Long
    Handler_Release = 1
End Function

' ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler::Invoke
' Receives the initialization result from WebView2
Private Function Handler_Invoke(ByVal this As LongPtr, ByVal errorCode As Long, ByVal pEnvironment As LongPtr) As Long
    'まだ何もしない
    Debug.Print "ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler::Invoke is called."
    Handler_Invoke = 0
End Function

' Initializes the WebView2 Environment
Private Sub CreateWebView2Environment()

    Static pa As PointerAccessor
    
    ' 1. PointerAccessorのセットアップ（pa.arr を pa.sa に紐付け）
    If pa.sa.cDims = 0 Then
        pa.sa.cDims = 1
        pa.sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        pa.sa.cbElements = PtrSize
        pa.sa.cLocks = 1
        MemLongPtr(VarPtr(pa)) = VarPtr(pa.sa)
    End If

    ' 2. m_Handler 内に vTable を構築
    ' vTable自体を m_Handler.Functions に配置し、pVTableを自分自身のFunctions配列に向ける
    m_Handler.pVTable = VarPtr(m_Handler.Functions(0))
    
    ' pa.arr を使って vTable の内容（関数ポインタ）を書き込む
    ' pa.sa.pvData を書き込み先に向け、arr(n) 経由で流し込む
    pa.sa.pvData = VarPtr(m_Handler.Functions(0))
    pa.sa.rgsabound0.cElements = 4
    
    pa.arr(0) = GetAddr(AddressOf Handler_QueryInterface)
    pa.arr(1) = GetAddr(AddressOf Handler_AddRef)
    pa.arr(2) = GetAddr(AddressOf Handler_Release)
    pa.arr(3) = GetAddr(AddressOf Handler_Invoke)
    
    ' 3. 後片付け（配列の安全な解除）
    pa.sa.rgsabound0.cElements = 0
    pa.sa.pvData = 0

    ' 4. WebView2環境作成の呼び出し
    ' pObject として、m_Handler 自体のアドレス（最初の要素は vTableへのポインタ）を渡す
    Dim userDataPath As String: userDataPath = "C:\Temp\VBA_WebView2"
    If Dir("C:\Temp", vbDirectory) = "" Then MkDir "C:\Temp"

    Dim hr As Long
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(userDataPath), 0, VarPtr(m_Handler))

    If hr <> S_OK Then
        MsgBox "Failed to initialize WebView2. HRESULT: " & hr, vbCritical
    End If

    DoEvents

End Sub

Private Property Let MemLongPtr(ByVal addr As LongPtr, ByVal newValue As LongPtr)
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        .arr(0) = newValue
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Property

'https://github.com/WNKLER/RefTypes/discussions/3#discussion-8595790
Private Sub WritePtrNatively(ByRef ptrs() As LONG_PTR, ByVal ptr As LongPtr)
    ptrs(0) = ptr
End Sub
