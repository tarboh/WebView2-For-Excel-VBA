Attribute VB_Name = "mWebView2Api"
'Module mWebView2Api

Option Explicit

' ???????????????????????????????????????????????????????????????????????
' mWebView2Api
' WebView2 用の定数、列挙型、構造体定義
' ???????????????????????????????????????????????????????????????????????

#If Win64 Then
#Else
    #Error "このモジュールは 64 ビット VBA (x64) が必要です"
#End If

' ???????????????????????????????????????????????
' Section 1: GUID 構造体
' ???????????????????????????????????????????????

Public Type guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

' --- ハンドラエントリ ---
Public Type HandlerEntry
    Key As String
    vtblBlock As LongPtr
    fakeObj As LongPtr
    isEvent As Boolean
    pInterface As LongPtr
    removeIndex As Long
    token As LongPtr
    pExtraRelease As LongPtr
    active As Boolean
End Type

' ???????????????????????????????????????????????
' Section 2: TYPEATTR 構造体 (ITypeInfo 用)
' ???????????????????????????????????????????????

' x64 レイアウト (合計 96 bytes)
' BuildFuncPtrCache で cFuncs フィールドを参照する。

Public Type TYPEATTR
    guid As guid           ' 0-15   (16 bytes)
    lcid As Long           ' 16-19
    dwReserved As Long     ' 20-23
    memidConstructor As Long ' 24-27
    memidDestructor As Long  ' 28-31
    lpstrSchema As LongPtr ' 32-39
    cbSizeInstance As Long  ' 40-43
    typekind As Long        ' 44-47
    cFuncs As Integer       ' 48-49  ★ メソッド数
    cVars As Integer        ' 50-51
    cImplTypes As Integer   ' 52-53
    cbSizeVft As Integer    ' 54-55
    cbAlignment As Integer  ' 56-57
    wTypeFlags As Integer   ' 58-59
    wMajorVerNum As Integer ' 60-61
    wMinorVerNum As Integer ' 62-63
    tdescAlias As LongPtr   ' 64-71  (TYPEDESC)
    tdescAlias2 As LongPtr  ' 72-79
    idldescType As LongPtr  ' 80-87  (IDLDESC)
    idldescType2 As LongPtr ' 88-95
End Type


' ???????????????????????????????????????????????
' Section 3: FUNCDESC 構造体 (ITypeInfo 用)
' ???????????????????????????????????????????????

' x64 レイアウト
' BuildFuncPtrCache で memid と oVft を参照する。

Public Type FUNCDESC
    memid As Long              ' 0-3    MEMBERID
    padding1 As Long           ' 4-7    (alignment)
    lprgscode As LongPtr       ' 8-15
    lprgelemdescParam As LongPtr ' 16-23
    funckind As Long           ' 24-27
    invkind As Long            ' 28-31
    callconv As Long           ' 32-35
    cParams As Integer         ' 36-37
    cParamsOpt As Integer      ' 38-39
    oVft As Integer            ' 40-41  ★ VTable オフセット
    cScodes As Integer         ' 42-43
    elemdescFunc_tdesc As LongPtr  ' 44-51  (ELEMDESC.TYPEDESC)
    elemdescFunc_tdesc2 As LongPtr ' 52-59
    elemdescFunc_idldesc As LongPtr ' 60-67  (ELEMDESC.IDLDESC)
    elemdescFunc_idldesc2 As LongPtr ' 68-75
    wFuncFlags As Integer      ' 76-77
    padding2 As Integer        ' 78-79  (alignment)
    padding3 As Long           ' 80-83
End Type


' ???????????????????????????????????????????????
' Section 4: WebView2 列挙型
' ???????????????????????????????????????????????

' --- COREWEBVIEW2_CAPTURE_PREVIEW_IMAGE_FORMAT ---
Public Enum COREWEBVIEW2_CAPTURE_PREVIEW_IMAGE_FORMAT
    COREWEBVIEW2_CAPTURE_PREVIEW_IMAGE_FORMAT_PNG = 0
    COREWEBVIEW2_CAPTURE_PREVIEW_IMAGE_FORMAT_JPEG = 1
End Enum

' --- COREWEBVIEW2_MOVE_FOCUS_REASON ---
Public Enum COREWEBVIEW2_MOVE_FOCUS_REASON
    COREWEBVIEW2_MOVE_FOCUS_REASON_PROGRAMMATIC = 0
    COREWEBVIEW2_MOVE_FOCUS_REASON_NEXT = 1
    COREWEBVIEW2_MOVE_FOCUS_REASON_PREVIOUS = 2
End Enum

' --- COREWEBVIEW2_WEB_RESOURCE_CONTEXT ---
Public Enum COREWEBVIEW2_WEB_RESOURCE_CONTEXT
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_ALL = 0
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_DOCUMENT = 1
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_STYLESHEET = 2
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_IMAGE = 3
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_MEDIA = 4
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_FONT = 5
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_SCRIPT = 6
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_XML_HTTP_REQUEST = 7
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_FETCH = 8
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_TEXT_TRACK = 9
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_EVENT_SOURCE = 10
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_WEBSOCKET = 11
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_MANIFEST = 12
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_SIGNED_EXCHANGE = 13
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_PING = 14
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_CSP_VIOLATION_REPORT = 15
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_OTHER = 16
End Enum

' --- COREWEBVIEW2_KEY_EVENT_KIND ---
Public Enum COREWEBVIEW2_KEY_EVENT_KIND
    COREWEBVIEW2_KEY_EVENT_KIND_KEY_DOWN = 0
    COREWEBVIEW2_KEY_EVENT_KIND_KEY_UP = 1
    COREWEBVIEW2_KEY_EVENT_KIND_SYSTEM_KEY_DOWN = 2
    COREWEBVIEW2_KEY_EVENT_KIND_SYSTEM_KEY_UP = 3
End Enum

' --- COREWEBVIEW2_PERMISSION_KIND ---
Public Enum COREWEBVIEW2_PERMISSION_KIND
    COREWEBVIEW2_PERMISSION_KIND_UNKNOWN_PERMISSION = 0
    COREWEBVIEW2_PERMISSION_KIND_MICROPHONE = 1
    COREWEBVIEW2_PERMISSION_KIND_CAMERA = 2
    COREWEBVIEW2_PERMISSION_KIND_GEOLOCATION = 3
    COREWEBVIEW2_PERMISSION_KIND_NOTIFICATIONS = 4
    COREWEBVIEW2_PERMISSION_KIND_OTHER_SENSORS = 5
    COREWEBVIEW2_PERMISSION_KIND_CLIPBOARD_READ = 6
End Enum

' --- COREWEBVIEW2_PERMISSION_STATE ---
Public Enum COREWEBVIEW2_PERMISSION_STATE
    COREWEBVIEW2_PERMISSION_STATE_DEFAULT = 0
    COREWEBVIEW2_PERMISSION_STATE_ALLOW = 1
    COREWEBVIEW2_PERMISSION_STATE_DENY = 2
End Enum

' --- COREWEBVIEW2_PROCESS_FAILED_KIND ---
Public Enum COREWEBVIEW2_PROCESS_FAILED_KIND
    COREWEBVIEW2_PROCESS_FAILED_KIND_BROWSER_PROCESS_EXITED = 0
    COREWEBVIEW2_PROCESS_FAILED_KIND_RENDER_PROCESS_EXITED = 1
    COREWEBVIEW2_PROCESS_FAILED_KIND_RENDER_PROCESS_UNRESPONSIVE = 2
End Enum

' --- COREWEBVIEW2_SCRIPT_DIALOG_KIND ---
Public Enum COREWEBVIEW2_SCRIPT_DIALOG_KIND
    COREWEBVIEW2_SCRIPT_DIALOG_KIND_ALERT = 0
    COREWEBVIEW2_SCRIPT_DIALOG_KIND_CONFIRM = 1
    COREWEBVIEW2_SCRIPT_DIALOG_KIND_PROMPT = 2
    COREWEBVIEW2_SCRIPT_DIALOG_KIND_BEFOREUNLOAD = 3
End Enum

' --- COREWEBVIEW2_WEB_ERROR_STATUS ---
Public Enum COREWEBVIEW2_WEB_ERROR_STATUS
    COREWEBVIEW2_WEB_ERROR_STATUS_UNKNOWN = 0
    COREWEBVIEW2_WEB_ERROR_STATUS_CERTIFICATE_COMMON_NAME_IS_INCORRECT = 1
    COREWEBVIEW2_WEB_ERROR_STATUS_CERTIFICATE_EXPIRED = 2
    COREWEBVIEW2_WEB_ERROR_STATUS_CLIENT_CERTIFICATE_CONTAINS_ERRORS = 3
    COREWEBVIEW2_WEB_ERROR_STATUS_CERTIFICATE_REVOKED = 4
    COREWEBVIEW2_WEB_ERROR_STATUS_CERTIFICATE_IS_INVALID = 5
    COREWEBVIEW2_WEB_ERROR_STATUS_SERVER_UNREACHABLE = 6
    COREWEBVIEW2_WEB_ERROR_STATUS_TIMEOUT = 7
    COREWEBVIEW2_WEB_ERROR_STATUS_ERROR_HTTP_INVALID_SERVER_RESPONSE = 8
    COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_ABORTED = 9
    COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_RESET = 10
    COREWEBVIEW2_WEB_ERROR_STATUS_DISCONNECTED = 11
    COREWEBVIEW2_WEB_ERROR_STATUS_CANNOT_CONNECT = 12
    COREWEBVIEW2_WEB_ERROR_STATUS_HOST_NAME_NOT_RESOLVED = 13
    COREWEBVIEW2_WEB_ERROR_STATUS_OPERATION_CANCELED = 14
    COREWEBVIEW2_WEB_ERROR_STATUS_REDIRECT_FAILED = 15
    COREWEBVIEW2_WEB_ERROR_STATUS_UNEXPECTED_ERROR = 16
End Enum

' --- COREWEBVIEW2_PHYSICAL_KEY_STATUS ---
' AcceleratorKeyPressed イベントの args から取得可能な構造体
Public Type COREWEBVIEW2_PHYSICAL_KEY_STATUS
    RepeatCount As Long        ' 0-3
    ScanCode As Long           ' 4-7
    IsExtendedKey As Long      ' 8-11   (BOOL)
    IsMenuKeyDown As Long      ' 12-15  (BOOL)
    WasKeyDown As Long         ' 16-19  (BOOL)
    IsKeyReleased As Long      ' 20-23  (BOOL)
End Type


' ???????????????????????????????????????????????
' Section 5: HRESULT 定数
' ???????????????????????????????????????????????

Public Const S_OK As Long = 0
Public Const S_FALSE As Long = 1
Public Const E_FAIL As Long = &H80004005
Public Const E_POINTER As Long = &H80004003
Public Const E_NOINTERFACE As Long = &H80004002
Public Const E_INVALIDARG As Long = &H80070057
Public Const E_PENDING As Long = &H8000000A
Public Const E_ABORT As Long = &H80004004
Public Const E_ACCESSDENIED As Long = &H80070005
Public Const E_OUTOFMEMORY As Long = &H8007000E
Public Const E_UNEXPECTED As Long = &H8000FFFF
Public Const E_NOTIMPL As Long = &H80004001
Public Const CLASS_E_NOAGGREGATION As Long = &H80040110

' WebView2 固有の HRESULT
Public Const HRESULT_FROM_WIN32_ERROR_FILE_NOT_FOUND As Long = &H80070002
Public Const HRESULT_FROM_WIN32_ERROR_FILE_EXISTS As Long = &H80070050


' ???????????????????????????????????????????????
' Section 6: その他の定数
' ???????????????????????????????????????????????

' VarType 拡張
Public Const vbLongPtr As Integer = 20  ' VT_I8 on x64

' STGM フラグ (IStream / SHCreateStreamOnFileW 用)
Public Const STGM_READ As Long = &H0
Public Const STGM_WRITE As Long = &H1
Public Const STGM_READWRITE As Long = &H2
Public Const STGM_CREATE As Long = &H1000
Public Const STGM_SHARE_DENY_NONE As Long = &H40
Public Const STGM_SHARE_DENY_READ As Long = &H30
Public Const STGM_SHARE_DENY_WRITE As Long = &H20
Public Const STGM_SHARE_EXCLUSIVE As Long = &H10

' CryptStringToBinary フラグ
Public Const CRYPT_STRING_BASE64 As Long = 1
Public Const CRYPT_STRING_HEX As Long = 4

' VirtualAlloc / VirtualFree
Public Const MEM_COMMIT As Long = &H1000
Public Const MEM_RESERVE As Long = &H2000
Public Const MEM_RELEASE As Long = &H8000
Public Const PAGE_EXECUTE_READWRITE As Long = &H40
Public Const PAGE_READWRITE As Long = &H4

' GetDeviceCaps
Public Const LOGPIXELSX As Long = 88
Public Const LOGPIXELSY As Long = 90

' DispCallFunc
Public Const CC_STDCALL As Long = 4
Public Const CC_CDECL As Long = 1

