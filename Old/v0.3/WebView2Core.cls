VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebView2Core"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' WebView2Core.cls  (Handler管理集約版)
' =============================================================================
' 低レベル COM 操作基盤 + 統一ハンドラ管理。
'
' ハンドラ管理API (Tab/Manager から呼ぶ):
'   RegisterEvent()      - COM イベント購読 (add_/remove_ ペア)
'   RegisterEventEx()    - 同上 + pExtraRelease (DTP用)
'   EnsureCallback()     - 完了コールバック (再利用)
'   CreateOneShot()      - 使い捨てハンドラ (CDP用)
'   FreeOneShot()        - 使い捨て解放
'   UnregisterEvent()    - 個別解除
'   UnregisterByPrefix() - プレフィックス一括解除
'   LookupKeyByFakeObj() - fakeObj → fullKey 逆引き
'
' キー名前空間:
'   Tab  → "T{id}.WV.NavigationStarting" 等
'   Mgr  → "MGR.__env_completed" 等
' =============================================================================

#If Win64 Then
#Else
    #Error "このクラスは 64 ビット VBA (x64) が必要です"
#End If

' --- Win32 API ---
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" ( _
    ByVal dest As LongPtr, ByVal src As LongPtr, ByVal Length As LongPtr)
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef dest As Any, ByRef src As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, _
    ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
Private Declare PtrSafe Function VirtualFree Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, _
    ByVal dwFreeType As Long) As Long
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" ( _
    ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, _
    ByVal cc As Long, ByVal vtReturn As Integer, _
    ByVal cActuals As Long, ByRef prgvt As Any, _
    ByRef prgpvarg As Any, ByRef pvargResult As Any) As Long
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal pv As LongPtr)
Private Declare PtrSafe Sub SysFreeString Lib "oleaut32" (ByVal bstr As LongPtr)
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" ( _
    ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" ( _
    ByVal hdc As LongPtr, ByVal nIndex As Long) As Long

Private Const MEM_COMMIT As Long = &H1000
Private Const MEM_RESERVE As Long = &H2000
Private Const MEM_RELEASE As Long = &H8000
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private Const CC_STDCALL As Long = 4
Private Const E_NOINTERFACE As Long = &H80004002
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90


' =============================================================================
' ハンドラエントリ (Private)
' =============================================================================

Private Type HandlerEntry
    fullKey As String
    vtblBlock As LongPtr
    fakeObj As LongPtr
    isEvent As Boolean
    pInterface As LongPtr
    removeIndex As Long
    token As LongPtr
    pExtraRelease As LongPtr
    active As Boolean
End Type


' =============================================================================
' メンバ変数
' =============================================================================

Private m_pStubQI As LongPtr
Private m_pStubAR As LongPtr

Private m_handlers As Object        ' Dictionary: fullKey -> index
Private m_fakeObjToKey As Object    ' Dictionary: CStr(fakeObj) -> fullKey
Private m_hData() As HandlerEntry
Private m_hCount As Long


' =============================================================================
' ライフサイクル
' =============================================================================

Private Sub Class_Initialize()
    m_pStubQI = MakeStub_QI()
    m_pStubAR = MakeStub_AR()
    Set m_handlers = CreateObject("Scripting.Dictionary")
    m_handlers.CompareMode = 1
    Set m_fakeObjToKey = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    UnregisterAll
    If m_pStubQI <> 0 Then VirtualFree m_pStubQI, 0, MEM_RELEASE: m_pStubQI = 0
    If m_pStubAR <> 0 Then VirtualFree m_pStubAR, 0, MEM_RELEASE: m_pStubAR = 0
    Set m_handlers = Nothing
    Set m_fakeObjToKey = Nothing
End Sub


' =============================================================================
' スタブ生成
' =============================================================================

Private Function MakeStub_QI() As LongPtr
    Dim a As LongPtr
    a = VirtualAlloc(0, 16, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If a = 0 Then Exit Function
    Dim b(7) As Byte, hr As Long: hr = E_NOINTERFACE
    b(0) = &HB8: CopyMemory b(1), hr, 4: b(5) = &HC3
    RtlMoveMemory a, VarPtr(b(0)), 6
    MakeStub_QI = a
End Function

Private Function MakeStub_AR() As LongPtr
    Dim a As LongPtr
    a = VirtualAlloc(0, 16, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If a = 0 Then Exit Function
    Dim b(7) As Byte
    b(0) = &HB8: b(1) = 1: b(2) = 0: b(3) = 0: b(4) = 0: b(5) = &HC3
    RtlMoveMemory a, VarPtr(b(0)), 6
    MakeStub_AR = a
End Function


' =============================================================================
' 関数ポインタキャッシュ構築
' =============================================================================

Public Sub BuildFuncPtrCache(ByVal pObj As LongPtr, ByVal dict As Object)
    dict.RemoveAll
    Dim pTI As LongPtr, hr As Long
    hr = CallCOM3(pObj, 4, 0&, 0&, VarPtr(pTI))
    If hr <> 0 Then Exit Sub

    Dim pTA As LongPtr
    hr = CallCOM1(pTI, 3, VarPtr(pTA))
    If hr <> 0 Then GoTo Cleanup

    Dim uTA As TYPEATTR: CopyMemory uTA, ByVal pTA, LenB(uTA)
    Dim pVT As LongPtr:  CopyMemory pVT, ByVal pObj, LenB(pVT)

    Dim i As Long
    For i = 0 To uTA.cFuncs - 1
        Dim pFD As LongPtr
        hr = CallCOM2(pTI, 5, CLng(i), VarPtr(pFD))
        If hr <> 0 Then GoTo NextF
        Dim uFD As FUNCDESC: CopyMemory uFD, ByVal pFD, LenB(uFD)
        Dim pB As LongPtr: pB = 0
        hr = CallCOM5(pTI, 12, uFD.memid, VarPtr(pB), 0, 0, 0)
        If hr = 0 And pB <> 0 Then
            Dim fn As String: fn = PtrToStrW(pB)
            SysFreeString pB: pB = 0
            Dim pA As LongPtr: CopyMemory pA, ByVal (pVT + uFD.oVft), LenB(pA)
            dict(fn) = pA
        End If
        CallCOM1V pTI, 20, pFD
NextF:
    Next i
    CallCOM1V pTI, 19, pTA
Cleanup:
    CallCOM0V pTI, 2
End Sub


' =============================================================================
' VTable + サンク生成 (内部)
' =============================================================================

Private Function MakeVTBlock(ByVal pOwner As LongPtr, ByVal pFunc As LongPtr) As LongPtr
    If pFunc = 0 Then Exit Function
    Dim blk As LongPtr
    blk = VirtualAlloc(0, 80, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If blk = 0 Then Exit Function

    Dim vtbl As LongPtr:  vtbl = blk
    Dim thk As LongPtr: thk = blk + 32

    Dim b(47) As Byte, n As Long
    b(n) = &H4D: b(n + 1) = &H89: b(n + 2) = &HC1: n = n + 3   ' mov r9,r8
    b(n) = &H49: b(n + 1) = &H89: b(n + 2) = &HD0: n = n + 3   ' mov r8,rdx
    b(n) = &H48: b(n + 1) = &H89: b(n + 2) = &HCA: n = n + 3   ' mov rdx,rcx
    b(n) = &H48: b(n + 1) = &HB9: n = n + 2                     ' mov rcx,imm64
    CopyMemory b(n), pOwner, 8: n = n + 8
    b(n) = &H48: b(n + 1) = &HB8: n = n + 2                     ' mov rax,imm64
    CopyMemory b(n), pFunc, 8: n = n + 8
    b(n) = &HFF: b(n + 1) = &HE0: n = n + 2                     ' jmp rax
    RtlMoveMemory thk, VarPtr(b(0)), CLngPtr(n)

    RtlMoveMemory vtbl, VarPtr(m_pStubQI), 8
    RtlMoveMemory vtbl + 8, VarPtr(m_pStubAR), 8
    RtlMoveMemory vtbl + 16, VarPtr(m_pStubAR), 8
    RtlMoveMemory vtbl + 24, VarPtr(thk), 8

    MakeVTBlock = vtbl
End Function

Private Function MakeFakeObj(ByVal vtbl As LongPtr) As LongPtr
    Dim p As LongPtr
    p = VirtualAlloc(0, 8, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If p = 0 Then Exit Function
    RtlMoveMemory p, VarPtr(vtbl), 8
    MakeFakeObj = p
End Function

Private Sub FreeBlock(ByRef vtbl As LongPtr, ByRef fo As LongPtr)
    If fo <> 0 Then VirtualFree fo, 0, MEM_RELEASE: fo = 0
    If vtbl <> 0 Then VirtualFree vtbl, 0, MEM_RELEASE: vtbl = 0
End Sub


' =============================================================================
' 統一ハンドラ管理 API
' =============================================================================

' WebView2Core に追加
Public Function AddRefCOM(ByVal pInterface As LongPtr) As Long
    If pInterface = 0 Then Exit Function
    Dim res As Variant
    DispCallFunc pInterface, 1 * LenB(pInterface), CC_STDCALL, vbLong, 0, 0, 0, res
    AddRefCOM = CLng(res)
End Function

Public Function ReleaseCOM(ByVal pInterface As LongPtr) As Long
    If pInterface = 0 Then Exit Function
    Dim res As Variant
    DispCallFunc pInterface, 2 * LenB(pInterface), CC_STDCALL, vbLong, 0, 0, 0, res
    ReleaseCOM = CLng(res)
End Function

Public Function RegisterEvent( _
    ByVal prefix As String, _
    ByVal pOwner As LongPtr, _
    ByVal pInterface As LongPtr, _
    ByVal addIndex As Long, _
    ByVal removeIndex As Long, _
    ByVal callbackMethod As String, _
    ByVal localKey As String, _
    ByVal funcPtrs As Object) As String

    If pInterface = 0 Then Exit Function
    Dim fk As String: fk = prefix & localKey
    If m_handlers.Exists(fk) Then Exit Function
    If Not funcPtrs.Exists(callbackMethod) Then
        Debug.Print "RegisterEvent: not found - " & callbackMethod
        Exit Function
    End If

    Dim pFunc As LongPtr: pFunc = funcPtrs(callbackMethod)
    Dim vtbl As LongPtr: vtbl = MakeVTBlock(pOwner, pFunc)
    If vtbl = 0 Then Exit Function
    Dim fo As LongPtr: fo = MakeFakeObj(vtbl)
    If fo = 0 Then FreeBlock vtbl, fo: Exit Function

    Dim token As LongPtr, hr As Long
    hr = dcf(pInterface, addIndex, callbackMethod, fo, VarPtr(token))
    If hr <> 0 Then
        Debug.Print "RegisterEvent: add failed - " & fk & " hr=&H" & Hex(hr)
        FreeBlock vtbl, fo: Exit Function
    End If

    m_hCount = m_hCount + 1
    ReDim Preserve m_hData(0 To m_hCount - 1)
    With m_hData(m_hCount - 1)
        .fullKey = fk: .vtblBlock = vtbl: .fakeObj = fo
        .isEvent = True: .pInterface = pInterface
        .removeIndex = removeIndex: .token = token: .active = True
    End With
    m_handlers(fk) = m_hCount - 1
    m_fakeObjToKey(CStr(fo)) = fk
    RegisterEvent = fk
End Function

Public Function RegisterEventEx( _
    ByVal prefix As String, ByVal pOwner As LongPtr, _
    ByVal pInterface As LongPtr, ByVal addIndex As Long, _
    ByVal removeIndex As Long, ByVal callbackMethod As String, _
    ByVal localKey As String, ByVal funcPtrs As Object, _
    ByVal pExtraRelease As LongPtr) As String

    Dim fk As String
    fk = RegisterEvent(prefix, pOwner, pInterface, addIndex, removeIndex, _
                        callbackMethod, localKey, funcPtrs)
    If fk <> "" And pExtraRelease <> 0 Then
        m_hData(m_handlers(fk)).pExtraRelease = pExtraRelease
    End If
    RegisterEventEx = fk
End Function

Public Function EnsureCallback( _
    ByVal prefix As String, ByVal pOwner As LongPtr, _
    ByVal localKey As String, ByVal callbackMethod As String, _
    ByVal funcPtrs As Object) As LongPtr

    Dim fk As String: fk = prefix & localKey
    If m_handlers.Exists(fk) Then
        Dim ei As Long: ei = m_handlers(fk)
        If m_hData(ei).active Then
            EnsureCallback = m_hData(ei).fakeObj: Exit Function
        End If
    End If

    If Not funcPtrs.Exists(callbackMethod) Then Exit Function
    Dim pFunc As LongPtr: pFunc = funcPtrs(callbackMethod)
    Dim vtbl As LongPtr: vtbl = MakeVTBlock(pOwner, pFunc)
    If vtbl = 0 Then Exit Function
    Dim fo As LongPtr: fo = MakeFakeObj(vtbl)
    If fo = 0 Then FreeBlock vtbl, fo: Exit Function

    m_hCount = m_hCount + 1
    ReDim Preserve m_hData(0 To m_hCount - 1)
    With m_hData(m_hCount - 1)
        .fullKey = fk: .vtblBlock = vtbl: .fakeObj = fo
        .isEvent = False: .active = True
    End With
    m_handlers(fk) = m_hCount - 1
    m_fakeObjToKey(CStr(fo)) = fk
    EnsureCallback = fo
End Function

Public Function CreateOneShot( _
    ByVal pOwner As LongPtr, ByVal callbackMethod As String, _
    ByVal funcPtrs As Object, _
    ByRef outVtbl As LongPtr, ByRef outFO As LongPtr) As Boolean

    If Not funcPtrs.Exists(callbackMethod) Then Exit Function
    outVtbl = MakeVTBlock(pOwner, funcPtrs(callbackMethod))
    If outVtbl = 0 Then Exit Function
    outFO = MakeFakeObj(outVtbl)
    If outFO = 0 Then FreeBlock outVtbl, outFO: Exit Function
    CreateOneShot = True
End Function

Public Sub FreeOneShot(ByRef vtbl As LongPtr, ByRef fo As LongPtr)
    FreeBlock vtbl, fo
End Sub

Public Function UnregisterEvent(ByVal fullKey As String) As Long
    If Not m_handlers.Exists(fullKey) Then UnregisterEvent = -1: Exit Function
    Dim idx As Long: idx = m_handlers(fullKey)
    With m_hData(idx)
        If Not .active Then Exit Function
        If .isEvent And .pInterface <> 0 Then
            UnregisterEvent = dcf(.pInterface, .removeIndex, "remove_" & .fullKey, .token)
        End If
        If .pExtraRelease <> 0 Then dcf .pExtraRelease, 2, "Release_extra": .pExtraRelease = 0
        Dim fk As String: fk = CStr(.fakeObj)
        If m_fakeObjToKey.Exists(fk) Then m_fakeObjToKey.Remove fk
        FreeBlock .vtblBlock, .fakeObj
        .active = False: .pInterface = 0: .token = 0
    End With
    m_handlers.Remove fullKey
End Function

Public Sub UnregisterByPrefix(ByVal prefix As String)
    If m_handlers.Count = 0 Then Exit Sub
    Dim keys As Variant: keys = m_handlers.keys
    Dim i As Long
    For i = UBound(keys) To LBound(keys) Step -1
        Dim k As String: k = CStr(keys(i))
        If Left$(k, Len(prefix)) = prefix Then
            Dim idx As Long: idx = m_handlers(k)
            If m_hData(idx).isEvent Then
                UnregisterEvent k
            Else
                Dim fk As String: fk = CStr(m_hData(idx).fakeObj)
                If m_fakeObjToKey.Exists(fk) Then m_fakeObjToKey.Remove fk
                FreeBlock m_hData(idx).vtblBlock, m_hData(idx).fakeObj
                m_hData(idx).active = False
                m_handlers.Remove k
            End If
        End If
    Next i
End Sub

Private Sub UnregisterAll()
    If m_handlers Is Nothing Then Exit Sub
    If m_handlers.Count = 0 Then Exit Sub
    Dim keys As Variant: keys = m_handlers.keys
    Dim i As Long
    For i = UBound(keys) To LBound(keys) Step -1
        Dim k As String: k = CStr(keys(i))
        If m_handlers.Exists(k) Then
            Dim idx As Long: idx = m_handlers(k)
            If m_hData(idx).isEvent Then
                UnregisterEvent k
            Else
                Dim fk As String: fk = CStr(m_hData(idx).fakeObj)
                If Not m_fakeObjToKey Is Nothing Then
                    If m_fakeObjToKey.Exists(fk) Then m_fakeObjToKey.Remove fk
                End If
                FreeBlock m_hData(idx).vtblBlock, m_hData(idx).fakeObj
                m_hData(idx).active = False
                If m_handlers.Exists(k) Then m_handlers.Remove k
            End If
        End If
    Next i
    m_hCount = 0: Erase m_hData
End Sub

Public Function LookupKeyByFakeObj(ByVal fakeObj As LongPtr) As String
    Dim k As String: k = CStr(fakeObj)
    If m_fakeObjToKey.Exists(k) Then LookupKeyByFakeObj = m_fakeObjToKey(k)
End Function

Public Function HasHandler(ByVal fullKey As String) As Boolean
    HasHandler = m_handlers.Exists(fullKey)
End Function

Public Sub DumpHandlers()
    Debug.Print "=== Handlers (" & m_hCount & ") ==="
    If m_handlers.Count = 0 Then Debug.Print "(none)": Exit Sub
    Dim keys As Variant: keys = m_handlers.keys
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        Dim k As String: k = CStr(keys(i))
        Dim idx As Long: idx = m_handlers(k)
        With m_hData(idx)
            Debug.Print "  [" & idx & "] " & .fullKey & _
                " ev=" & .isEvent & " act=" & .active
        End With
    Next i
End Sub


' =============================================================================
' dcf
' =============================================================================

Public Function dcf( _
    ByVal pInterface As LongPtr, ByVal vtblIndex As Long, _
    ByVal funcName As String, ParamArray args() As Variant) As Long

    If pInterface = 0 Then
        Debug.Print "dcf: null - " & funcName
        dcf = &H80004003: Exit Function
    End If

    Dim pVT As LongPtr: CopyMemory pVT, ByVal pInterface, LenB(pVT)
    Dim pM As LongPtr:  CopyMemory pM, ByVal (pVT + vtblIndex * LenB(pVT)), LenB(pM)

    Dim argc As Long
    On Error Resume Next
    argc = UBound(args) - LBound(args) + 1
    If Err.Number <> 0 Then argc = 0
    On Error GoTo 0

    Dim vt() As Integer, vp() As LongPtr, vals() As Variant
    If argc > 0 Then
        ReDim vt(0 To argc - 1): ReDim vp(0 To argc - 1): ReDim vals(0 To argc - 1)
        Dim i As Long
        For i = 0 To argc - 1
            vals(i) = args(LBound(args) + i)
            Select Case VarType(vals(i))
                Case vbLong
                    vt(i) = vbLong
                Case vbLongLong
                    vt(i) = vbLongLong
                Case vbDouble
                    vt(i) = vbDouble
                Case vbInteger
                    vals(i) = CLng(CInt(vals(i)))
                    vt(i) = vbLong
                Case vbBoolean
                    vals(i) = CLng(IIf(CBool(vals(i)), 1, 0))
                    vt(i) = vbLong
                Case vbString
                    vals(i) = StrPtr(CStr(vals(i)))
                    vt(i) = vbLongLong
                Case Else
                    vt(i) = vbLongLong
            End Select
            vp(i) = VarPtr(vals(i))
        Next i
    End If

    Dim res As Variant, hr As Long
    If argc > 0 Then
        hr = DispCallFunc(pInterface, CLngPtr(vtblIndex) * LenB(pVT), _
            CC_STDCALL, vbLong, argc, vt(0), vp(0), res)
    Else
        hr = DispCallFunc(pInterface, CLngPtr(vtblIndex) * LenB(pVT), _
            CC_STDCALL, vbLong, 0, 0, 0, res)
    End If

    If hr <> 0 Then
        If Len(funcName) > 0 Then Debug.Print "dcf CALL fail: " & funcName & " &H" & Hex(hr)
        dcf = hr
    Else
        Dim lr As Long: lr = CLng(res)
        If lr <> 0 And Len(funcName) > 0 Then
            Debug.Print "dcf METHOD fail: " & funcName & " &H" & Hex(lr)
        End If
        dcf = lr
    End If
End Function


' =============================================================================
' ユーティリティ
' =============================================================================

Public Function PtrToStrW(ByVal p As LongPtr) As String
    If p = 0 Then Exit Function
    Dim cch As Long: cch = lstrlenW(p)
    If cch = 0 Then Exit Function
    PtrToStrW = String$(cch, vbNullChar)
    RtlMoveMemory StrPtr(PtrToStrW), p, CLngPtr(cch * 2)
End Function

Public Function PtsToPx(ByVal pts As Single, ByVal isV As Boolean) As Long
    Dim hdc As LongPtr: hdc = GetDC(0)
    Dim d As Long
    If isV Then d = GetDeviceCaps(hdc, LOGPIXELSY) Else d = GetDeviceCaps(hdc, LOGPIXELSX)
    ReleaseDC 0, hdc
    PtsToPx = CLng(pts * d / 72)
End Function

Public Sub CreateDeepFolder(ByVal fp As String)
    If fp = "" Then Exit Sub
    If Right$(fp, 1) = "\" Then fp = Left$(fp, Len(fp) - 1)
    If Dir(fp, vbDirectory) <> "" Then Exit Sub
    Dim p As String, pos As Long: pos = InStrRev(fp, "\")
    If pos > 0 Then p = Left$(fp, pos - 1): CreateDeepFolder p
    MkDir fp
End Sub


' =============================================================================
' 内部 DispCallFunc ヘルパー
' =============================================================================

Private Function CallCOM0V(ByVal p As LongPtr, ByVal idx As Long) As Long
    Dim r As Variant
    CallCOM0V = DispCallFunc(p, idx * LenB(p), CC_STDCALL, vbLong, 0, 0, 0, r)
End Function
Private Function CallCOM1(ByVal p As LongPtr, ByVal idx As Long, ByVal a1 As LongPtr) As Long
    Dim r As Variant, v(0) As Variant, t(0) As Integer, vp(0) As LongPtr
    v(0) = a1: t(0) = vbLongPtr: vp(0) = VarPtr(v(0))
    CallCOM1 = DispCallFunc(p, idx * LenB(p), CC_STDCALL, vbLong, 1, t(0), vp(0), r)
    If CallCOM1 = 0 Then CallCOM1 = CLng(r)
End Function
Private Sub CallCOM1V(ByVal p As LongPtr, ByVal idx As Long, ByVal a1 As LongPtr)
    Dim r As Variant, v(0) As Variant, t(0) As Integer, vp(0) As LongPtr
    v(0) = a1: t(0) = vbLongPtr: vp(0) = VarPtr(v(0))
    DispCallFunc p, idx * LenB(p), CC_STDCALL, vbLong, 1, t(0), vp(0), r
End Sub
Private Function CallCOM2(ByVal p As LongPtr, ByVal idx As Long, _
    ByVal a1 As Long, ByVal a2 As LongPtr) As Long
    Dim r As Variant, v(1) As Variant, t(1) As Integer, vp(1) As LongPtr
    v(0) = a1: v(1) = a2: t(0) = vbLong: t(1) = vbLongPtr
    vp(0) = VarPtr(v(0)): vp(1) = VarPtr(v(1))
    CallCOM2 = DispCallFunc(p, idx * LenB(p), CC_STDCALL, vbLong, 2, t(0), vp(0), r)
    If CallCOM2 = 0 Then CallCOM2 = CLng(r)
End Function
Private Function CallCOM3(ByVal p As LongPtr, ByVal idx As Long, _
    ByVal a1 As Long, ByVal a2 As Long, ByVal a3 As LongPtr) As Long
    Dim r As Variant, v(2) As Variant, t(2) As Integer, vp(2) As LongPtr
    v(0) = a1: v(1) = a2: v(2) = a3
    t(0) = vbLong: t(1) = vbLong: t(2) = vbLongPtr
    vp(0) = VarPtr(v(0)): vp(1) = VarPtr(v(1)): vp(2) = VarPtr(v(2))
    CallCOM3 = DispCallFunc(p, idx * LenB(p), CC_STDCALL, vbLong, 3, t(0), vp(0), r)
    If CallCOM3 = 0 Then CallCOM3 = CLng(r)
End Function
Private Function CallCOM5(ByVal p As LongPtr, ByVal idx As Long, _
    ByVal a1 As Long, ByVal a2 As LongPtr, ByVal a3 As LongPtr, _
    ByVal a4 As LongPtr, ByVal a5 As LongPtr) As Long
    Dim r As Variant, v(4) As Variant, t(4) As Integer, vp(4) As LongPtr
    v(0) = a1: v(1) = a2: v(2) = a3: v(3) = a4: v(4) = a5
    t(0) = vbLong: t(1) = vbLongPtr: t(2) = vbLongPtr: t(3) = vbLongPtr: t(4) = vbLongPtr
    vp(0) = VarPtr(v(0)): vp(1) = VarPtr(v(1)): vp(2) = VarPtr(v(2))
    vp(3) = VarPtr(v(3)): vp(4) = VarPtr(v(4))
    CallCOM5 = DispCallFunc(p, idx * LenB(p), CC_STDCALL, vbLong, 5, t(0), vp(0), r)
    If CallCOM5 = 0 Then CallCOM5 = CLng(r)
End Function


