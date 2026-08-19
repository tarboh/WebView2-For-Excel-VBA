Attribute VB_Name = "ModHookSupport"
'==============================================================================
' ModHookSupport.bas  (x64専用 / 方針3)
'
' WndProc サブクラス化用のサンク生成と EbMode スタブパッチ。
'
' 動作概要:
'   サブクラス化プロシージャ (= インスタンス別サンク) は OS から
'   WndProc(hwnd, msg, wParam, lParam) の4引数で呼ばれる。
'
'   [サンク]
'     1. g_CurrentThis = pThis
'     2. スタブ (EbModeAnchor) を CALL
'          - Break中: スタブが RET 0 で戻る → サンクが 0 を返す
'          - 通常: スタブが書き換え済みの +37 (= ディスパッチャ) へ JMP
'     3. CALL 結果 (RAX) をそのまま返す
'
'   [ディスパッチャ] (スタブ JMP の先、末尾呼び出し)
'     入り口の時点でレジスタは:
'       RCX=hwnd, RDX=msg, R8=wParam, R9=lParam
'     ここで this を第1引数に挿入する必要がある:
'       RCX=pThis, RDX=hwnd, R8=msg, R9=wParam, [RSP+28h]=lParam
'     → 引数シフト後、vtable[WndProc_Slot] を CALL して戻る。
'==============================================================================
Option Explicit

Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, _
    ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
Private Declare PtrSafe Function VirtualFree Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, _
    ByVal dwFreeType As Long) As Long
Private Declare PtrSafe Function VirtualProtect Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, _
    ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" ( _
    ByVal Destination As LongPtr, ByVal Source As LongPtr, _
    ByVal Length As LongPtr)

Private Const NullPtr As LongLong = 0^
Private Const MEM_COMMIT As Long = &H1000
Private Const MEM_RESERVE As Long = &H2000
Private Const MEM_RELEASE As Long = &H8000
Private Const PAGE_EXECUTE_READWRITE As Long = &H40

' スタブ +37 のパッチ位置 (ダンプから確定)
Private Const STUB_PATCH_OFFSET As Long = &H37
Private Const STUB_PATCH_SIZE As Long = 8

'==============================================================================
' モジュール状態
'==============================================================================
Public g_CurrentThis As LongPtr    ' ディスパッチャが読む現在のインスタンス

Private m_Initialized As Boolean
Private m_pDispatcher As LongPtr
Private m_pStub As LongPtr
Private m_OrigUserCode As LongPtr
Private m_VTableSlot As Long

'==============================================================================
' EbMode チェック用アンカー (スタブ本体)
'==============================================================================
Public Sub EbModeAnchor()
    ' 空実装。VBA コンパイラが EbMode チェック入りのプロローグを生成する。
End Sub

Public Sub CopyMemory(ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)
    RtlMoveMemory Destination, Source, Length
End Sub

'==============================================================================
' スタブ +37 パッチ + ディスパッチャ生成 (一度だけ実行)
'==============================================================================
Public Sub EnsureDispatcherInstalled(ByVal vtSlot As Long)
    If m_Initialized Then
        If m_VTableSlot <> vtSlot Then
            Err.Raise 5, "EnsureDispatcherInstalled", _
                "Dispatcher already installed with different vtSlot"
        End If
        Exit Sub
    End If

    m_pStub = VBA.Int(AddressOf EbModeAnchor)
    m_VTableSlot = vtSlot
    m_pDispatcher = BuildDispatcher(vtSlot)

    Dim pPatch As LongPtr
    pPatch = m_pStub + STUB_PATCH_OFFSET
    CopyMemory VarPtr(m_OrigUserCode), pPatch, STUB_PATCH_SIZE

    Dim oldProt As Long
    If VirtualProtect(pPatch, STUB_PATCH_SIZE, PAGE_EXECUTE_READWRITE, oldProt) = 0 Then
        Err.Raise 7, "EnsureDispatcherInstalled", "VirtualProtect failed"
    End If
    CopyMemory pPatch, VarPtr(m_pDispatcher), STUB_PATCH_SIZE
    VirtualProtect pPatch, STUB_PATCH_SIZE, oldProt, oldProt

    m_Initialized = True
End Sub

'==============================================================================
' ディスパッチャ生成 (WndProc 4引数版)
'
'   入り時点: RCX=hwnd, RDX=msg, R8=wParam, R9=lParam
'   出し時点: RCX=pThis, RDX=hwnd, R8=msg, R9=wParam, [RSP+28h]=lParam
'
'   x64 呼出規約では第5引数以降はスタックに積む。[RSP+20h]以降が
'   第5引数以降のスロット (ただし+00?+18はシャドースペース)。
'   つまり [RSP+28h] が第5引数の位置。
'==============================================================================
Private Function BuildDispatcher(ByVal vtSlot As Long) As LongPtr
    Dim code(0 To 127) As Byte
    Dim i As Long: i = 0

    ' sub rsp, 38h    48 83 EC 38    ; shadow(20h) + 1引数(8) + align(8) = 30h+8
    '                                ; CALL時のpush retでさらに+8 → RSP+38hで16byte整列
    code(i) = &H48: i = i + 1
    code(i) = &H83: i = i + 1
    code(i) = &HEC: i = i + 1
    code(i) = &H38: i = i + 1

    ' mov [rsp+28h], r9    4C 89 4C 24 28    ; lParam をスタックへ (第5引数)
    code(i) = &H4C: i = i + 1
    code(i) = &H89: i = i + 1
    code(i) = &H4C: i = i + 1
    code(i) = &H24: i = i + 1
    code(i) = &H28: i = i + 1

    ' mov r9, r8      4D 89 C1    ; wParam を第4引数へ
    code(i) = &H4D: i = i + 1
    code(i) = &H89: i = i + 1
    code(i) = &HC1: i = i + 1

    ' mov r8, rdx     49 89 D0    ; msg を第3引数へ
    code(i) = &H49: i = i + 1
    code(i) = &H89: i = i + 1
    code(i) = &HD0: i = i + 1

    ' mov rdx, rcx    48 89 CA    ; hwnd を第2引数へ
    code(i) = &H48: i = i + 1
    code(i) = &H89: i = i + 1
    code(i) = &HCA: i = i + 1

    ' mov rax, &g_CurrentThis       48 B8 <imm64>
    code(i) = &H48: i = i + 1
    code(i) = &HB8: i = i + 1
    Dim pG As LongPtr: pG = VarPtr(g_CurrentThis)
    PutBytes code, i, VarPtr(pG), 8: i = i + 8
    ' mov rcx, [rax]                48 8B 08    ; 第1引数 = this
    code(i) = &H48: i = i + 1
    code(i) = &H8B: i = i + 1
    code(i) = &H8:  i = i + 1

    ' mov rax, [rcx]                48 8B 01    ; vtable
    code(i) = &H48: i = i + 1
    code(i) = &H8B: i = i + 1
    code(i) = &H1:  i = i + 1

    ' call [rax + vtSlot*8]         FF 50 <disp8>
    code(i) = &HFF: i = i + 1
    code(i) = &H50: i = i + 1
    code(i) = CByte(vtSlot * 8): i = i + 1

    ' add rsp, 38h    48 83 C4 38
    code(i) = &H48: i = i + 1
    code(i) = &H83: i = i + 1
    code(i) = &HC4: i = i + 1
    code(i) = &H38: i = i + 1

    ' ret              C3
    code(i) = &HC3: i = i + 1

    Dim pMem As LongPtr
    pMem = VirtualAlloc(NullPtr, i, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If pMem = NullPtr Then
        Err.Raise 7, "BuildDispatcher", "VirtualAlloc failed"
    End If
    CopyMemory pMem, VarPtr(code(0)), i
    BuildDispatcher = pMem
End Function

'==============================================================================
' インスタンス別サンク生成 (メッセージフィルタ付き)
'   OS から WndProc(hwnd, msg, wParam, lParam) で呼ばれる。
'   ***重要***: 描画メッセージ等すべてを VBA に回すと UserForm が破綻するため、
'               msg が msgMatch1 または msgMatch2 のときだけ VBA へ入り、
'               それ以外は元の WndProc (pOrigProc) へ末尾ジャンプする。
'
'   x64 レジスタ: RCX=hwnd, RDX=msg, R8=wParam, R9=lParam
'   アセンブリ (疑似):
'     cmp edx, msgMatch1
'     je  VBAPath
'     cmp edx, msgMatch2
'     je  VBAPath
'     mov rax, pOrigProc
'     jmp rax                    ; 元の WndProc へ末尾委譲 (スタックそのまま)
'   VBAPath:
'     sub rsp, 28h
'     mov [g_CurrentThis], pThis
'     call stub
'     add rsp, 28h
'     ret
'==============================================================================
Public Function BuildInstanceThunk(ByVal pThis As LongPtr, _
                                    ByVal msgMatch1 As Long, _
                                    ByVal msgMatch2 As Long, _
                                    ByVal pOrigProc As LongPtr, _
                                    ByRef outSize As LongPtr) As LongPtr
    If Not m_Initialized Then
        Err.Raise 5, "BuildInstanceThunk", _
            "Call EnsureDispatcherInstalled first"
    End If

    Dim code(0 To 127) As Byte
    Dim i As Long: i = 0

    ' --- フィルタ部 ---
    ' cmp edx, imm32    81 FA <imm32>
    code(i) = &H81: i = i + 1
    code(i) = &HFA: i = i + 1
    PutBytes code, i, VarPtr(msgMatch1), 4: i = i + 4
    ' je rel8  (VBAPath へ、後でオフセット埋め戻し)
    code(i) = &H74: i = i + 1
    Dim je1Pos As Long: je1Pos = i
    code(i) = &H0: i = i + 1  ' プレースホルダ

    ' cmp edx, imm32
    code(i) = &H81: i = i + 1
    code(i) = &HFA: i = i + 1
    PutBytes code, i, VarPtr(msgMatch2), 4: i = i + 4
    ' je rel8
    code(i) = &H74: i = i + 1
    Dim je2Pos As Long: je2Pos = i
    code(i) = &H0: i = i + 1

    ' --- 非マッチ: 元 WndProc へ末尾ジャンプ ---
    ' mov rax, pOrigProc    48 B8 <imm64>
    code(i) = &H48: i = i + 1
    code(i) = &HB8: i = i + 1
    PutBytes code, i, VarPtr(pOrigProc), 8: i = i + 8
    ' jmp rax               FF E0
    code(i) = &HFF: i = i + 1
    code(i) = &HE0: i = i + 1

    ' --- VBAPath: ここが je の飛び先 ---
    Dim vbaPathPos As Long: vbaPathPos = i
    ' je の rel8 を埋める
    code(je1Pos) = CByte(vbaPathPos - (je1Pos + 1))
    code(je2Pos) = CByte(vbaPathPos - (je2Pos + 1))

    ' sub rsp, 28h    48 83 EC 28
    code(i) = &H48: i = i + 1
    code(i) = &H83: i = i + 1
    code(i) = &HEC: i = i + 1
    code(i) = &H28: i = i + 1

    ' mov rax, &g_CurrentThis    48 B8 <imm64>
    code(i) = &H48: i = i + 1
    code(i) = &HB8: i = i + 1
    Dim pG As LongPtr: pG = VarPtr(g_CurrentThis)
    PutBytes code, i, VarPtr(pG), 8: i = i + 8
    ' mov r10, pThis             49 BA <imm64>
    code(i) = &H49: i = i + 1
    code(i) = &HBA: i = i + 1
    PutBytes code, i, VarPtr(pThis), 8: i = i + 8
    ' mov [rax], r10             4C 89 10
    code(i) = &H4C: i = i + 1
    code(i) = &H89: i = i + 1
    code(i) = &H10: i = i + 1

    ' mov rax, m_pStub           48 B8 <imm64>
    code(i) = &H48: i = i + 1
    code(i) = &HB8: i = i + 1
    PutBytes code, i, VarPtr(m_pStub), 8: i = i + 8
    ' call rax                   FF D0
    code(i) = &HFF: i = i + 1
    code(i) = &HD0: i = i + 1

    ' add rsp, 28h    48 83 C4 28
    code(i) = &H48: i = i + 1
    code(i) = &H83: i = i + 1
    code(i) = &HC4: i = i + 1
    code(i) = &H28: i = i + 1
    ' ret              C3
    code(i) = &HC3: i = i + 1

    outSize = i
    Dim pMem As LongPtr
    pMem = VirtualAlloc(NullPtr, outSize, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If pMem = NullPtr Then
        Err.Raise 7, "BuildInstanceThunk", "VirtualAlloc failed"
    End If
    CopyMemory pMem, VarPtr(code(0)), outSize
    BuildInstanceThunk = pMem
End Function

Public Sub FreeThunk(ByVal pThunk As LongPtr)
    If pThunk <> NullPtr Then VirtualFree pThunk, 0, MEM_RELEASE
End Sub

Private Sub PutBytes(ByRef buf() As Byte, ByVal offset As Long, _
                      ByVal src As LongPtr, ByVal n As Long)
    CopyMemory VarPtr(buf(offset)), src, n
End Sub


