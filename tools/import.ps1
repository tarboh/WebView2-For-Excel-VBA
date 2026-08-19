<#
    import.ps1 -- src/ の VBA モジュールをブックへ書き戻す

    使い方:
        powershell -ExecutionPolicy Bypass -File tools\import.ps1                      # 下見 (既定)
        powershell -ExecutionPolicy Bypass -File tools\import.ps1 -Apply               # 書き戻す
        powershell -ExecutionPolicy Bypass -File tools\import.ps1 -Apply -AllowNew     # 新規追加も許す
        powershell -ExecutionPolicy Bypass -File tools\import.ps1 -Apply -Modules Wv2Json,Wv2Url

    ★既定は下見 (dry run)★ 何が変わるかを報告するだけでブックには触らない。
    実際に書き戻すには -Apply を明示すること (設計原則96)。

    設計 (E-1 で合意した 7 点):
        1. import 前にブックを book\backup\ へタイムスタンプ付きで複製する
        2. 対象は src\*.bas と src\*.cls のみ。src\document\ と forms\ は対象外
        3. VBComponents.Remove + Import は使わない。削除が遅延して同名 Import が
           Wv2Pane1 に化ける罠があるため。CodeModule を全削除して AddFromString で
           流し込む (コンポーネントを消さないので名前と属性が保たれる)
        4. 内容が変わったモジュールだけ入れる
        5. 事前ゲート: check_cp932.py を走らせ、1 件でも出たら中止
        6. 事後検証: 書き戻したモジュールを読み直して src\ と一致することを確認する
        7. Excel は不可視で開いて保存する

    E-3 で追加した「新規モジュールの追加」:
        - ★既存の差し替えは AddFromString、新規の追加は Import のハイブリッド★
          Remove しないので 3. の罠は当たらない。Import ならヘッダーの属性
          (VB_PredeclaredId / VB_Exposed 等) と名前が自動で正しく付く
        - Import に渡す一時ファイルは ★CP932 で書く★。事前ゲートが表現可能性を
          保証しているので確実 (仕様事実35 / 36)
        - 新規追加には -AllowNew が要る。src\ にゴミを置いたとき勝手に増えるのを防ぐ
        - Attribute VB_Name がファイル名と違ったら中止する。Import は VB_Name で
          名前が決まるので、ズレると別名のモジュールができる

    文字コード:
        AddFromString は BSTR (UTF-16) を渡すので、ファイル側の文字コードは
        PowerShell が UTF-8 として読んだ時点で解決している。ただし VBA は
        受け取った文字列をプロジェクトのコードページ (CP932) で保持するため、
        CP932 に無い文字はここで失われる (仕様事実35)。
        だから 5. の事前ゲートは必須。
#>

[CmdletBinding()]
param(
    [string]   $Book,
    [string]   $SrcDir,
    [string[]] $Modules,
    [switch]   $Apply,
    [switch]   $AllowNew,
    [switch]   $NoBackup,
    [switch]   $SkipGate,
    [switch]   $SkipHeaderCheck,
    [switch]   $SkipDriftCheck
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot

if (-not $Book) {
    $cand = @(Get-ChildItem (Join-Path $repoRoot 'book') -Filter '*.xlsm' -ErrorAction SilentlyContinue)
    if ($cand.Count -eq 0) { throw "book フォルダに .xlsm が見つからない。-Book で指定すること。" }
    if ($cand.Count -gt 1) { throw ("book フォルダに .xlsm が複数ある: " + ($cand.Name -join ', ')) }
    $Book = $cand[0].FullName
}
$Book = (Resolve-Path $Book).Path
if (-not $SrcDir) { $SrcDir = Join-Path $repoRoot 'src' }
$SrcDir = (Resolve-Path $SrcDir).Path

$modeLabel = '下見のみ (ブックには触らない)'
if ($Apply) { $modeLabel = '★書き戻す★' }
if ($Apply -and $AllowNew) { $modeLabel = '★書き戻す + 新規追加を許可★' }

Write-Host ("ブック  : " + $Book)
Write-Host ("ソース  : " + $SrcDir)
Write-Host ("モード  : " + $modeLabel)
Write-Host ""

$cp932 = [System.Text.Encoding]::GetEncoding(932)

# ---------------------------------------------------------------
# モジュールファイルをヘッダーと本体に割る
#   .cls : VERSION 行 / BEGIN..END ブロック / Attribute VB_xxx の連なり
#   .bas : Attribute VB_Name の 1 行
# メンバー属性 (Attribute m_browser.VB_VarHelpID) は本体側なので落とさない
# ---------------------------------------------------------------
function Split-ModuleFile([string]$path) {
    $text = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
    $text = $text -replace "`r`n", "`n"
    $text = $text -replace "`r", "`n"
    $all = @($text -split "`n")
    if ($all.Count -gt 0 -and $all[$all.Count - 1] -eq '') { $all = @($all[0..($all.Count - 2)]) }

    $i = 0
    while ($i -lt $all.Count) {
        $l = $all[$i]
        if ($l -match '^VERSION\s') { $i++; continue }
        if ($l -eq 'BEGIN') {
            $i++
            while ($i -lt $all.Count -and $all[$i] -ne 'END') { $i++ }
            if ($i -lt $all.Count) { $i++ }
            continue
        }
        if ($l -match '^Attribute VB_[A-Za-z]+ = ') { $i++; continue }
        break
    }

    $header = @()
    if ($i -gt 0) { $header = @($all[0..($i - 1)]) }
    $body = @()
    if ($i -lt $all.Count) { $body = @($all[$i..($all.Count - 1)]) }
    # 末尾の空行を全部落とす。AddFromString が改行を 1 つ足すので、
    # 落としておかないと import のたびに空行が 1 行ずつ増え続ける (仕様事実40)
    while ($body.Count -gt 0 -and $body[$body.Count - 1] -eq '') {
        $body = @($body[0..($body.Count - 2)])
    }

    $attrs = @{}
    foreach ($l in $header) {
        if ($l -match '^Attribute (VB_[A-Za-z]+) = (.*)$') {
            $attrs[$Matches[1]] = $Matches[2].Trim()
        }
    }

    return [pscustomobject]@{
        Header = ($header -join "`n")
        Body   = ($body -join "`n")
        Attrs  = $attrs
        Full   = $text
    }
}

# CodeModule.Lines はメンバー属性行 (Attribute m_browser.VB_VarHelpID = -1) を
# 返さないが、Export は書き出す (仕様事実39)。比較のときだけ両側から落とす。
# 書き込みには残す (AddFromString は受け付け、メンバー属性として保存する)。
function Strip-MemberAttrs([string]$s) {
    $keep = @()
    foreach ($l in ($s -split "`n")) {
        if ($l -match '^Attribute\s+\w+\.\w+\s*=') { continue }
        $keep += $l
    }
    return ($keep -join "`n")
}

function Normalize([string]$s) {
    if ($null -eq $s) { return '' }
    $s = $s -replace "`r`n", "`n"
    $s = $s -replace "`r", "`n"
    return ($s -replace "`n+$", '')
}

# ---- 5. 事前ゲート -------------------------------------------------------
if (-not $SkipGate) {
    $checker = Join-Path $PSScriptRoot 'check_cp932.py'
    if (Test-Path $checker) {
        Write-Host "事前ゲート: check_cp932.py"
        & python $checker $SrcDir
        if ($LASTEXITCODE -ne 0) {
            Write-Host ""
            Write-Host "★中止★ CP932 で往復できない文字がある。ブックへ入れると永久に失われる (仕様事実35)。" -ForegroundColor Red
            Write-Host "上の代替候補に置き換えてから再実行すること。" -ForegroundColor Red
            exit 1
        }
        Write-Host ""
    } else {
        Write-Host "警告: check_cp932.py が見つからないので事前ゲートを飛ばした。" -ForegroundColor Yellow
    }
}

# ---- 対象ファイルの収集 --------------------------------------------------
$files = @(Get-ChildItem $SrcDir -File | Where-Object { @('.bas', '.cls') -contains $_.Extension } | Sort-Object Name)
if ($Modules -and $Modules.Count -gt 0) {
    $files = @($files | Where-Object { $Modules -contains $_.BaseName })
    if ($files.Count -eq 0) { throw ("-Modules に一致するファイルが無い: " + ($Modules -join ', ')) }
}

# ---- 論点5: VB_Name とファイル名の一致を先に検査 -------------------------
$parsed = @{}
$nameErrors = @()
foreach ($f in $files) {
    $m = Split-ModuleFile $f.FullName
    $parsed[$f.BaseName] = $m
    $vbName = $m.Attrs['VB_Name']
    if ($null -eq $vbName) {
        $nameErrors += ($f.Name + " : Attribute VB_Name が無い")
    } else {
        $vbName = $vbName.Trim('"')
        if ($vbName -ne $f.BaseName) {
            $nameErrors += ($f.Name + " : Attribute VB_Name = " + $vbName + " がファイル名と違う")
        }
    }
}
if ($nameErrors.Count -gt 0) {
    Write-Host "★中止★ Attribute VB_Name とファイル名が食い違っている:" -ForegroundColor Red
    foreach ($e in $nameErrors) { Write-Host ("  " + $e) -ForegroundColor Red }
    Write-Host "Import は VB_Name で名前が決まるので、ズレたまま入れると別名のモジュールができる。" -ForegroundColor Red
    exit 1
}

Write-Host ("対象: " + $files.Count + " ファイル")
Write-Host ""

# ---- 3. バックアップ (書き戻すときだけ) ----------------------------------
if ($Apply -and (-not $NoBackup)) {
    $backupDir = Join-Path (Split-Path -Parent $Book) 'backup'
    if (-not (Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $dest = Join-Path $backupDir ([System.IO.Path]::GetFileNameWithoutExtension($Book) + '_' + $stamp + '.xlsm')
    Copy-Item $Book $dest
    Write-Host ("バックアップ: " + $dest)
    Write-Host ""
}

# ---- Excel 起動 ---------------------------------------------------------
$excel = $null
$wb = $null
$rows = @()
$changedNames = @()
$addedNames = @()
$blocked = @()
$headerWarn = @()
$drifted = @()
$tmpDir = Join-Path ([System.IO.Path]::GetTempPath()) ('wv2import_' + [Guid]::NewGuid().ToString('N'))

try {
    New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3   # マクロ強制無効。Workbook_Open を発火させない

    $readOnly = -not $Apply
    for ($i = 1; $i -le 10; $i++) {
        try { $wb = $excel.Workbooks.Open($Book, 0, $readOnly); break }
        catch { Start-Sleep -Milliseconds 800 }
    }
    if ($null -eq $wb) { throw "Workbooks.Open に 10 回失敗した。Excel を一度終了してから再実行すること。" }

    $vbp = $null
    try { $vbp = $wb.VBProject } catch { }
    if ($null -eq $vbp) { throw "VBProject にアクセスできない。トラスト センターの設定を確認すること。" }

    $byName = @{}
    foreach ($c in $vbp.VBComponents) { $byName[$c.Name] = $c }

    foreach ($f in $files) {
        $name = $f.BaseName
        $m = $parsed[$name]
        $body = $m.Body
        $bodyCmp = Normalize (Strip-MemberAttrs $body)

        # ================= 新規追加 (E-3) =================
        if (-not $byName.ContainsKey($name)) {
            if (-not $AllowNew) {
                $blocked += $name
                $rows += [pscustomobject]@{ モジュール = $name; 状態 = '★ブックに無い (-AllowNew が要る)★'; 行数 = ($body -split "`n").Count }
                continue
            }
            if (-not $Apply) {
                $rows += [pscustomobject]@{ モジュール = $name; 状態 = '新規追加の予定'; 行数 = ($body -split "`n").Count }
                continue
            }

            # 一時ファイルを CP932 + CRLF で書いて Import する
            # ★CRLF は必須★ LF だけで書くと Import がヘッダーブロック
            #   (VERSION 1.0 CLASS / BEGIN..END) を認識できず、クラスモジュールが
            #   標準モジュールとして作られてヘッダーがコードに落ちる (E-3 で踏んだ)
            $tmp = Join-Path $tmpDir ($name + $f.Extension)
            $full = $m.Full -replace "`n", "`r`n"
            if (-not $full.EndsWith("`r`n")) { $full = $full + "`r`n" }
            $bytes = $cp932.GetBytes($full)
            # 往復検証 (事前ゲートを通っていれば必ず通る。二重の保険)
            if ($cp932.GetString($bytes) -ne $full) {
                throw ($name + " : CP932 への変換が非可逆。事前ゲートを確認すること。")
            }
            [System.IO.File]::WriteAllBytes($tmp, $bytes)

            $newComp = $vbp.VBComponents.Import($tmp)
            $actualName = $newComp.Name
            $actualType = $newComp.Type
            $byName[$actualName] = $newComp

            $wantType = 1
            if ($f.Extension -eq '.cls') { $wantType = 2 }

            $state = '★追加後の名前が違う: ' + $actualName + '★'
            if ($actualName -ne $name) {
                # そのまま
            } elseif ($actualType -ne $wantType) {
                $state = '★種別が違う (期待 ' + $wantType + ' / 実際 ' + $actualType + ')★'
            } else {
                $cmN = $newComp.CodeModule
                $backN = ''
                if ($cmN.CountOfLines -ge 1) { $backN = Normalize $cmN.Lines(1, $cmN.CountOfLines) }
                if ((Normalize (Strip-MemberAttrs $backN)) -eq $bodyCmp) {
                    $state = '新規追加した (照合 OK)'
                    $addedNames += $name
                } else {
                    $state = '★追加したが本体の照合に失敗★'
                }
            }
            $rows += [pscustomobject]@{ モジュール = $name; 状態 = $state; 行数 = ('新規 -> ' + $newComp.CodeModule.CountOfLines) }
            continue
        }

        # ================= 既存の差し替え =================
        $comp = $byName[$name]
        if (@(1, 2) -notcontains $comp.Type) {
            $rows += [pscustomobject]@{ モジュール = $name; 状態 = ('★種別が違うので飛ばした (Type=' + $comp.Type + ')★'); 行数 = '-' }
            continue
        }

        $cm = $comp.CodeModule
        $cur = ''
        if ($cm.CountOfLines -ge 1) { $cur = Normalize $cm.Lines(1, $cm.CountOfLines) }
        $curCmp = Normalize (Strip-MemberAttrs $cur)

        if ($curCmp -eq $bodyCmp) {
            $rows += [pscustomobject]@{ モジュール = $name; 状態 = '変更なし'; 行数 = $cm.CountOfLines }
            continue
        }

        $changedNames += $name
        $before = $cm.CountOfLines
        $after = ($body -split "`n").Count

        if ($Apply) {
            if ($cm.CountOfLines -ge 1) { $cm.DeleteLines(1, $cm.CountOfLines) }
            $cm.AddFromString($body)
            $back = ''
            if ($cm.CountOfLines -ge 1) { $back = Normalize $cm.Lines(1, $cm.CountOfLines) }
            $state = '★書き戻し後の照合に失敗★'
            if ((Normalize (Strip-MemberAttrs $back)) -eq $bodyCmp) { $state = '書き戻した (照合 OK)' }
            $rows += [pscustomobject]@{ モジュール = $name; 状態 = $state; 行数 = ($before.ToString() + ' -> ' + $cm.CountOfLines) }
        } else {
            $rows += [pscustomobject]@{ モジュール = $name; 状態 = '要更新'; 行数 = ($before.ToString() + ' -> ' + $after) }
        }
    }

    # ---- 論点4: ヘッダー属性の照合 --------------------------------------
    # VBIDE には VB_PredeclaredId 等を読む API が無いので、Export して読む。
    # VBIDE から設定する手段も無いので、食い違ったら警告するだけ (手で直す)。
    if (-not $SkipHeaderCheck) {
        foreach ($f in $files) {
            $name = $f.BaseName
            if (-not $byName.ContainsKey($name)) { continue }
            $comp = $byName[$name]
            if (@(1, 2) -notcontains $comp.Type) { continue }
            $hp = Join-Path $tmpDir ('hdr_' + $name + $f.Extension)
            $comp.Export($hp)
            $hb = [System.IO.File]::ReadAllBytes($hp)
            $ht = $cp932.GetString($hb)
            $actual = @{}
            foreach ($l in ($ht -replace "`r`n", "`n" -split "`n")) {
                if ($l -match '^Attribute (VB_[A-Za-z]+) = (.*)$') { $actual[$Matches[1]] = $Matches[2].Trim() }
                elseif ($l -notmatch '^(VERSION|BEGIN|END|\s+\w+ =|Attribute VB_)') { break }
            }
            foreach ($k in $parsed[$name].Attrs.Keys) {
                $want = $parsed[$name].Attrs[$k]
                $got = $actual[$k]
                if ($got -ne $want) {
                    $headerWarn += ($name + " : " + $k + " が src=" + $want + " / ブック=" + $got)
                }
            }
        }
    }

    if ($Apply -and ($changedNames.Count + $addedNames.Count) -gt 0) {
        $wb.Save()
        Write-Host ("ブックを保存した (差し替え " + $changedNames.Count + " / 新規 " + $addedNames.Count + ")")
        Write-Host ""

    }
}
finally {
    if ($null -ne $wb)    { try { $wb.Close($false) } catch { } }
    if ($null -ne $excel) { try { $excel.Quit() } catch { } }
    foreach ($o in @($wb, $excel)) {
        if ($null -ne $o) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if (Test-Path $tmpDir) { Remove-Item $tmpDir -Recurse -Force -ErrorAction SilentlyContinue }
}

# ---- 保存後のドリフト検査 (E-3) -----------------------------------------
# ★VBA は識別子の大小をプロジェクト全体で統一する (仕様事実38)★
#   新規モジュールが既存と同じ綴りの識別子を別の大小で宣言していると、
#   コンパイル時に「触っていないモジュール」の中身まで書き換わる。
#   動作は変わらないが git diff が汚れ、本当の変更を覆い隠すので必ず表に出す。
#
# ★この正規化は、import を行った PowerShell プロセスからは観測できない★
#   Excel を終了させて新しい COM インスタンスを作っても、正規化前の綴りが返る
#   (E-3 で実測)。そのため自前で比較すると「ドリフトなし」と嘘をつく。
#   別プロセスで export.ps1 を呼んでバイト照合する形にしてある。
if ($Apply -and ($changedNames.Count + $addedNames.Count) -gt 0 -and (-not $SkipDriftCheck)) {
    Write-Host "保存後のドリフト検査 (別プロセスで export し直して照合) ..."
    $dTmp = Join-Path ([System.IO.Path]::GetTempPath()) ('wv2drift_' + [Guid]::NewGuid().ToString('N'))
    try {
        New-Item -ItemType Directory -Path $dTmp -Force | Out-Null
        $expPath = Join-Path $PSScriptRoot 'export.ps1'
        & powershell -ExecutionPolicy Bypass -NoProfile -File $expPath -Book $Book -OutDir $dTmp -Force | Out-Null

        $dSrc = Join-Path $dTmp 'src'
        if (-not (Test-Path $dSrc)) {
            Write-Host "  ★ドリフト検査ができなかった (export に失敗)★" -ForegroundColor Yellow
        } else {
            foreach ($f in $files) {
                $nm = $f.BaseName
                if ($changedNames -contains $nm) { continue }
                if ($addedNames -contains $nm) { continue }
                $g = Join-Path $dSrc $f.Name
                if (-not (Test-Path $g)) { continue }
                $h1 = (Get-FileHash $f.FullName -Algorithm MD5).Hash
                $h2 = (Get-FileHash $g -Algorithm MD5).Hash
                if ($h1 -ne $h2) { $drifted += $nm }
            }
            if ($drifted.Count -eq 0) { Write-Host "  ドリフトなし" }
        }
    }
    finally {
        if (Test-Path $dTmp) { Remove-Item $dTmp -Recurse -Force -ErrorAction SilentlyContinue }
    }
    Write-Host ""
}

$rows | Format-Table -AutoSize

$failed = @($rows | Where-Object { $_.状態 -like '*失敗*' -or $_.状態 -like '*名前が違う*' })
Write-Host ("差し替え: " + $changedNames.Count + " / 新規: " + $addedNames.Count + " / 対象: " + $files.Count)

if ($blocked.Count -gt 0) {
    Write-Host ""
    Write-Host ("ブックに無いモジュール: " + ($blocked -join ', ')) -ForegroundColor Yellow
    Write-Host "  追加してよければ -AllowNew を付けて再実行すること。" -ForegroundColor Yellow
}

if ($drifted.Count -gt 0) {
    Write-Host ""
    Write-Host "★書き戻していないのに中身が変わったモジュール★" -ForegroundColor Yellow
    Write-Host ("  " + ($drifted -join ', ')) -ForegroundColor Yellow
    Write-Host "  VBA が識別子の大小をプロジェクト全体で統一したため (仕様事実38)。" -ForegroundColor Yellow
    Write-Host "  動作は変わらないが、export.ps1 -Force を走らせて src\ に取り込むこと。" -ForegroundColor Yellow
    Write-Host "  取り込まないと、次回以降ずっと差分として出続ける。" -ForegroundColor Yellow
}

if ($headerWarn.Count -gt 0) {
    Write-Host ""
    Write-Host "ヘッダー属性が食い違っている:" -ForegroundColor Yellow
    foreach ($w in $headerWarn) { Write-Host ("  " + $w) -ForegroundColor Yellow }
    Write-Host "  ★VBIDE からは設定できない属性なので、VBE で手当てすること。★" -ForegroundColor Yellow
}

if ($failed.Count -gt 0) {
    Write-Host ""
    Write-Host "★照合に失敗したモジュールがある★" -ForegroundColor Red
    Write-Host "  book\backup\ の複製から戻すこと。" -ForegroundColor Red
    exit 1
}

if (-not $Apply) {
    Write-Host ""
    Write-Host "※ 下見のみ。ブックには触っていない。書き戻すには -Apply を付けること。"
} else {
    Write-Host ""
    Write-Host "※ 書き戻し後、tools\export.ps1 -Force を走らせて src\ と一致することを確認するとよい。"
}
