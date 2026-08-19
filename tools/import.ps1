<#
    import.ps1 -- src/ の VBA モジュールをブックへ書き戻す

    使い方:
        powershell -ExecutionPolicy Bypass -File tools\import.ps1           # 下見 (既定)
        powershell -ExecutionPolicy Bypass -File tools\import.ps1 -Apply    # 実際に書き戻す
        powershell -ExecutionPolicy Bypass -File tools\import.ps1 -Apply -Modules Wv2Json,Wv2Url

    ★既定は下見 (dry run)★ 何が変わるかを報告するだけでブックには触らない。
    実際に書き戻すには -Apply を明示すること。

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
    [switch]   $NoBackup,
    [switch]   $SkipGate
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

Write-Host ("ブック  : " + $Book)
Write-Host ("ソース  : " + $SrcDir)
Write-Host ("モード  : " + $modeLabel)
Write-Host ""

# ---------------------------------------------------------------
# ヘッダーを落として「コード本体」だけを取り出す
#   .cls : VERSION 行 / BEGIN..END ブロック / Attribute VB_xxx の連なり
#   .bas : Attribute VB_Name の 1 行
# メンバー属性 (Attribute m_browser.VB_VarHelpID) は本体側なので落とさない
# ---------------------------------------------------------------
function Get-CodeBody([string]$path) {
    $text = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
    $text = $text -replace "`r`n", "`n"
    $text = $text -replace "`r", "`n"
    $lines = [System.Collections.Generic.List[string]]::new()
    foreach ($l in ($text -split "`n")) { [void]$lines.Add($l) }
    # 末尾の空要素 (最終改行に由来する 1 個) だけ落とす
    if ($lines.Count -gt 0 -and $lines[$lines.Count - 1] -eq '') { $lines.RemoveAt($lines.Count - 1) }

    $i = 0
    while ($i -lt $lines.Count) {
        $l = $lines[$i]
        if ($l -match '^VERSION\s') { $i++; continue }
        if ($l -eq 'BEGIN') {
            $i++
            while ($i -lt $lines.Count -and $lines[$i] -ne 'END') { $i++ }
            if ($i -lt $lines.Count) { $i++ }
            continue
        }
        if ($l -match '^Attribute VB_[A-Za-z]+ = ') { $i++; continue }
        break
    }
    $body = @()
    if ($i -lt $lines.Count) { $body = @($lines[$i..($lines.Count - 1)]) }
    # 末尾の空行を全部落とす。AddFromString が改行を 1 つ足すので、
    # 落としておかないと import のたびに空行が 1 行ずつ増え続ける
    while ($body.Count -gt 0 -and $body[$body.Count - 1] -eq '') {
        $body = @($body[0..($body.Count - 2)])
    }
    return ($body -join "`n")
}

# CodeModule.Lines はメンバー属性行 (Attribute m_browser.VB_VarHelpID = -1) を
# 返さないが、Export は書き出す。比較のときだけ両側から落とす。
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
$missing = @()

try {
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
        $body = Get-CodeBody $f.FullName

        if (-not $byName.ContainsKey($name)) {
            $missing += $name
            $rows += [pscustomobject]@{ モジュール = $name; 状態 = '★ブックに存在しない★'; 行数 = ($body -split "`n").Count }
            continue
        }

        $comp = $byName[$name]
        $cm = $comp.CodeModule
        $cur = ''
        if ($cm.CountOfLines -ge 1) { $cur = Normalize $cm.Lines(1, $cm.CountOfLines) }

        $curCmp  = Normalize (Strip-MemberAttrs $cur)
        $bodyCmp = Normalize (Strip-MemberAttrs $body)

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
            # --- 6. 事後検証: 読み直して一致するか ---
            $back = ''
            if ($cm.CountOfLines -ge 1) { $back = Normalize $cm.Lines(1, $cm.CountOfLines) }
            $state = '★書き戻し後の照合に失敗★'
            if ((Normalize (Strip-MemberAttrs $back)) -eq $bodyCmp) { $state = '書き戻した (照合 OK)' }
            $rows += [pscustomobject]@{ モジュール = $name; 状態 = $state; 行数 = ($before.ToString() + ' -> ' + $cm.CountOfLines) }
        } else {
            $rows += [pscustomobject]@{ モジュール = $name; 状態 = '要更新'; 行数 = ($before.ToString() + ' -> ' + $after) }
        }
    }

    if ($Apply -and $changedNames.Count -gt 0) {
        $wb.Save()
        Write-Host ("ブックを保存した (" + $changedNames.Count + " モジュール)")
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
}

$rows | Format-Table -AutoSize

$failed = @($rows | Where-Object { $_.状態 -like '*失敗*' })
Write-Host ("変更あり: " + $changedNames.Count + " / " + $files.Count)

if ($missing.Count -gt 0) {
    Write-Host ""
    Write-Host ("ブックに存在しないモジュール: " + ($missing -join ', ')) -ForegroundColor Yellow
    Write-Host "  本スクリプトは既存コンポーネントの中身の差し替えだけを行う。" -ForegroundColor Yellow
    Write-Host "  新規モジュールは VBE で空のモジュールを作ってから再実行すること。" -ForegroundColor Yellow
}

if ($failed.Count -gt 0) {
    Write-Host ""
    Write-Host "★書き戻し後の照合に失敗したモジュールがある★" -ForegroundColor Red
    Write-Host "  CP932 に無い文字が落ちた可能性が高い。バックアップから戻すこと。" -ForegroundColor Red
    exit 1
}

if (-not $Apply) {
    Write-Host ""
    Write-Host "※ 下見のみ。ブックには触っていない。書き戻すには -Apply を付けること。"
} else {
    Write-Host ""
    Write-Host "※ 書き戻し後、tools\export.ps1 -Force を走らせて src\ と一致することを確認するとよい。"
}
