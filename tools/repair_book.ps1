<#
    repair_book.ps1 -- 露出したメンバー属性行を取り除く (コンポーネントの作り直し)

    使い方:
        powershell -ExecutionPolicy Bypass -File tools\repair_book.ps1 -Modules Wv2Browser
        powershell -ExecutionPolicy Bypass -File tools\repair_book.ps1 -Modules Wv2Browser -Apply

    ★既定は下見★ 実行するには -Apply が要る (設計原則96)。

    何をするか:
        1. ブックを book\backup\ へ複製する
        2. 【第1段】ブックを開いて対象コンポーネントを Remove し、保存して閉じ、
           ★Excel を完全に終了する★
        3. 【第2段】新しい Excel で開き直し、src\ のファイルを Import して保存する
        4. check_book.py で重複が消えたことを確かめる

    ★なぜ Remove + Import なのか★
        メンバー属性行が二重になると、片方が VBE のコードペインに露出して
        コンパイルエラーになる。ところが CodeModule.Lines は Attribute 行を
        返さないので、CodeModule.DeleteLines では消せない。
        コンポーネントごと作り直すしかない。

    ★なぜ 2 段階に分けるのか★
        VBComponents.Remove は削除が遅延する。同じセッションで同名を Import すると
        Wv2Browser1 のような別名で作られる (E-1 で記録済みの罠)。
        Excel を終了させてから Import すれば、削除が確定した状態で始められる。
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string[]] $Modules,
    [string]   $Book,
    [string]   $SrcDir,
    [switch]   $Apply,
    [switch]   $NoBackup
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

if (-not $Book) {
    $cand = @(Get-ChildItem (Join-Path $repoRoot 'book') -Filter '*.xlsm' -ErrorAction SilentlyContinue)
    if ($cand.Count -ne 1) { throw "book フォルダの .xlsm を 1 本に絞れない。-Book で指定すること。" }
    $Book = $cand[0].FullName
}
$Book = (Resolve-Path $Book).Path
if (-not $SrcDir) { $SrcDir = Join-Path $repoRoot 'src' }
$SrcDir = (Resolve-Path $SrcDir).Path

$cp932 = [System.Text.Encoding]::GetEncoding(932)

Write-Host ("ブック  : " + $Book)
Write-Host ("対象    : " + ($Modules -join ', '))
$modeLabel = '下見のみ (ブックには触らない)'
if ($Apply) { $modeLabel = '★作り直す★' }
Write-Host ("モード  : " + $modeLabel)
Write-Host ""

# ---- 対象ファイルの存在確認 ----------------------------------------------
$srcFiles = @{}
foreach ($m in $Modules) {
    $f = $null
    foreach ($ext in @('.cls', '.bas')) {
        $p = Join-Path $SrcDir ($m + $ext)
        if (Test-Path $p) { $f = $p; break }
    }
    if ($null -eq $f) { throw ("src に " + $m + ".cls / .bas が見つからない") }
    $srcFiles[$m] = $f
    Write-Host ("  " + $m + " <- " + $f)
}
Write-Host ""

if (-not $Apply) {
    Write-Host "※ 下見のみ。実行するには -Apply を付けること。"
    Write-Host "   実行すると対象コンポーネントを一度削除し、src から Import し直す。"
    return
}

# ---- バックアップ ---------------------------------------------------------
if (-not $NoBackup) {
    $backupDir = Join-Path (Split-Path -Parent $Book) 'backup'
    if (-not (Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $dest = Join-Path $backupDir ([System.IO.Path]::GetFileNameWithoutExtension($Book) + '_repair_' + $stamp + '.xlsm')
    Copy-Item $Book $dest
    Write-Host ("バックアップ: " + $dest)
    Write-Host ""
}

function Open-Book([bool]$readOnly) {
    $ex = New-Object -ComObject Excel.Application
    $ex.Visible = $false
    $ex.DisplayAlerts = $false
    $ex.EnableEvents = $false
    $ex.AutomationSecurity = 3
    $w = $null
    for ($i = 1; $i -le 10; $i++) {
        try { $w = $ex.Workbooks.Open($Book, 0, $readOnly); break }
        catch { Start-Sleep -Milliseconds 800 }
    }
    if ($null -eq $w) { throw "Workbooks.Open に失敗した。Excel を終了してから再実行すること。" }
    return @($ex, $w)
}

function Close-Book($ex, $w, [bool]$save) {
    if ($null -ne $w) {
        try { if ($save) { $w.Save() } } catch { }
        try { $w.Close($false) } catch { }
    }
    if ($null -ne $ex) { try { $ex.Quit() } catch { } }
    foreach ($o in @($w, $ex)) {
        if ($null -ne $o) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# ---- 【第1段】削除 --------------------------------------------------------
Write-Host "【第1段】コンポーネントを削除する"
$r = Open-Book $false
$ex1 = $r[0]; $wb1 = $r[1]
try {
    $vbp = $wb1.VBProject
    foreach ($m in $Modules) {
        $comp = $null
        try { $comp = $vbp.VBComponents.Item($m) } catch { }
        if ($null -eq $comp) {
            Write-Host ("  " + $m + " : ブックに無い (飛ばす)") -ForegroundColor Yellow
            continue
        }
        $vbp.VBComponents.Remove($comp)
        Write-Host ("  " + $m + " : Remove した")
    }
}
finally {
    Close-Book $ex1 $wb1 $true
}

Write-Host "  Excel を終了して削除を確定させる ..."
Start-Sleep -Seconds 3
$still = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
if ($still.Count -gt 0) {
    Write-Host ("  警告: Excel がまだ " + $still.Count + " プロセス残っている") -ForegroundColor Yellow
    Start-Sleep -Seconds 3
}
Write-Host ""

# ---- 【第2段】Import ------------------------------------------------------
Write-Host "【第2段】src から Import し直す"
$tmpDir = Join-Path ([System.IO.Path]::GetTempPath()) ('wv2repair_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null

$r = Open-Book $false
$ex2 = $r[0]; $wb2 = $r[1]
$failed = @()
try {
    $vbp = $wb2.VBProject
    foreach ($m in $Modules) {
        $srcPath = $srcFiles[$m]
        $text = [System.IO.File]::ReadAllText($srcPath, [System.Text.Encoding]::UTF8)
        $text = $text -replace "`r`n", "`n"
        $text = $text -replace "`r", "`n"
        $text = $text -replace "`n", "`r`n"
        if (-not $text.EndsWith("`r`n")) { $text = $text + "`r`n" }

        $bytes = $cp932.GetBytes($text)
        if ($cp932.GetString($bytes) -ne $text) {
            throw ($m + " : CP932 への変換が非可逆。check_cp932.py を確認すること。")
        }
        $tmp = Join-Path $tmpDir ([System.IO.Path]::GetFileName($srcPath))
        [System.IO.File]::WriteAllBytes($tmp, $bytes)

        $newComp = $vbp.VBComponents.Import($tmp)
        if ($newComp.Name -ne $m) {
            $failed += ($m + " : Import 後の名前が " + $newComp.Name + " になった")
            Write-Host ("  " + $m + " : ★名前が " + $newComp.Name + " になった★") -ForegroundColor Red
        } else {
            Write-Host ("  " + $m + " : Import した (" + $newComp.CodeModule.CountOfLines + " 行)")
        }
    }
}
finally {
    Close-Book $ex2 $wb2 $true
    if (Test-Path $tmpDir) { Remove-Item $tmpDir -Recurse -Force -ErrorAction SilentlyContinue }
}

Write-Host ""
if ($failed.Count -gt 0) {
    Write-Host "★失敗した★" -ForegroundColor Red
    foreach ($f in $failed) { Write-Host ("  " + $f) -ForegroundColor Red }
    Write-Host "  book\backup\ の複製から戻すこと。" -ForegroundColor Red
    exit 1
}

Write-Host "作り直しが終わった。check_book.py で確かめること:"
Write-Host "  python tools\check_book.py"
