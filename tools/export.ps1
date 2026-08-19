<#
    export.ps1 -- ブックの VBA モジュールをファイルへ書き出す

    使い方:
        powershell -ExecutionPolicy Bypass -File tools\export.ps1
        powershell -ExecutionPolicy Bypass -File tools\export.ps1 -OutDir C:\temp\dump -Raw

    前提:
        - Excel の「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が ON
          (HKCU:\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM = 1)
        - Excel が起動していないこと

    出力:
        src\           標準モジュール (.bas) / クラスモジュール (.cls)
        src\document\  ThisWorkbook / Sheet* (import では置換できないので参照用)
        forms\         ユーザーフォーム (.frm + .frx) -- -Forms を付けたときだけ
                       .frx はエクスポートのたびに数バイト変化するため既定では触らない

    文字コード:
        VBE の Export は CP932 (システムの ANSI コードページ) で書き出す。
        本スクリプトはそれを UTF-8 / BOM なし / CRLF へ変換して保存する。
        変換後に CP932 へ戻して元バイト列と一致することを検証し、
        1 バイトでも違えば異常として報告する。
        .frx はバイナリなので一切触らない。

    安全性:
        - ブックは ReadOnly で開き、SaveChanges:=$false で閉じる。ブックは変更しない
        - マクロは AutomationSecurity で強制無効化する (Workbook_Open を発火させない)
        - ファイルを消す動作は無い (同名は上書き。既存があれば -Force が要る)
#>

[CmdletBinding()]
param(
    [string] $Book,
    [string] $OutDir,
    [switch] $Raw,
    [switch] $Forms,
    [switch] $Force
)

$ErrorActionPreference = 'Stop'

# ---- パスの決定 ---------------------------------------------------------
$repoRoot = Split-Path -Parent $PSScriptRoot

if (-not $Book) {
    $cand = @(Get-ChildItem (Join-Path $repoRoot 'book') -Filter '*.xlsm' -ErrorAction SilentlyContinue)
    if ($cand.Count -eq 0) { throw "book フォルダに .xlsm が見つからない。-Book で指定すること。" }
    if ($cand.Count -gt 1) { throw ("book フォルダに .xlsm が複数ある: " + ($cand.Name -join ', ') + " / -Book で指定すること。") }
    $Book = $cand[0].FullName
}
$Book = (Resolve-Path $Book).Path
if (-not $OutDir) { $OutDir = $repoRoot }
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
$OutDir = (Resolve-Path $OutDir).Path

$modeLabel = 'src / src\document / forms へ振り分け'
if ($Raw) { $modeLabel = 'Raw (フラット)' }

Write-Host ("ブック  : " + $Book)
Write-Host ("出力先  : " + $OutDir)
Write-Host ("モード  : " + $modeLabel)
Write-Host ""

# ---- 出力先の準備 -------------------------------------------------------
if ($Raw) {
    $dirStd = $OutDir; $dirCls = $OutDir; $dirFrm = $OutDir; $dirDoc = $OutDir
} else {
    $dirStd = Join-Path $OutDir 'src'
    $dirCls = Join-Path $OutDir 'src'
    $dirFrm = Join-Path $OutDir 'forms'
    $dirDoc = Join-Path (Join-Path $OutDir 'src') 'document'
}
foreach ($d in @($dirStd, $dirCls, $dirFrm, $dirDoc)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}

if (-not $Force) {
    $existing = @()
    foreach ($d in (@($dirStd, $dirCls, $dirFrm, $dirDoc) | Select-Object -Unique)) {
        $existing += @(Get-ChildItem $d -File -ErrorAction SilentlyContinue |
                       Where-Object { @('.bas', '.cls', '.frm', '.frx') -contains $_.Extension })
    }
    if ($existing.Count -gt 0) {
        Write-Host ("出力先に既に " + $existing.Count + " 本のモジュールがある。上書きしてよければ -Force を付けて再実行すること。") -ForegroundColor Yellow
        Write-Host ("  例: " + (($existing | Select-Object -First 5 | ForEach-Object { $_.Name }) -join ', '))
        return
    }
}

$cp932 = [System.Text.Encoding]::GetEncoding(932)
$utf8  = New-Object System.Text.UTF8Encoding($false)

# ---- Excel 起動 ---------------------------------------------------------
$excel = $null
$wb = $null
$rows = @()
$warn = @()
$skipped = @()

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3   # msoAutomationSecurityForceDisable

    for ($i = 1; $i -le 10; $i++) {
        try { $wb = $excel.Workbooks.Open($Book, 0, $true); break }
        catch { Start-Sleep -Milliseconds 800 }
    }
    if ($null -eq $wb) { throw "Workbooks.Open に 10 回失敗した。Excel を一度終了してから再実行すること。" }

    $vbp = $null
    try { $vbp = $wb.VBProject } catch { }
    if ($null -eq $vbp) {
        throw "VBProject にアクセスできない。Excel の [ファイル]-[オプション]-[トラスト センター]-[トラスト センターの設定]-[マクロの設定] で「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を ON にすること。"
    }

    foreach ($comp in $vbp.VBComponents) {
        $name = $comp.Name
        $type = $comp.Type          # 1=標準 2=クラス 3=フォーム 100=ドキュメント
        $lines = 0
        try { $lines = $comp.CodeModule.CountOfLines } catch { }

        switch ($type) {
            1   { $ext = '.bas'; $dest = $dirStd; $kind = '標準モジュール' }
            2   { $ext = '.cls'; $dest = $dirCls; $kind = 'クラスモジュール' }
            3   { $ext = '.frm'; $dest = $dirFrm; $kind = 'ユーザーフォーム' }
            100 { $ext = '.cls'; $dest = $dirDoc; $kind = 'ドキュメント' }
            default { $ext = '.txt'; $dest = $dirDoc; $kind = ('その他(Type=' + $type + ')') }
        }

        if (($type -eq 3) -and (-not $Forms)) {
            $skipped += ($name + ' (ユーザーフォーム。-Forms を付けたときだけ書き出す)')
            continue
        }
        $path = Join-Path $dest ($name + $ext)
        $comp.Export($path)

        # --- CP932 -> UTF-8 変換 (往復検証つき) ---
        $srcBytes  = [System.IO.File]::ReadAllBytes($path)
        $text = $cp932.GetString($srcBytes)
        $back = $cp932.GetBytes($text)
        $lossless = ($back.Length -eq $srcBytes.Length)
        if ($lossless) {
            for ($k = 0; $k -lt $srcBytes.Length; $k++) {
                if ($back[$k] -ne $srcBytes[$k]) { $lossless = $false; break }
            }
        }
        if (-not $lossless) {
            $warn += ($name + $ext + " : CP932 への往復が非可逆。UTF-8 変換を見送って原本のまま残した。")
        } else {
            [System.IO.File]::WriteAllText($path, $text, $utf8)
        }

        $size = 0
        if (Test-Path $path) { $size = (Get-Item $path).Length }

        $rows += [pscustomobject]@{
            名前   = $name
            種別   = $kind
            行数   = $lines
            バイト = $size
            出力先 = (Resolve-Path $path).Path.Substring($OutDir.Length).TrimStart('\')
        }
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

$rows | Sort-Object 種別, 名前 | Format-Table -AutoSize
Write-Host ("合計 " + $rows.Count + " コンポーネント、" + (($rows | Measure-Object 行数 -Sum).Sum) + " 行")

if ($warn.Count -gt 0) {
    Write-Host ""
    Write-Host "警告:" -ForegroundColor Yellow
    foreach ($w in $warn) { Write-Host ("  " + $w) -ForegroundColor Yellow }
}

if ($skipped.Count -gt 0) {
    Write-Host ""
    Write-Host "書き出さなかったもの:"
    foreach ($s in $skipped) { Write-Host ("  " + $s) }
}

Write-Host ""
Write-Host "※ ブックは読み取り専用で開き、保存せずに閉じた。ブックは変更されていない。"
Write-Host "※ .frx はバイナリなので変換していない。"
