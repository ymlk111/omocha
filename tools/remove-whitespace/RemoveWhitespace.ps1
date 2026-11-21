# ============================================================
#  RemoveWhitespace.ps1
#  ドロップしたファイルから 半角スペース・タブ・改行（＋全角スペース）を全て除去する
#
#  使い方:
#    同じフォルダにある RemoveWhitespace.bat にファイルをドラッグ＆ドロップする。
#    （.ps1 は Windows の標準設定ではドロップを受け付けないため bat 経由で起動する）
#
#  出力（元ファイルは変更しない）:
#    ・ファイルごと : 元ファイルと同じフォルダに「元の名前_nospace.拡張子」
#    ・まとめる     : 先頭ファイルのフォルダに「merged_nospace_日時.拡張子」
#                     結合順はファイル名順（エクスプローラーのドロップ順は当てにならないため）
#
#  文字コード:
#    BOM付きUTF-8 / UTF-8 / Shift_JIS / UTF-16 を自動判定し、出力も同じ文字コードで書く。
#    まとめる場合、入力の文字コードが混在していたら UTF-8（BOMなし）で書く。
#    NUL を含むファイルはバイナリとみなしてスキップする。
# ============================================================
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Paths
)

# ===== 設定 =====
$RemoveFullWidthSpace = $true        # $true: 全角スペース(U+3000)も除去する
$MergeSeparator       = ''           # まとめる時のファイル間の区切り。"`r`n" にすると1ファイル=1行になる
$OutputSuffix         = '_nospace'   # 出力ファイル名に付ける文字列
# ================

# PowerShell 7 で Shift_JIS を扱うための登録（5.1 では不要。失敗しても無視）
try { [System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance) } catch { }

$pattern = if ($RemoveFullWidthSpace) { '[ \t\r\n　]' } else { '[ \t\r\n]' }
$WhitespaceRegex = New-Object System.Text.RegularExpressions.Regex($pattern)

function Remove-Whitespace([string]$Text) {
    $WhitespaceRegex.Replace($Text, '')
}

# 文字コードを判定してテキストとして読み込む
function Read-TextFile([string]$Path) {
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $len   = $bytes.Length
    $skip  = 0

    if ($len -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $enc = New-Object System.Text.UTF8Encoding($true);            $name = 'UTF-8(BOM付き)'; $skip = 3
    }
    elseif ($len -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        $enc = New-Object System.Text.UnicodeEncoding($false, $true); $name = 'UTF-16LE';       $skip = 2
    }
    elseif ($len -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        $enc = New-Object System.Text.UnicodeEncoding($true, $true);  $name = 'UTF-16BE';       $skip = 2
    }
    else {
        if ([Array]::IndexOf($bytes, [byte]0) -ge 0) {
            throw 'バイナリファイルと判断したためスキップしました'
        }
        try {
            # 不正なバイト列で例外を出す厳密モードで UTF-8 として読めるか試す
            $strict = New-Object System.Text.UTF8Encoding($false, $true)
            [void]$strict.GetString($bytes)
            $enc = New-Object System.Text.UTF8Encoding($false); $name = 'UTF-8'
        }
        catch {
            $enc = [System.Text.Encoding]::GetEncoding(932);   $name = 'Shift_JIS'
        }
    }

    [pscustomobject]@{
        Text         = $enc.GetString($bytes, $skip, $len - $skip)
        Encoding     = $enc
        EncodingName = $name
    }
}

function Get-ErrorMessage($ErrorRecord) {
    $ErrorRecord.Exception.GetBaseException().Message
}

# ファイルごとに処理
function Invoke-Each($Files) {
    $ok = 0
    foreach ($f in $Files) {
        try {
            $r       = Read-TextFile $f.FullName
            $out     = Remove-Whitespace $r.Text
            $outPath = [System.IO.Path]::Combine($f.DirectoryName, $f.BaseName + $OutputSuffix + $f.Extension)
            [System.IO.File]::WriteAllText($outPath, $out, $r.Encoding)

            Write-Host ("OK  {0}  [{1}]  除去 {2:N0} 文字" -f $f.Name, $r.EncodingName, ($r.Text.Length - $out.Length)) -ForegroundColor Green
            Write-Host "    -> $outPath"
            $ok++
        }
        catch {
            Write-Host ("NG  {0}: {1}" -f $f.Name, (Get-ErrorMessage $_)) -ForegroundColor Red
        }
    }
    Write-Host ''
    Write-Host ("完了: {0} / {1} 件" -f $ok, @($Files).Count)
}

# 全てまとめて1ファイルに
function Invoke-Merge($Files) {
    $sorted = @($Files | Sort-Object Name)
    $parts  = New-Object System.Collections.Generic.List[string]
    $used   = @()

    Write-Host '結合順（ファイル名順）:'
    foreach ($f in $sorted) {
        try {
            $r = Read-TextFile $f.FullName
            $parts.Add((Remove-Whitespace $r.Text))
            $used += [pscustomobject]@{ File = $f; Result = $r }
            Write-Host ("  OK  {0}  [{1}]" -f $f.Name, $r.EncodingName) -ForegroundColor Green
        }
        catch {
            Write-Host ("  NG  {0}: {1}  ※結合から除外" -f $f.Name, (Get-ErrorMessage $_)) -ForegroundColor Red
        }
    }

    if ($used.Count -eq 0) {
        Write-Host '結合できるファイルがありませんでした。' -ForegroundColor Red
        return
    }

    # 文字コード: 全部同じならそれ、混在なら UTF-8(BOMなし)
    $encNames = @($used | ForEach-Object { $_.Result.EncodingName } | Select-Object -Unique)
    if ($encNames.Count -eq 1) {
        $enc     = $used[0].Result.Encoding
        $encName = $encNames[0]
    }
    else {
        $enc     = New-Object System.Text.UTF8Encoding($false)
        $encName = 'UTF-8（入力の文字コードが混在していたため）'
    }

    # 拡張子: 全部同じならそれ、違えば .txt
    $exts = @($used | ForEach-Object { $_.File.Extension.ToLower() } | Select-Object -Unique)
    $ext  = if ($exts.Count -eq 1 -and $exts[0]) { $used[0].File.Extension } else { '.txt' }

    $merged  = $parts -join $MergeSeparator
    $outName = 'merged' + $OutputSuffix + '_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + $ext
    $outPath = [System.IO.Path]::Combine($used[0].File.DirectoryName, $outName)
    [System.IO.File]::WriteAllText($outPath, $merged, $enc)

    Write-Host ''
    Write-Host ("完了: {0} / {1} 件を結合  [{2}]  {3:N0} 文字" -f $used.Count, $sorted.Count, $encName, $merged.Length) -ForegroundColor Green
    Write-Host "    -> $outPath"
}

function Invoke-Main {
    # ドロップされたもののうちファイルだけを対象にする
    $files = @()
    foreach ($p in $Paths) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        if (Test-Path -LiteralPath $p -PathType Leaf) {
            $files += Get-Item -LiteralPath $p -Force
        }
        else {
            Write-Host "スキップ（ファイルではありません）: $p" -ForegroundColor Yellow
        }
    }

    if ($files.Count -eq 0) {
        Write-Host '処理するファイルがありません。RemoveWhitespace.bat にファイルをドラッグ＆ドロップしてください。'
        return
    }

    $mode = 'each'
    if ($files.Count -ge 2) {
        Write-Host ("{0} 件のファイルが指定されました。処理方法を選んでください。" -f $files.Count)
        Write-Host '  1: ファイルごとに処理する'
        Write-Host '  2: 全てまとめて1ファイルにする'
        Write-Host '  q: 何もせず終了'
        while ($true) {
            $ans = "$(Read-Host '番号を入力')".Trim()
            if ($ans -in @('1', '１'))      { $mode = 'each';  break }
            if ($ans -in @('2', '２'))      { $mode = 'merge'; break }
            if ($ans -in @('q', 'ｑ', 'Ｑ')) { Write-Host 'キャンセルしました。'; return }
        }
        Write-Host ''
    }

    if ($mode -eq 'each') { Invoke-Each $files } else { Invoke-Merge $files }
}

try {
    Invoke-Main
}
catch {
    Write-Host ("エラー: {0}" -f (Get-ErrorMessage $_)) -ForegroundColor Red
}
Write-Host ''
[void](Read-Host 'Enterキーで閉じます')
