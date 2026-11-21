# ============================================================
#  RemoveWhitespace.ps1
#  Removes all spaces, tabs and line breaks (and full-width spaces)
#  from the dropped files.
#
#  Usage:
#    Drop files onto RemoveWhitespace.bat in the same folder.
#    (.ps1 files do not accept drag & drop on default Windows settings)
#
#  Output (original files are never modified):
#    each  : <name>_nospace.<ext> next to each original file
#    merge : merged_nospace_<datetime>.<ext> in the first file's folder
#            Files are joined in file-name order
#            (the order Explorer passes dropped files is unreliable).
#
#  Encoding:
#    Auto-detects UTF-8 (with/without BOM), Shift_JIS and UTF-16,
#    and writes the output in the same encoding.
#    When merging files with mixed encodings, output is UTF-8 (no BOM).
#    Files containing NUL bytes are treated as binary and skipped.
#
#  NOTE: Keep this file ASCII-only. Windows PowerShell 5.1 misreads
#        non-ASCII characters depending on the file encoding / BOM.
# ============================================================

# Capture arguments at script level ($args inside a function is different)
$Paths = @($args)

# ===== Settings =====
$RemoveFullWidthSpace = $true        # $true: also remove full-width space (U+3000)
$MergeSeparator       = ''           # separator between files when merging ("`r`n" = one file per line)
$OutputSuffix         = '_nospace'   # suffix added to output file names
# ====================

# Needed for Shift_JIS on PowerShell 7 (not needed on 5.1; ignore failure)
try { [System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance) } catch { }

$pattern = if ($RemoveFullWidthSpace) { '[ \t\r\n\u3000]' } else { '[ \t\r\n]' }
$WhitespaceRegex = New-Object System.Text.RegularExpressions.Regex($pattern)

function Remove-Whitespace([string]$Text) {
    $WhitespaceRegex.Replace($Text, '')
}

# Detect the encoding and read the file as text
function Read-TextFile([string]$Path) {
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $len   = $bytes.Length
    $skip  = 0

    if ($len -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $enc = New-Object System.Text.UTF8Encoding($true);            $name = 'UTF-8 BOM'; $skip = 3
    }
    elseif ($len -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        $enc = New-Object System.Text.UnicodeEncoding($false, $true); $name = 'UTF-16LE';  $skip = 2
    }
    elseif ($len -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        $enc = New-Object System.Text.UnicodeEncoding($true, $true);  $name = 'UTF-16BE';  $skip = 2
    }
    else {
        if ([Array]::IndexOf($bytes, [byte]0) -ge 0) {
            throw 'Skipped: looks like a binary file (contains NUL bytes)'
        }
        try {
            # Strict mode throws on invalid byte sequences
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

# Process each file separately
function Invoke-Each($Files) {
    $ok = 0
    foreach ($f in $Files) {
        try {
            $r       = Read-TextFile $f.FullName
            $out     = Remove-Whitespace $r.Text
            $outPath = [System.IO.Path]::Combine($f.DirectoryName, $f.BaseName + $OutputSuffix + $f.Extension)
            [System.IO.File]::WriteAllText($outPath, $out, $r.Encoding)

            Write-Host ("OK  {0}  [{1}]  removed {2:N0} chars" -f $f.Name, $r.EncodingName, ($r.Text.Length - $out.Length)) -ForegroundColor Green
            Write-Host "    -> $outPath"
            $ok++
        }
        catch {
            Write-Host ("NG  {0}: {1}" -f $f.Name, (Get-ErrorMessage $_)) -ForegroundColor Red
        }
    }
    Write-Host ''
    Write-Host ("Done: {0} / {1} files" -f $ok, @($Files).Count)
}

# Merge all files into one
function Invoke-Merge($Files) {
    $sorted = @($Files | Sort-Object Name)
    $parts  = New-Object System.Collections.Generic.List[string]
    $used   = @()

    Write-Host 'Merge order (by file name):'
    foreach ($f in $sorted) {
        try {
            $r = Read-TextFile $f.FullName
            $parts.Add((Remove-Whitespace $r.Text))
            $used += [pscustomobject]@{ File = $f; Result = $r }
            Write-Host ("  OK  {0}  [{1}]" -f $f.Name, $r.EncodingName) -ForegroundColor Green
        }
        catch {
            Write-Host ("  NG  {0}: {1}  (excluded from merge)" -f $f.Name, (Get-ErrorMessage $_)) -ForegroundColor Red
        }
    }

    if ($used.Count -eq 0) {
        Write-Host 'Nothing to merge.' -ForegroundColor Red
        return
    }

    # Encoding: keep it if all inputs share one, otherwise UTF-8 (no BOM)
    $encNames = @($used | ForEach-Object { $_.Result.EncodingName } | Select-Object -Unique)
    if ($encNames.Count -eq 1) {
        $enc     = $used[0].Result.Encoding
        $encName = $encNames[0]
    }
    else {
        $enc     = New-Object System.Text.UTF8Encoding($false)
        $encName = 'UTF-8 (inputs had mixed encodings)'
    }

    # Extension: keep it if all inputs share one, otherwise .txt
    $exts = @($used | ForEach-Object { $_.File.Extension.ToLower() } | Select-Object -Unique)
    $ext  = if ($exts.Count -eq 1 -and $exts[0]) { $used[0].File.Extension } else { '.txt' }

    $merged  = $parts -join $MergeSeparator
    $outName = 'merged' + $OutputSuffix + '_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + $ext
    $outPath = [System.IO.Path]::Combine($used[0].File.DirectoryName, $outName)
    [System.IO.File]::WriteAllText($outPath, $merged, $enc)

    Write-Host ''
    Write-Host ("Done: merged {0} / {1} files  [{2}]  {3:N0} chars" -f $used.Count, $sorted.Count, $encName, $merged.Length) -ForegroundColor Green
    Write-Host "    -> $outPath"
}

function Invoke-Main {
    # Use only the dropped items that are files
    $files = @()
    foreach ($p in $Paths) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        if (Test-Path -LiteralPath $p -PathType Leaf) {
            $files += Get-Item -LiteralPath $p -Force
        }
        else {
            Write-Host "Skipped (not a file): $p" -ForegroundColor Yellow
        }
    }

    if ($files.Count -eq 0) {
        Write-Host 'No files to process. Drop files onto RemoveWhitespace.bat.'
        return
    }

    $mode = 'each'
    if ($files.Count -ge 2) {
        Write-Host ("{0} files were given. How do you want to process them?" -f $files.Count)
        Write-Host '  1: Process each file separately'
        Write-Host '  2: Merge all into one file'
        Write-Host '  q: Quit without doing anything'
        while ($true) {
            # FormKC turns full-width input (e.g. from IME) into half-width
            $ans = "$(Read-Host 'Enter 1, 2 or q')".Normalize([System.Text.NormalizationForm]::FormKC).Trim()
            if ($ans -eq '1') { $mode = 'each';  break }
            if ($ans -eq '2') { $mode = 'merge'; break }
            if ($ans -eq 'q') { Write-Host 'Cancelled.'; return }
        }
        Write-Host ''
    }

    if ($mode -eq 'each') { Invoke-Each $files } else { Invoke-Merge $files }
}

try {
    Invoke-Main
}
catch {
    Write-Host ("Error: {0}" -f (Get-ErrorMessage $_)) -ForegroundColor Red
}
Write-Host ''
[void](Read-Host 'Press Enter to close')
