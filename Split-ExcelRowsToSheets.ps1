<#
.SYNOPSIS
    Excelの特定シートの特定行以降を、1行ずつ別シート(元シートのコピー)の指定行に貼り付けるスクリプト。

.DESCRIPTION
    指定された元シートを「データ行の数」だけコピーします。
    コピーした各シートでは元データ行はクリアされ、対象の1行だけが指定行に貼り付けられます。
    書式・列幅・数式・条件付き書式などはコピー元シートのものがそのまま保持されます。
    新シート名は「{元シート名}No{連番}」になります。

.PARAMETER FilePath
    対象のExcelファイルのフルパス(必須)。

.PARAMETER SheetName
    元になるシートの名前(必須)。

.PARAMETER StartRow
    データの開始行(必須)。この行以降を1行ずつ別シートに分割します。

.PARAMETER PasteRow
    新シートで貼り付け先となる行番号。既定は 1。

.PARAMETER PasteMode
    貼り付け方法。
      ValuesOnly : 値のみ(既定。書式は新シート側=テンプレートのものを使用)
      AllContents: 値・書式すべて貼り付け
      Formulas   : 数式を保持して貼り付け

.PARAMETER OutputPath
    出力先ファイルパス。省略時は元ファイルを上書き保存。

.PARAMETER KeepOpen
    指定すると処理後もExcelを閉じずに表示します(動作確認用)。

.EXAMPLE
    .\Split-ExcelRowsToSheets.ps1 -FilePath "C:\work\data.xlsx" -SheetName "明細" -StartRow 5

.EXAMPLE
    .\Split-ExcelRowsToSheets.ps1 -FilePath "C:\work\data.xlsx" -SheetName "明細" `
        -StartRow 5 -PasteRow 2 -PasteMode AllContents `
        -OutputPath "C:\work\data_split.xlsx"

.NOTES
    要件 : Windows + Microsoft Excel インストール済み
    実行 : PowerShell 5.1 以上
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$FilePath,

    [Parameter(Mandatory = $true)]
    [string]$SheetName,

    [Parameter(Mandatory = $true)]
    [ValidateRange(1, 1048576)]
    [int]$StartRow,

    [ValidateRange(1, 1048576)]
    [int]$PasteRow = 1,

    [ValidateSet('ValuesOnly', 'AllContents', 'Formulas')]
    [string]$PasteMode = 'ValuesOnly',

    [string]$OutputPath,

    [switch]$KeepOpen
)

# Excel定数 -----------------------------------------------------------------
$xlPasteValues   = -4163
$xlPasteAll      = -4104
$xlPasteFormulas = -4123
$xlCalcManual    = -4135
$xlCalcAutomatic = -4105

switch ($PasteMode) {
    'ValuesOnly'  { $pasteType = $xlPasteValues }
    'AllContents' { $pasteType = $xlPasteAll }
    'Formulas'    { $pasteType = $xlPasteFormulas }
}

# 事前チェック ---------------------------------------------------------------
if (-not (Test-Path -LiteralPath $FilePath)) {
    throw "ファイルが見つかりません: $FilePath"
}
$absPath = (Resolve-Path -LiteralPath $FilePath).Path

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = $absPath
}
else {
    $outDir = Split-Path -Path $OutputPath -Parent
    if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }
}

# Excel COM 起動 -------------------------------------------------------------
$excel    = $null
$workbook = $null
$startTime = Get-Date

try {
    Write-Host "Excelを起動しています..."
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible        = [bool]$KeepOpen
    $excel.DisplayAlerts  = $false
    $excel.ScreenUpdating = $false
    $excel.Calculation    = $xlCalcManual

    Write-Host "ファイルを開いています: $absPath"
    $workbook = $excel.Workbooks.Open($absPath)

    # 元シート取得
    $srcSheet = $null
    foreach ($s in $workbook.Worksheets) {
        if ($s.Name -eq $SheetName) { $srcSheet = $s; break }
    }
    if ($null -eq $srcSheet) {
        $names = ($workbook.Worksheets | ForEach-Object { $_.Name }) -join ', '
        throw "シートが見つかりません: '$SheetName' / 存在するシート: $names"
    }

    # 使用範囲から最終行・最終列を取得
    $usedRange = $srcSheet.UsedRange
    $lastRow   = $usedRange.Row + $usedRange.Rows.Count - 1
    $lastCol   = $usedRange.Column + $usedRange.Columns.Count - 1

    if ($StartRow -gt $lastRow) {
        Write-Warning "StartRow($StartRow) が最終行($lastRow)を超えています。処理対象の行がありません。"
        return
    }

    $dataRowCount = $lastRow - $StartRow + 1
    Write-Host ""
    Write-Host "===== 処理内容 ====="
    Write-Host "対象シート     : $SheetName"
    Write-Host "データ範囲     : 行 $StartRow ～ $lastRow (列 1 ～ $lastCol)"
    Write-Host "作成シート数   : $dataRowCount"
    Write-Host "貼り付けモード : $PasteMode"
    Write-Host "貼り付け先行   : $PasteRow"
    Write-Host "===================="
    Write-Host ""

    # 各行ごとに処理 ---------------------------------------------------------
    $counter = 0
    for ($row = $StartRow; $row -le $lastRow; $row++) {
        $counter++
        $newName = "{0}No{1}" -f $SheetName, $counter

        # シート名が31文字を超える場合は切り詰める(Excelの制限)
        if ($newName.Length -gt 31) {
            $suffix = "No$counter"
            $prefixLen = 31 - $suffix.Length
            $newName = $SheetName.Substring(0, [Math]::Min($SheetName.Length, $prefixLen)) + $suffix
        }

        # 既存の同名シートがあれば削除（再実行対応）
        foreach ($s in $workbook.Worksheets) {
            if ($s.Name -eq $newName) {
                $s.Delete() | Out-Null
                break
            }
        }

        # 元シートを丸ごとコピー（一番右側に配置）
        $lastSheet = $workbook.Worksheets.Item($workbook.Worksheets.Count)
        $srcSheet.Copy([System.Reflection.Missing]::Value, $lastSheet) | Out-Null

        # コピー直後のアクティブシートが新シート
        $newSheet = $workbook.ActiveSheet
        $newSheet.Name = $newName

        # 新シート上で StartRow 以降のデータ行を値だけクリア(書式は残す)
        $clearRange = $newSheet.Range(
            $newSheet.Cells.Item($StartRow, 1),
            $newSheet.Cells.Item($lastRow, $lastCol)
        )
        $clearRange.ClearContents() | Out-Null

        # 元シートの該当行をコピー → 新シートの PasteRow に貼り付け
        $srcRowRange = $srcSheet.Range(
            $srcSheet.Cells.Item($row, 1),
            $srcSheet.Cells.Item($row, $lastCol)
        )
        $dstRange = $newSheet.Cells.Item($PasteRow, 1)

        $srcRowRange.Copy() | Out-Null
        $dstRange.PasteSpecial($pasteType) | Out-Null
        $excel.CutCopyMode = $false

        # 進捗表示(10件ごと or 最終)
        if ($counter % 10 -eq 0 -or $counter -eq $dataRowCount) {
            Write-Host ("  進捗 {0,5}/{1} : 行 {2} → '{3}'" -f $counter, $dataRowCount, $row, $newName)
        }
    }

    # 保存 -------------------------------------------------------------------
    Write-Host ""
    Write-Host "保存中: $OutputPath"

    # 計算モードを戻してから保存
    $excel.Calculation = $xlCalcAutomatic

    if ($OutputPath -ieq $absPath) {
        $workbook.Save()
    }
    else {
        $ext = [System.IO.Path]::GetExtension($OutputPath).ToLower()
        switch ($ext) {
            '.xlsx' { $fmt = 51 }
            '.xlsm' { $fmt = 52 }
            '.xls'  { $fmt = 56 }
            default { $fmt = 51 }
        }
        $workbook.SaveAs($OutputPath, $fmt)
    }

    $elapsed = (Get-Date) - $startTime
    Write-Host ""
    Write-Host ("完了 : {0}シート作成 / 経過 {1:mm}:{1:ss}" -f $dataRowCount, $elapsed)
}
catch {
    Write-Error $_
    throw
}
finally {
    # 後始末 -----------------------------------------------------------------
    if ($workbook -and -not $KeepOpen) {
        $workbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel -and -not $KeepOpen) {
        $excel.ScreenUpdating = $true
        $excel.DisplayAlerts  = $true
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
