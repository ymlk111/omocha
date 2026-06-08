<#
.SYNOPSIS
    ログファイルへの追記を監視し、新規追記行を Excel のアクティブシートへ
    1行＝1セルで自動転記する常駐 GUI ツール。

.DESCRIPTION
    - 監視方式: Timer ポーリング(1秒)でファイルサイズの変化を検知
    - ON 時点のファイルサイズを基準オフセットとし、以降の追記分のみ転記
    - ローテーション/クリア(サイズ減少)時は先頭から読み直し
    - 書き込み先は GUI で指定したブックのアクティブシート
    - アクティブシート切替を検知したら書き込み位置を開始セルにリセット
    - ログの文字コードは Shift-JIS

.NOTES
    実行環境: Windows + Excel インストール済み / PowerShell
#>

# ===== アセンブリ読み込み =====
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ===== グローバル状態 =====
$script:isWatching      = $false      # 監視 ON/OFF
$script:baseOffset      = 0           # 基準バイトオフセット(ON時点のサイズ)
$script:lastSize        = 0           # 前回確認時のファイルサイズ
$script:writeRow        = 0           # 次に書き込む行(開始セル基準のオフセット)
$script:startCol        = 0           # 書き込み開始列(1始まり)
$script:startRow        = 0           # 書き込み開始行(1始まり)
$script:lastSheetName   = ""          # 直近のアクティブシート名(切替検知用)
$script:targetWorkbook  = $null       # 対象ブックの COM オブジェクト
$script:excelApp        = $null       # Excel.Application の COM オブジェクト
$script:writtenCount    = 0           # 累計書き込み行数

# Shift-JIS エンコーディング
$script:sjis = [System.Text.Encoding]::GetEncoding("Shift_JIS")

# ===== ユーティリティ: 起動中ブック一覧の取得 =====
function Get-OpenWorkbooks {
    <#
        実行中の Excel に接続し、開いているブック名の一覧を返す。
        Excel が起動していない場合は空配列を返す。
    #>
    $result = @()
    try {
        # 既に起動している Excel インスタンスに接続
        $script:excelApp = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        foreach ($wb in $script:excelApp.Workbooks) {
            $result += $wb.Name
        }
    }
    catch {
        # GetActiveObject 失敗 = Excel 未起動
        $script:excelApp = $null
    }
    return $result
}

# ===== ユーティリティ: セルアドレス解析("B5" -> 行2,列5) =====
function Parse-CellAddress {
    param([string]$addr)
    # 例: "B5" -> 列文字列 "B", 行 "5"
    if ($addr -notmatch '^([A-Za-z]+)([0-9]+)$') {
        return $null
    }
    $colLetters = $matches[1].ToUpper()
    $rowNum     = [int]$matches[2]

    # 列文字列を列番号(1始まり)へ変換
    $colNum = 0
    foreach ($ch in $colLetters.ToCharArray()) {
        $colNum = $colNum * 26 + ([int][char]$ch - [int][char]'A' + 1)
    }
    return [PSCustomObject]@{ Row = $rowNum; Col = $colNum }
}

# ===== ユーティリティ: 対象ブック取得 =====
function Resolve-TargetWorkbook {
    param([string]$wbName)
    if ($null -eq $script:excelApp) { return $null }
    foreach ($wb in $script:excelApp.Workbooks) {
        if ($wb.Name -eq $wbName) { return $wb }
    }
    return $null
}

# ===== ログの新規追記分を読み取り行配列で返す =====
function Read-NewLines {
    param([string]$path)

    $lines = @()
    try {
        $fs = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, `
                  [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        try {
            $currentSize = $fs.Length

            # --- ローテーション/クリア判定: サイズが基準より小さい ---
            if ($currentSize -lt $script:baseOffset) {
                # 先頭から読み直す
                $script:baseOffset = 0
            }

            # 増加が無ければ何もしない
            if ($currentSize -le $script:baseOffset) {
                $script:lastSize = $currentSize
                return @()
            }

            # 基準オフセット位置までシーク
            $fs.Seek($script:baseOffset, [System.IO.SeekOrigin]::Begin) | Out-Null
            $readLen = $currentSize - $script:baseOffset
            $buffer  = New-Object byte[] $readLen
            $read    = $fs.Read($buffer, 0, $readLen)

            # Shift-JIS でデコード
            $text = $script:sjis.GetString($buffer, 0, $read)

            # --- 完全な行のみ処理(末尾が改行で終わっていない場合は最終行を保留) ---
            $endsWithNewline = $text.EndsWith("`n")
            # 改行で分割(CRLF/LF両対応)
            $rawLines = $text -replace "`r`n", "`n" -split "`n"

            if ($endsWithNewline) {
                # 末尾は分割により空要素になるので除去
                $rawLines = $rawLines[0..($rawLines.Count - 2)]
                # 読み取った全バイトを確定
                $script:baseOffset = $currentSize
            }
            else {
                # 最終行は未完(改行待ち)なので保留し、その手前までをオフセット確定
                $lastPartial = $rawLines[-1]
                $rawLines = $rawLines[0..($rawLines.Count - 2)]
                # 保留分のバイト数を差し引いてオフセット確定
                $partialBytes = $script:sjis.GetByteCount($lastPartial)
                $script:baseOffset = $currentSize - $partialBytes
            }

            $lines = @($rawLines)
            $script:lastSize = $currentSize
        }
        finally {
            $fs.Close()
            $fs.Dispose()
        }
    }
    catch {
        # 読み取り失敗時は空を返す(常駐は継続)
        return @()
    }
    return $lines
}

# ===== Excel への書き込み =====
function Write-LinesToExcel {
    param([string[]]$lines)

    if ($lines.Count -eq 0) { return }
    if ($null -eq $script:targetWorkbook) { return }

    try {
        $sheet = $script:targetWorkbook.ActiveSheet

        # --- アクティブシート切替検知 -> 書き込み位置リセット ---
        if ($sheet.Name -ne $script:lastSheetName) {
            $script:lastSheetName = $sheet.Name
            $script:writeRow = 0
        }

        foreach ($line in $lines) {
            $row = $script:startRow + $script:writeRow
            $cell = $sheet.Cells.Item($row, $script:startCol)
            $cell.Value2 = $line
            $script:writeRow++
            $script:writtenCount++
        }
    }
    catch {
        # ユーザーのセル手動編集中など COM 例外でも常駐を落とさない
    }
}

# ===== ポーリング処理(Timer Tick) =====
function Invoke-Poll {
    if (-not $script:isWatching) { return }

    $path = $txtLogPath.Text
    if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path $path)) {
        return
    }

    $newLines = Read-NewLines -path $path
    if ($newLines.Count -gt 0) {
        Write-LinesToExcel -lines $newLines
        Update-Status
    }
}

# ===== 状態表示更新 =====
function Update-Status {
    if ($script:isWatching) {
        $sheetName = "?"
        try { $sheetName = $script:targetWorkbook.ActiveSheet.Name } catch {}
        $lblStatus.Text = "監視中: ON  /  書込先シート: $sheetName  /  累計: $($script:writtenCount) 行"
        $lblStatus.ForeColor = [System.Drawing.Color]::Green
    }
    else {
        $lblStatus.Text = "監視: OFF"
        $lblStatus.ForeColor = [System.Drawing.Color]::Gray
    }
}

# =====================================================================
#  GUI 構築
# =====================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "ログ自動転記ツール"
$form.Size = New-Object System.Drawing.Size(560, 320)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# --- 対象ブック選択 ---
$lblBook = New-Object System.Windows.Forms.Label
$lblBook.Text = "対象ブック:"
$lblBook.Location = New-Object System.Drawing.Point(20, 22)
$lblBook.Size = New-Object System.Drawing.Size(90, 22)
$form.Controls.Add($lblBook)

$cmbBook = New-Object System.Windows.Forms.ComboBox
$cmbBook.Location = New-Object System.Drawing.Point(115, 20)
$cmbBook.Size = New-Object System.Drawing.Size(300, 22)
$cmbBook.DropDownStyle = "DropDownList"
$form.Controls.Add($cmbBook)

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "更新"
$btnRefresh.Location = New-Object System.Drawing.Point(425, 19)
$btnRefresh.Size = New-Object System.Drawing.Size(100, 24)
$form.Controls.Add($btnRefresh)

# --- 監視ログファイル選択 ---
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "ログファイル:"
$lblLog.Location = New-Object System.Drawing.Point(20, 62)
$lblLog.Size = New-Object System.Drawing.Size(90, 22)
$form.Controls.Add($lblLog)

$txtLogPath = New-Object System.Windows.Forms.TextBox
$txtLogPath.Location = New-Object System.Drawing.Point(115, 60)
$txtLogPath.Size = New-Object System.Drawing.Size(300, 22)
$txtLogPath.ReadOnly = $true
$form.Controls.Add($txtLogPath)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "参照..."
$btnBrowse.Location = New-Object System.Drawing.Point(425, 59)
$btnBrowse.Size = New-Object System.Drawing.Size(100, 24)
$form.Controls.Add($btnBrowse)

# --- 開始セル入力 ---
$lblCell = New-Object System.Windows.Forms.Label
$lblCell.Text = "開始セル:"
$lblCell.Location = New-Object System.Drawing.Point(20, 102)
$lblCell.Size = New-Object System.Drawing.Size(90, 22)
$form.Controls.Add($lblCell)

$txtCell = New-Object System.Windows.Forms.TextBox
$txtCell.Location = New-Object System.Drawing.Point(115, 100)
$txtCell.Size = New-Object System.Drawing.Size(100, 22)
$txtCell.Text = "B5"
$form.Controls.Add($txtCell)

$lblCellHint = New-Object System.Windows.Forms.Label
$lblCellHint.Text = "(例: B5 → ここから下方向へ1行=1セル)"
$lblCellHint.Location = New-Object System.Drawing.Point(225, 102)
$lblCellHint.Size = New-Object System.Drawing.Size(300, 22)
$lblCellHint.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblCellHint)

# --- ON/OFF トグル ---
$btnToggle = New-Object System.Windows.Forms.Button
$btnToggle.Text = "監視 開始 (OFF)"
$btnToggle.Location = New-Object System.Drawing.Point(115, 145)
$btnToggle.Size = New-Object System.Drawing.Size(410, 40)
$btnToggle.BackColor = [System.Drawing.Color]::LightGray
$btnToggle.Font = New-Object System.Drawing.Font("Yu Gothic UI", 11, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnToggle)

# --- 状態表示 ---
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "監視: OFF"
$lblStatus.Location = New-Object System.Drawing.Point(20, 205)
$lblStatus.Size = New-Object System.Drawing.Size(505, 24)
$lblStatus.ForeColor = [System.Drawing.Color]::Gray
$lblStatus.Font = New-Object System.Drawing.Font("Yu Gothic UI", 10)
$form.Controls.Add($lblStatus)

# --- 補足説明 ---
$lblNote = New-Object System.Windows.Forms.Label
$lblNote.Text = "※ ON 以降に追記された行のみ転記します。シートを切り替えると開始セルから再開します。"
$lblNote.Location = New-Object System.Drawing.Point(20, 235)
$lblNote.Size = New-Object System.Drawing.Size(505, 40)
$lblNote.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblNote)

# --- ポーリング用 Timer ---
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1000   # 1秒
$timer.Add_Tick({ Invoke-Poll })

# =====================================================================
#  イベントハンドラ
# =====================================================================

# ブック一覧 更新
$btnRefresh.Add_Click({
    $cmbBook.Items.Clear()
    $books = Get-OpenWorkbooks
    if ($books.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "起動中の Excel ブックが見つかりません。Excel でブックを開いてから「更新」してください。",
            "情報", "OK", "Information") | Out-Null
        return
    }
    foreach ($b in $books) { $cmbBook.Items.Add($b) | Out-Null }
    $cmbBook.SelectedIndex = 0
})

# ログファイル参照
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "ログ/テキスト (*.log;*.txt)|*.log;*.txt|すべてのファイル (*.*)|*.*"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtLogPath.Text = $dlg.FileName
    }
})

# ON/OFF トグル
$btnToggle.Add_Click({
    if (-not $script:isWatching) {
        # ---- OFF -> ON ----
        # 入力チェック
        if ([string]::IsNullOrWhiteSpace($cmbBook.Text)) {
            [System.Windows.Forms.MessageBox]::Show("対象ブックを選択してください。","入力エラー","OK","Warning") | Out-Null
            return
        }
        if ([string]::IsNullOrWhiteSpace($txtLogPath.Text) -or -not (Test-Path $txtLogPath.Text)) {
            [System.Windows.Forms.MessageBox]::Show("有効なログファイルを指定してください。","入力エラー","OK","Warning") | Out-Null
            return
        }
        $cellInfo = Parse-CellAddress $txtCell.Text
        if ($null -eq $cellInfo) {
            [System.Windows.Forms.MessageBox]::Show("開始セルの形式が不正です(例: B5)。","入力エラー","OK","Warning") | Out-Null
            return
        }

        # 対象ブック解決
        $script:targetWorkbook = Resolve-TargetWorkbook $cmbBook.Text
        if ($null -eq $script:targetWorkbook) {
            [System.Windows.Forms.MessageBox]::Show("対象ブックが見つかりません。「更新」で一覧を取り直してください。","エラー","OK","Error") | Out-Null
            return
        }

        # 書き込み位置の初期化
        $script:startRow = $cellInfo.Row
        $script:startCol = $cellInfo.Col
        $script:writeRow = 0
        $script:writtenCount = 0
        try { $script:lastSheetName = $script:targetWorkbook.ActiveSheet.Name } catch { $script:lastSheetName = "" }

        # ON 時点のファイルサイズを基準オフセットに(以前の既存ログは書かない)
        try {
            $fi = New-Object System.IO.FileInfo($txtLogPath.Text)
            $script:baseOffset = $fi.Length
            $script:lastSize   = $fi.Length
        }
        catch {
            $script:baseOffset = 0
            $script:lastSize   = 0
        }

        # 状態切替
        $script:isWatching = $true
        $btnToggle.Text = "監視 停止 (ON)"
        $btnToggle.BackColor = [System.Drawing.Color]::LightGreen
        # ON 中は設定変更を禁止
        $cmbBook.Enabled = $false
        $btnRefresh.Enabled = $false
        $txtCell.Enabled = $false
        $btnBrowse.Enabled = $false
        $timer.Start()
        Update-Status
    }
    else {
        # ---- ON -> OFF ----
        $script:isWatching = $false
        $timer.Stop()
        $btnToggle.Text = "監視 開始 (OFF)"
        $btnToggle.BackColor = [System.Drawing.Color]::LightGray
        $cmbBook.Enabled = $true
        $btnRefresh.Enabled = $true
        $txtCell.Enabled = $true
        $btnBrowse.Enabled = $true
        Update-Status
    }
})

# フォーム終了時: Timer 停止と COM 解放
$form.Add_FormClosing({
    $timer.Stop()
    $timer.Dispose()
    # COM オブジェクト解放(Excel 本体は閉じない)
    if ($null -ne $script:targetWorkbook) {
        [Runtime.InteropServices.Marshal]::ReleaseComObject($script:targetWorkbook) | Out-Null
        $script:targetWorkbook = $null
    }
    if ($null -ne $script:excelApp) {
        [Runtime.InteropServices.Marshal]::ReleaseComObject($script:excelApp) | Out-Null
        $script:excelApp = $null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
})

# 起動時に一度ブック一覧を取得
$btnRefresh.PerformClick()

# フォーム表示
[void]$form.ShowDialog()
