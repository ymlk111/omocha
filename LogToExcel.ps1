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
    - グローバルホットキー(Ctrl+Alt+L)で ON/OFF をトグル(アプリ非アクティブでも有効)

.NOTES
    実行環境: Windows + Excel インストール済み / PowerShell
#>

# ===== アセンブリ読み込み =====
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ===== グローバルホットキー受信用クラス =====
Add-Type -ReferencedAssemblies System.Windows.Forms @"
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public class HotkeyWindow : NativeWindow {
    public event Action HotkeyPressed;          // ホットキー押下時に発火
    private const int WM_HOTKEY = 0x0312;
    public int HotkeyId = 9000;

    [DllImport("user32.dll")] public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
    [DllImport("user32.dll")] public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    public HotkeyWindow(IntPtr handle) { this.AssignHandle(handle); }   // 既存フォームの窓に相乗り
    public bool Register(uint modifiers, uint vk) { return RegisterHotKey(this.Handle, HotkeyId, modifiers, vk); }
    public void Unregister() { UnregisterHotKey(this.Handle, HotkeyId); }

    protected override void WndProc(ref Message m) {
        // WM_HOTKEY を捕捉してイベント発火
        if (m.Msg == WM_HOTKEY && (int)m.WParam == HotkeyId && HotkeyPressed != null) {
            HotkeyPressed();
        }
        base.WndProc(ref m);
    }
}
"@

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
$script:hotkey          = $null       # グローバルホットキー

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

            # --- 生バイト列で最後の改行(LF=0x0A)位置を探す ---
            # 完全な行の末尾までだけをデコード対象にすることで、
            # マルチバイト文字の途中で切れて文字化け/オフセットずれが起きるのを防ぐ。
            # (Shift-JIS の2バイト目に 0x0A は出現しないため 0x0A 走査は安全)
            $lastNL = -1
            for ($i = $read - 1; $i -ge 0; $i--) {
                if ($buffer[$i] -eq 0x0A) { $lastNL = $i; break }
            }

            # 完全な行がまだ無ければ何も確定せず次回に持ち越し(改行待ち)
            if ($lastNL -lt 0) {
                $script:lastSize = $currentSize
                return @()
            }

            # 改行までの完全な行のみ Shift-JIS でデコード(必ず文字境界で切れる)
            $text = $script:sjis.GetString($buffer, 0, $lastNL + 1)

            # 消費した生バイト数だけオフセットを進める(残り半端なバイトは次回へ持ち越し)
            $script:baseOffset = $script:baseOffset + ($lastNL + 1)

            # 改行で分割(CRLF/LF両対応)。末尾の空要素を除去
            $rawLines = ($text -replace "`r`n", "`n") -split "`n"
            $rawLines = $rawLines[0..($rawLines.Count - 2)]

            $lines = @($rawLines)
            $script:lastSize = $currentSize
        }
        finally {
            $fs.Close()
            $fs.Dispose()
        }
    }
 catch {
        # 一時的にエラー内容を表示(原因切り分け用)
        [System.Windows.Forms.MessageBox]::Show("読み取り失敗: $($_.Exception.Message)","DEBUG","OK","Error") | Out-Null
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
$lblNote.Text = "※ ON 以降に追記された行のみ転記します。シート切替で開始セルから再開。 ホットキー: Ctrl+Alt+L"
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

# フォーム終了時: Timer 停止・ホットキー解除・COM 解放
$form.Add_FormClosing({
    $timer.Stop()
    $timer.Dispose()
    # グローバルホットキー解除(登録解除しないと OS 側に残る)
    if ($null -ne $script:hotkey) {
        $script:hotkey.Unregister()
        $script:hotkey.ReleaseHandle()
        $script:hotkey = $null
    }
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

# グローバルホットキー登録(アプリ非アクティブでも有効) 既定: Ctrl + Alt + L
$form.Add_Shown({
    # ShowDialog 後はウィンドウハンドルが確定しているのでここで登録
    $script:hotkey = New-Object HotkeyWindow($form.Handle)
    # UI スレッド上で実行されるためトグル(COM操作)も安全
    $script:hotkey.add_HotkeyPressed({ $btnToggle.PerformClick() })
    # 修飾: MOD_ALT(0x1) + MOD_CONTROL(0x2) + MOD_NOREPEAT(0x4000) / 仮想キー: L=0x4C
    $ok = $script:hotkey.Register((0x1 -bor 0x2 -bor 0x4000), 0x4C)
    if (-not $ok) {
        [System.Windows.Forms.MessageBox]::Show(
            "グローバルホットキー(Ctrl+Alt+L)の登録に失敗しました。他アプリと競合している可能性があります。",
            "警告","OK","Warning") | Out-Null
    }
})

# 起動時に一度ブック一覧を取得
$btnRefresh.PerformClick()

# フォーム表示
[void]$form.ShowDialog()
