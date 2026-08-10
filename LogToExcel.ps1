<#
.SYNOPSIS
    ログファイルへの追記を監視し、新規追記行を Excel のアクティブシートへ
    1行＝1セルで自動転記する常駐 GUI ツール。 (v2.1: 欠落対策・堅牢化版)

.DESCRIPTION
    - 監視方式: Timer ポーリング(1秒)でファイルサイズの変化を検知
    - ON 時点のファイルサイズを基準オフセットとし、以降の追記分のみ転記
    - ローテーション/クリア(サイズ減少)時は先頭から読み直し
    - 書き込み先は GUI で指定したブックのアクティブシート
    - シートごとに書き込み位置を保持(切替で戻っても以前の転記を上書きしない)
    - 文字コードは 自動判定(UTF-8→Shift-JIS) / UTF-8固定 / Shift-JIS固定 を選択可
    - グローバルホットキー(Ctrl+Alt+L)で ON/OFF をトグル(アプリ非アクティブでも有効)
    - ［半角数字］で始まる行を無視するモード

.NOTES
    実行環境: Windows + Excel インストール済み / Windows PowerShell 5.1 (powershell.exe)
              ※ PowerShell 7 (pwsh.exe) は非対応 (Marshal.GetActiveObject が存在しないため)
    ファイルは UTF-8 (BOM付き) で保存すること (powershell.exe の文字化け防止)

    v2.1 の主な修正点:
      [欠落対策(主因)]
      - 保留キュー方式: 読み取った行はキューに保持し、Excel への書き込みが
        成功した分だけ消費する。busy(セル編集中/ドラッグ中: 0x800AC472,
        0x80010001, 0x8001010A 等)で失敗しても次 Tick で自動再試行 → 欠落しない
      - 1セル上限(32,767文字)超の行は複数セルに分割して全文を保持
        (超過行は下方向に複数セルへまたがり、以降の行はその分だけ下にずれる)
      - NumberFormat="@" を設定してから範囲一括書き込み
        ("=..." の数式解釈・日付/数値の自動変換を防ぎ、COM呼び出し回数も激減)
      [重要度: 高]
      - Test-Path を -LiteralPath 化 (パスに [ ] を含むと監視されないバグの修正)
      - PowerShell 7 実行時は起動時にエラー表示して終了
      - ブック閉鎖 / Excel 終了(切断系エラー)を検知して自動停止+通知
      - 保留キューに上限(超過分は古い方から破棄し件数を表示)
      - 開始セルの範囲検証 / シート最終行(1,048,576)到達で安全停止
      [重要度: 中]
      - lastSize を利用したローテーション検知の強化(縮小→再肥大化も捕捉)
      - シート名→書き込み位置の辞書化(切替で戻っても続きから書く)
      - 文字コードの固定設定を追加 / 1Tick の読み取り量に上限(巨大追記対策)
      - 読み取り失敗の連続回数をステータスに表示
#>

# ===== アセンブリ読み込み =====
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ===== 実行環境チェック (Windows PowerShell 5.1 専用) =====
if ($PSVersionTable.PSEdition -eq 'Core') {
    [System.Windows.Forms.MessageBox]::Show(
        "このツールは Windows PowerShell 5.1 (powershell.exe) で実行してください。`n" +
        "PowerShell 7 (pwsh.exe) では Excel への接続 (Marshal.GetActiveObject) が利用できません。",
        "実行環境エラー", "OK", "Error") | Out-Null
    return
}

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

# ===== 定数 =====
$EXCEL_CELL_MAX = 32767        # Excel の 1セル最大文字数(超過行は複数セルに分割)
$EXCEL_MAX_ROW  = 1048576      # ワークシート最大行
$EXCEL_MAX_COL  = 16384        # ワークシート最大列 (XFD)
$MAX_READ_BYTES = 4MB          # 1Tick あたりの読み取り上限(巨大追記時のメモリ保護)
$MAX_PENDING    = 100000       # 保留キューの上限行数(超過分は古い方から破棄)

# ===== グローバル状態 =====
$script:isWatching     = $false      # 監視 ON/OFF
$script:baseOffset     = 0           # 基準バイトオフセット(ここまで消費済み)
$script:lastSize       = 0           # 前回確認時のファイルサイズ(ローテ検知に使用)
$script:startCol       = 0           # 書き込み開始列(1始まり)
$script:startRow       = 0           # 書き込み開始行(1始まり)
$script:sheetRows      = @{}         # シート名 -> 次に書く行オフセット(開始セル基準)
$script:targetWorkbook = $null       # 対象ブックの COM オブジェクト
$script:excelApp       = $null       # Excel.Application の COM オブジェクト
$script:writtenCount   = 0           # 累計書き込みセル数
$script:pending        = New-Object System.Collections.Generic.Queue[string]  # 未書込行(書込成功まで保持)
$script:lastError      = ""          # 直近の書き込みエラー(再試行中の表示用)
$script:flushFailCount = 0           # 書き込み連続失敗回数
$script:readErrCount   = 0           # 読み取り連続失敗回数
$script:droppedCount   = 0           # キューあふれで破棄した行数
$script:inPoll         = $false      # ポーリング再入防止フラグ
$script:hotkey         = $null       # グローバルホットキー

# エンコーディング
$script:encUtf8        = New-Object System.Text.UTF8Encoding($false, $true)   # 不正バイトで例外(自動判定用)
$script:encUtf8Lenient = New-Object System.Text.UTF8Encoding($false, $false)  # 例外にしない(固定/強制確定用)
$script:encSjis        = [System.Text.Encoding]::GetEncoding("Shift_JIS")

# ===== ユーティリティ: 起動中ブック一覧の取得 =====
function Get-OpenWorkbooks {
    <#
        実行中の Excel に接続し、開いているブック名の一覧を返す。
        Excel が起動していない場合は空配列を返す。
        (注: 複数の Excel インスタンスがある場合は片方しか見えない)
    #>
    $result = @()
    try {
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

# ===== ユーティリティ: セルアドレス解析("B5" -> 行5,列2) =====
function Parse-CellAddress {
    param([string]$addr)
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

    # 範囲検証: "B0" や "ZZZ1" のような Excel に存在しない座標を弾く
    if ($rowNum -lt 1 -or $rowNum -gt $EXCEL_MAX_ROW) { return $null }
    if ($colNum -lt 1 -or $colNum -gt $EXCEL_MAX_COL) { return $null }

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

# ===== バイト列のデコード(文字コード設定に従う) =====
function Decode-Bytes {
    param([byte[]]$buffer, [int]$length, [bool]$forced = $false)
    switch ($cmbEnc.SelectedIndex) {
        1 { return $script:encUtf8Lenient.GetString($buffer, 0, $length) }   # UTF-8 固定
        2 { return $script:encSjis.GetString($buffer, 0, $length) }          # Shift-JIS 固定
        default {
            # 自動判定: UTF-8 として解釈できれば UTF-8、不正バイトがあれば Shift-JIS
            if ($forced) {
                # 強制確定チャンクは文字境界で切れている可能性があるため寛容な UTF-8 で
                return $script:encUtf8Lenient.GetString($buffer, 0, $length)
            }
            try   { return $script:encUtf8.GetString($buffer, 0, $length) }
            catch { return $script:encSjis.GetString($buffer, 0, $length) }
        }
    }
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

            # --- ローテーション/クリア判定 ---
            #  - 前回確認時よりサイズが減った: ポーリング間隔内に
            #    「クリア -> 基準オフセット以上まで再肥大化」したケースも捕捉できる
            #  - 基準オフセットより小さい: 従来判定
            if ($currentSize -lt $script:lastSize -or $currentSize -lt $script:baseOffset) {
                # 先頭から読み直す(ローテ後の内容はすべて新規扱い)
                $script:baseOffset = 0
            }

            # 増加が無ければ何もしない
            if ($currentSize -le $script:baseOffset) {
                $script:lastSize = $currentSize
                $script:readErrCount = 0
                return @()
            }

            # 基準オフセット位置までシーク
            $fs.Seek($script:baseOffset, [System.IO.SeekOrigin]::Begin) | Out-Null
            $readLen = $currentSize - $script:baseOffset
            if ($readLen -gt $MAX_READ_BYTES) {
                # 1Tick の読み取り量に上限(残りは次 Tick 以降で処理)
                # ※ Int32 超の巨大追記で Read が破綻するエッジも同時に解消
                $readLen = $MAX_READ_BYTES
            }
            $buffer = New-Object byte[] ([int]$readLen)
            $read   = $fs.Read($buffer, 0, [int]$readLen)

            # --- 生バイト列で最後の改行(LF=0x0A)位置を探す ---
            # 完全な行の末尾までだけをデコード対象にすることで、
            # マルチバイト文字の途中で切れて文字化け/オフセットずれが起きるのを防ぐ。
            # (UTF-8/Shift-JIS とも 2バイト目以降に 0x0A は出現しないため 0x0A 走査は安全)
            # ※ 改行が CR のみのファイルは行を確定できない(非対応)
            $lastNL = -1
            for ($i = $read - 1; $i -ge 0; $i--) {
                if ($buffer[$i] -eq 0x0A) { $lastNL = $i; break }
            }

            $forced = $false
            if ($lastNL -lt 0) {
                if ($read -ge $MAX_READ_BYTES) {
                    # 上限まで読んでも改行が無い超巨大行:
                    # 恒久スタックを避けるためチャンク全体を強制的に1行として確定する。
                    # (文字境界で切れる可能性はあるが、セル分割で全文は保持される)
                    $lastNL = $read - 1
                    $forced = $true
                }
                else {
                    # 完全な行がまだ無ければ何も確定せず次回に持ち越し(改行待ち)
                    $script:lastSize = $currentSize
                    $script:readErrCount = 0
                    return @()
                }
            }

            # 改行までの完全な行のみデコード(必ず文字境界で切れる)
            $decodeLen = $lastNL + 1
            $text = Decode-Bytes -buffer $buffer -length $decodeLen -forced $forced
            # 先頭に UTF-8 BOM が残った場合は除去
            $text = $text.TrimStart([char]0xFEFF)

            # 消費した生バイト数だけオフセットを進める(残り半端なバイトは次回へ持ち越し)
            $script:baseOffset = $script:baseOffset + $decodeLen

            # 改行で分割(CRLF/LF両対応)
            $rawLines = ($text -replace "`r`n", "`n") -split "`n"
            if (-not $forced) {
                # 末尾は必ず改行で終わっているため、最後の空要素を除去
                $rawLines = $rawLines[0..($rawLines.Count - 2)]
            }

            $lines = @($rawLines)
            $script:lastSize = $currentSize
            $script:readErrCount = 0
        }
        finally {
            $fs.Close()
            $fs.Dispose()
        }
    }
    catch {
        # 読み取り失敗時は空を返す(常駐は継続)。連続失敗はステータスに表示
        $script:readErrCount++
        return @()
    }
    return $lines
}

# ===== Excel への書き込み(保留キューを一括転記) =====
function Flush-PendingToExcel {
    if ($script:pending.Count -eq 0) { return }
    if ($null -eq $script:targetWorkbook) { return }

    try {
        $sheet     = $script:targetWorkbook.ActiveSheet
        $sheetName = $sheet.Name

        # --- シートごとの書き込み位置を取得(初出シートは開始セルから) ---
        # 以前のシートに戻っても続きから書くため、過去の転記を上書きしない
        if (-not $script:sheetRows.ContainsKey($sheetName)) {
            $script:sheetRows[$sheetName] = 0
        }
        $top = $script:startRow + $script:sheetRows[$sheetName]

        # --- 最終行チェック ---
        if ($top -gt $EXCEL_MAX_ROW) {
            Stop-Watching "シートの最終行($EXCEL_MAX_ROW)に達したため監視を停止しました。(シート: $sheetName)"
            return
        }
        $n   = $script:pending.Count
        $fit = [int][Math]::Min($n, $EXCEL_MAX_ROW - $top + 1)

        # --- 2次元配列を組み立てて範囲一括書き込み ---
        # 1セルずつの COM 呼び出しをやめることで大幅に高速化し、
        # 書き込み中にユーザー操作と衝突する時間窓も最小化する
        $data = [Array]::CreateInstance([object], $fit, 1)
        $i = 0
        foreach ($line in $script:pending) {
            if ($i -ge $fit) { break }
            $data[$i, 0] = $line
            $i++
        }

        $range = $sheet.Range($sheet.Cells.Item($top, $script:startCol),
                              $sheet.Cells.Item($top + $fit - 1, $script:startCol))
        $range.NumberFormat = "@"     # 文字列書式: "=..." の数式解釈や日付/数値の自動変換を防ぐ
        $range.Value2 = $data

        # ---- ここまで成功して初めて「消費」する(失敗時は保留のまま再試行) ----
        for ($i = 0; $i -lt $fit; $i++) { $null = $script:pending.Dequeue() }
        $script:sheetRows[$sheetName] = $script:sheetRows[$sheetName] + $fit
        $script:writtenCount += $fit
        $script:lastError = ""
        $script:flushFailCount = 0

        if ($fit -lt $n) {
            # 最終行までで書ききれなかった(以降は書けないため安全停止)
            Stop-Watching "シートの最終行に達したため監視を停止しました。未転記 $($script:pending.Count) 行が残っています。"
        }
    }
    catch {
        # busy 系 (0x800AC472 / 0x80010001 / 0x8001010A: セル編集中・ドラッグ中・
        # ダイアログ表示中など) は保留キューに残したまま次 Tick で自動再試行 → 欠落しない
        $script:lastError = $_.Exception.Message
        $script:flushFailCount++

        # 切断系エラー(ブックが閉じられた / Excel が終了した)は再試行しても
        # 回復しないため、検知したら自動停止して通知する
        $hrHex  = "{0:X8}" -f $_.Exception.HResult
        $isDead = ($_.Exception -is [Runtime.InteropServices.InvalidComObjectException]) -or
                  (@("80010108", "800706BA", "800401FD") -contains $hrHex)
                  # 80010108: RPC_E_DISCONNECTED / 800706BA: RPCサーバー利用不可 /
                  # 800401FD: オブジェクトがサーバーに未接続
        if ($isDead) {
            Stop-Watching "対象ブックまたは Excel との接続が失われたため監視を停止しました。`n($($script:lastError))"
        }
    }
}

# ===== ポーリング処理(Timer Tick) =====
function Invoke-Poll {
    if ($script:inPoll) { return }   # 再入防止(COM 呼び出し中の Tick 再入対策)
    $script:inPoll = $true
    try {
        if (-not $script:isWatching) { return }

        $path = $txtLogPath.Text
        # -LiteralPath: パスに [ ] が含まれるとワイルドカード解釈で常に false になるため必須
        if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path)) {
            return
        }

        $newLines = Read-NewLines -path $path
        if ($newLines.Count -gt 0 -and $chkIgnore.Checked) {
            # ［半角数字］で始まる行を無視するモード(先頭のタブ/空白は許容)
            $newLines = @($newLines | Where-Object { $_ -notmatch '^\s*\[[0-9]+\]' })
        }

        foreach ($l in $newLines) {
            if ($l.Length -le $EXCEL_CELL_MAX) {
                $script:pending.Enqueue($l)
            }
            else {
                # セル上限(32,767文字)超の行は複数セルに分割して全文を保持する
                # (超過行は下方向に複数セルへまたがる)
                $pos = 0
                while ($pos -lt $l.Length) {
                    $len = [Math]::Min($EXCEL_CELL_MAX, $l.Length - $pos)
                    # サロゲートペア(絵文字等)の途中で切らない
                    if ($len -lt ($l.Length - $pos) -and
                        [char]::IsHighSurrogate($l[$pos + $len - 1])) {
                        $len--
                    }
                    $script:pending.Enqueue($l.Substring($pos, $len))
                    $pos += $len
                }
            }
        }

        # 保留キュー上限: Excel へ長時間書けない場合のメモリ保護(古い方から破棄)
        while ($script:pending.Count -gt $MAX_PENDING) {
            $null = $script:pending.Dequeue()
            $script:droppedCount++
        }

        if ($script:pending.Count -gt 0) {
            Flush-PendingToExcel
        }
        Update-Status
    }
    finally {
        $script:inPoll = $false
    }
}

# ===== 状態表示更新 =====
function Update-Status {
    if (-not $script:isWatching) {
        $lblStatus.Text = "監視: OFF"
        $lblStatus.ForeColor = [System.Drawing.Color]::Gray
        return
    }

    $pos = "?"
    try {
        $sn  = $script:targetWorkbook.ActiveSheet.Name
        $off = 0
        if ($script:sheetRows.ContainsKey($sn)) { $off = $script:sheetRows[$sn] }
        $pos = "$sn (次: $($script:startRow + $off) 行目)"
    } catch {}

    $msg = "監視中: ON / 書込先: $pos / 累計: $($script:writtenCount) セル"
    if ($script:pending.Count -gt 0) { $msg += " / 保留: $($script:pending.Count) 行" }
    if ($script:droppedCount -gt 0)  { $msg += " / あふれ破棄: $($script:droppedCount) 行" }
    if ($script:readErrCount -gt 0)  { $msg += " / 読取失敗: 連続 $($script:readErrCount) 回" }

    if ($script:lastError) {
        $msg += "`n書込リトライ中($($script:flushFailCount)回目): $($script:lastError)"
        $lblStatus.ForeColor = [System.Drawing.Color]::DarkOrange
    }
    else {
        $lblStatus.ForeColor = [System.Drawing.Color]::Green
    }
    $lblStatus.Text = $msg
}

# ===== 監視開始 =====
function Start-Watching {
    # ---- 入力チェック ----
    if ([string]::IsNullOrWhiteSpace($cmbBook.Text)) {
        [System.Windows.Forms.MessageBox]::Show("対象ブックを選択してください。","入力エラー","OK","Warning") | Out-Null
        return
    }
    if ([string]::IsNullOrWhiteSpace($txtLogPath.Text) -or -not (Test-Path -LiteralPath $txtLogPath.Text)) {
        [System.Windows.Forms.MessageBox]::Show("有効なログファイルを指定してください。","入力エラー","OK","Warning") | Out-Null
        return
    }
    $cellInfo = Parse-CellAddress $txtCell.Text
    if ($null -eq $cellInfo) {
        [System.Windows.Forms.MessageBox]::Show(
            "開始セルの形式が不正です(例: B5)。`n行は 1～$EXCEL_MAX_ROW、列は A～XFD の範囲で指定してください。",
            "入力エラー","OK","Warning") | Out-Null
        return
    }

    # ---- 対象ブック解決 ----
    $script:targetWorkbook = Resolve-TargetWorkbook $cmbBook.Text
    if ($null -eq $script:targetWorkbook) {
        [System.Windows.Forms.MessageBox]::Show("対象ブックが見つかりません。「更新」で一覧を取り直してください。","エラー","OK","Error") | Out-Null
        return
    }

    # ---- 状態初期化 ----
    $script:startRow       = $cellInfo.Row
    $script:startCol       = $cellInfo.Col
    $script:sheetRows      = @{}
    $script:pending.Clear()
    $script:writtenCount   = 0
    $script:droppedCount   = 0
    $script:readErrCount   = 0
    $script:flushFailCount = 0
    $script:lastError      = ""

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

    # ---- 状態切替 ----
    $script:isWatching = $true
    $btnToggle.Text = "監視 停止 (ON)"
    $btnToggle.BackColor = [System.Drawing.Color]::LightGreen
    # ON 中は設定変更を禁止
    $cmbBook.Enabled = $false
    $btnRefresh.Enabled = $false
    $txtCell.Enabled = $false
    $btnBrowse.Enabled = $false
    $cmbEnc.Enabled = $false
    $timer.Start()
    Update-Status
}

# ===== 監視停止 =====
function Stop-Watching {
    param([string]$reason = "")   # 自動停止時の理由(空なら通常のOFF操作)

    $script:isWatching = $false
    $timer.Stop()
    $btnToggle.Text = "監視 開始 (OFF)"
    $btnToggle.BackColor = [System.Drawing.Color]::LightGray
    $cmbBook.Enabled = $true
    $btnRefresh.Enabled = $true
    $txtCell.Enabled = $true
    $btnBrowse.Enabled = $true
    $cmbEnc.Enabled = $true
    Update-Status

    if ($reason) {
        [System.Windows.Forms.MessageBox]::Show($reason, "監視停止", "OK", "Warning") | Out-Null
    }
}

# =====================================================================
#  GUI 構築
# =====================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "ログ自動転記ツール v2.1"
$form.Size = New-Object System.Drawing.Size(560, 400)
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

# --- 文字コード選択 ---
$lblEnc = New-Object System.Windows.Forms.Label
$lblEnc.Text = "文字コード:"
$lblEnc.Location = New-Object System.Drawing.Point(20, 136)
$lblEnc.Size = New-Object System.Drawing.Size(90, 22)
$form.Controls.Add($lblEnc)

$cmbEnc = New-Object System.Windows.Forms.ComboBox
$cmbEnc.Location = New-Object System.Drawing.Point(115, 134)
$cmbEnc.Size = New-Object System.Drawing.Size(220, 22)
$cmbEnc.DropDownStyle = "DropDownList"
[void]$cmbEnc.Items.Add("自動判定 (UTF-8 → Shift-JIS)")
[void]$cmbEnc.Items.Add("UTF-8 固定")
[void]$cmbEnc.Items.Add("Shift-JIS 固定")
$cmbEnc.SelectedIndex = 0
$form.Controls.Add($cmbEnc)

# --- ［半角数字］始まり行を無視するモード ---
$chkIgnore = New-Object System.Windows.Forms.CheckBox
$chkIgnore.Text = "［半角数字］で始まる行を無視する (先頭のタブ/空白は許容)"
$chkIgnore.Location = New-Object System.Drawing.Point(20, 164)
$chkIgnore.Size = New-Object System.Drawing.Size(505, 22)
$chkIgnore.Checked = $false
$form.Controls.Add($chkIgnore)

# --- ON/OFF トグル ---
$btnToggle = New-Object System.Windows.Forms.Button
$btnToggle.Text = "監視 開始 (OFF)"
$btnToggle.Location = New-Object System.Drawing.Point(115, 194)
$btnToggle.Size = New-Object System.Drawing.Size(410, 40)
$btnToggle.BackColor = [System.Drawing.Color]::LightGray
$btnToggle.Font = New-Object System.Drawing.Font("Yu Gothic UI", 11, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnToggle)

# --- 状態表示(2行分: 2行目は再試行中のエラー表示) ---
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "監視: OFF"
$lblStatus.Location = New-Object System.Drawing.Point(20, 244)
$lblStatus.Size = New-Object System.Drawing.Size(505, 44)
$lblStatus.ForeColor = [System.Drawing.Color]::Gray
$lblStatus.Font = New-Object System.Drawing.Font("Yu Gothic UI", 9)
$form.Controls.Add($lblStatus)

# --- 補足説明 ---
$lblNote = New-Object System.Windows.Forms.Label
$lblNote.Text = "※ ON 以降に追記された行のみ転記します。シートごとに書き込み位置を保持し、" + "`n" +
                "    書き込み失敗時は自動で再試行します。 ホットキー: Ctrl+Alt+L"
$lblNote.Location = New-Object System.Drawing.Point(20, 292)
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
        Start-Watching
    }
    else {
        Stop-Watching
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
            "グローバルホットキー(Ctrl+Alt+L)の登録に失敗しました。他アプリと競合している可能性があります。`n(このツールを二重起動している場合もこの警告が出ます)",
            "警告","OK","Warning") | Out-Null
    }
})

# 起動時に一度ブック一覧を取得
$btnRefresh.PerformClick()

# フォーム表示
[void]$form.ShowDialog()
