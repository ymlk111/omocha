<#
.SYNOPSIS
    CloudWatch Logs をキーワードで検索し、結果を CSV に出力します（画面あり・なしの両対応）。

.DESCRIPTION
    MFA などで AWS CLI の認証が済んでいる前提です。期間はすべて JST で扱います。

    起動方法は 2 通りあります。
      ・画面あり : 引数なしで実行する（または -Gui を付ける）と入力画面が開きます。
      ・画面なし : -LogGroup を指定すると、そのまま検索して CSV を出力します。

    検索条件（-Keyword）は省略できます。省略すると、期間内のログを全件取得します。
    結果は何件あっても 1 つの CSV に時刻順で出力します（途中で中断・失敗した場合は .partial ファイルに途中まで残ります）。

    1 回のクエリは limit 10000 で実行します。上限に達した場合は、取得済みの 10,000 件を活かしたまま
    「最後に取得した時刻」から続きを取得します（残り件数から分割数を決め、重複は自動で除外）。
    外部モジュールは使いません（PowerShell 標準コマンドと .NET、aws コマンドのみ。画面は .NET 標準の Windows Forms）。

.PARAMETER LogGroup
    検索対象のロググループ名。カンマ区切りで複数指定できます（最大 50）。

.PARAMETER Keyword
    検索条件。省略すると期間内の全件を取得します。
    カンマ区切りで複数指定できます。既定は AND（すべて含む）。-Or を付けると OR（いずれかを含む）。
    大文字・小文字は区別されます。

.PARAMETER Days
    JST の日付基準で、今日を含めて何日分さかのぼるか。
    1 = 今日 0:00 から現在まで / 3 = 一昨日 0:00 から現在まで。既定は 1、最大 31。

.PARAMETER From
    開始日時（JST）。指定すると -Days より優先されます。例: 2026-10-01 または "2026-10-01 13:30"

.PARAMETER To
    終了日時（JST）。-From と一緒に使います。日付だけの場合はその日の終わりまで。省略時は現在まで。

.PARAMETER Or
    複数の検索条件を OR で結合します。

.PARAMETER OutFile
    出力 CSV のパス。省略時はカレントフォルダに cwlogs_yyyyMMdd_HHmmss.csv を作成します。

.PARAMETER AwsProfile
    使用する AWS CLI のプロファイル名（省略時は既定のプロファイル／環境変数）。

.PARAMETER Region
    リージョン（省略時は AWS CLI の設定に従います）。

.PARAMETER Gui
    入力画面を開きます。ほかの引数を一緒に指定すると、その値が画面の初期値になります。

.EXAMPLE
    .\Search-CwLogs.ps1
    入力画面を開きます。

.EXAMPLE
    .\Search-CwLogs.ps1 -Gui -LogGroup /app/web,/app/batch -AwsProfile dev
    ロググループとプロファイルを入力済みの状態で画面を開きます。

.EXAMPLE
    .\Search-CwLogs.ps1 -LogGroup /aws/lambda/order-api -Keyword "req-12345" -Days 3
    画面なしで、直近 3 日分を検索します。

.EXAMPLE
    .\Search-CwLogs.ps1 -LogGroup /app/web -Keyword "timeout","refused" -Or -From 2026-10-01 -To 2026-10-02 -OutFile C:\work\result.csv

.EXAMPLE
    .\Search-CwLogs.ps1 -LogGroup /app/web -From "2026-10-05 13:00" -To "2026-10-05 14:00"
    検索条件なしで、その 1 時間のログを全件取得します。
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]]$LogGroup,

    [Parameter(Position = 1)]
    [string[]]$Keyword,

    [ValidateRange(1, 31)]
    [int]$Days = 1,

    [string]$From,

    [string]$To,

    [switch]$Or,

    [string]$OutFile,

    [string]$AwsProfile,

    [string]$Region,

    [switch]$Gui
)

$ErrorActionPreference = 'Stop'

# ---- 設定値 -----------------------------------------------------------------
$JstOffset     = [TimeSpan]::FromHours(9)
$MaxRows       = 10000   # 1 回のクエリの limit
$TargetRows    = 5000    # 上限超過後に続きを分割取得するとき、1 クエリあたりに狙う件数（偏りに備えて上限の半分）
$MaxRangeDays  = 31      # スキャン課金の事故防止。これを超える期間はエラーにする
$AwsTimeoutSec = 300     # aws コマンド 1 回あたりの待ち時間の上限
$Invariant     = [Globalization.CultureInfo]::InvariantCulture
$Utf8NoBom     = New-Object System.Text.UTF8Encoding($false)
$Utf8Bom       = New-Object System.Text.UTF8Encoding($true)
$EpochUtc      = New-Object DateTime(1970, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)

# ---- 実行中の状態 -----------------------------------------------------------
$script:Ui              = $null    # 画面の部品（画面ありのときだけ入る）
$script:Running         = $false   # 画面で検索を実行中か
$script:CancelRequested = $false   # 画面の「中止」が押されたか
$script:IgnoreCancel    = $false   # 後始末の間は中止を無視する
$script:AwsPath         = $null
$script:AwsEncoding     = $Utf8NoBom
$script:AwsUtf8Output   = $false
$script:CommonArgs      = @()
$script:LogGroups       = @()
$script:QueryString     = ''
$script:RequestFile     = $null
$script:ActiveQueryId   = $null
$script:BytesScanned    = [double]0
$script:QueryCount      = 0
$script:Truncated       = $false
$script:LastOutPath     = $null
$script:Writer          = $null    # CSV の書き込み先（取得しながら順次書き出す）
$script:RowCount        = 0        # CSV に書き出した件数
$script:LastProgress    = [DateTime]::UtcNow
# 続きの取得は「最後に取得した秒」から始めるため、その 1 秒分だけ前回と重なる。重複除外用に覚えておく
$script:OverlapSecond   = [long]-1
$script:OverlapKeys     = New-Object 'System.Collections.Generic.HashSet[string]'

# =============================================================================
#  共通処理（画面あり・なしの両方で使う）
# =============================================================================

# 進行状況を表示する。画面ありならログ欄へ、画面なしならコンソールへ出す
function Write-Status {
    param([string]$Text, [switch]$IsWarning)

    if ($script:Ui) {
        $prefix = ''
        if ($IsWarning) { $prefix = '【警告】' }
        $script:Ui.Log.AppendText($prefix + $Text + "`r`n")
        [System.Windows.Forms.Application]::DoEvents()
    }
    elseif ($IsWarning) {
        Write-Warning $Text
    }
    else {
        Write-Host $Text
    }
}

# 画面の再描画とボタン操作を処理し、「中止」が押されていたら処理を打ち切る
function Test-Cancel {
    if ($script:Ui) { [System.Windows.Forms.Application]::DoEvents() }
    if ($script:CancelRequested -and -not $script:IgnoreCancel) {
        throw '中止しました。'
    }
}

# 指定ミリ秒だけ待つ（待っている間も画面が固まらないようにする）
function Wait-Interval {
    param([int]$Milliseconds)

    $until = [DateTime]::UtcNow.AddMilliseconds($Milliseconds)
    while ([DateTime]::UtcNow -lt $until) {
        Test-Cancel
        Start-Sleep -Milliseconds 50
    }
}

# JST の日時文字列を DateTimeOffset に変換する。-EndOfDay 指定時、日付だけなら翌日 0:00 を返す
function ConvertFrom-JstText {
    param([string]$Text, [switch]$EndOfDay)

    $value   = $Text.Trim()
    $formats = [string[]]@(
        'yyyy-MM-dd', 'yyyy/MM/dd',
        'yyyy-MM-dd HH:mm', 'yyyy/MM/dd HH:mm',
        'yyyy-MM-dd HH:mm:ss', 'yyyy/MM/dd HH:mm:ss'
    )
    $parsed = [datetime]::MinValue
    $ok = [datetime]::TryParseExact($value, $formats, $Invariant, [Globalization.DateTimeStyles]::None, [ref]$parsed)
    if (-not $ok) {
        throw "日時の形式が不正です: '$Text'（例: 2026-10-01 または '2026-10-01 13:30'）"
    }
    $result = New-Object DateTimeOffset($parsed, $JstOffset)
    if ($EndOfDay -and $value.Length -le 10) {
        $result = $result.AddDays(1)
    }
    return $result
}

# コマンドライン引数 1 つを、空白や引用符を含んでいても正しく渡せる形にする（Windows の規則）
function ConvertTo-ArgumentText {
    param([string]$Value)

    if ($Value.Length -gt 0 -and $Value -notmatch '[\s"]') { return $Value }
    $escaped = [regex]::Replace($Value, '(\\*)"', '$1$1\"')
    $escaped = [regex]::Replace($escaped, '(\\+)$', '$1$1')
    return '"' + $escaped + '"'
}

# AWS CLI を実行し、標準出力を文字列で返す。失敗時は CLI のエラーメッセージ付きで例外にする。
# 文字コードを明示して読み取るので、コンソールの設定（chcp）に左右されない
function Invoke-Aws {
    param([string[]]$Arguments, [switch]$Bare)

    $all = @($Arguments)
    if (-not $Bare) { $all += $script:CommonArgs }

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName               = $script:AwsPath
    $startInfo.Arguments              = (@($all | ForEach-Object { ConvertTo-ArgumentText -Value $_ }) -join ' ')
    $startInfo.UseShellExecute        = $false
    $startInfo.CreateNoWindow         = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError  = $true
    $startInfo.StandardOutputEncoding = $script:AwsEncoding
    $startInfo.StandardErrorEncoding  = $script:AwsEncoding
    # 日本語を含む検索語・ログが文字化けしないよう、CLI との受け渡しを UTF-8 にそろえる
    $startInfo.EnvironmentVariables['AWS_CLI_FILE_ENCODING'] = 'UTF-8'
    if ($script:AwsUtf8Output) {
        $startInfo.EnvironmentVariables['AWS_CLI_OUTPUT_ENCODING'] = 'UTF-8'
    }

    $label   = (@($Arguments | Select-Object -First 2) -join ' ')
    $process = [System.Diagnostics.Process]::Start($startInfo)
    try {
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        $deadline   = [DateTime]::UtcNow.AddSeconds($AwsTimeoutSec)
        while (-not $process.WaitForExit(50)) {
            Test-Cancel
            if ([DateTime]::UtcNow -gt $deadline) {
                throw "aws $label が $AwsTimeoutSec 秒たっても終わりません（MFA コードの入力待ちになっている可能性があります）。"
            }
        }
        $stdoutText = $stdoutTask.Result
        $stderrText = ([string]$stderrTask.Result).Trim()
        if ($process.ExitCode -ne 0) {
            throw "aws $label が失敗しました（終了コード $($process.ExitCode)）。`r`n$stderrText"
        }
        return $stdoutText
    }
    finally {
        try { if (-not $process.HasExited) { $process.Kill() } } catch { }
        $process.Dispose()
    }
}

# aws コマンドの場所・バージョン・文字コードを確認し、共通の引数を組み立てる
function Initialize-Aws {
    param([string]$ProfileName, [string]$RegionName)

    $command = Get-Command aws -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $command) {
        throw 'aws コマンドが見つかりません。AWS CLI v2 をインストールしてください。'
    }
    $script:AwsPath       = $command.Path
    $script:AwsUtf8Output = $false
    $script:AwsEncoding   = $Utf8NoBom
    $script:CommonArgs    = @('--output', 'json', '--no-cli-pager')
    if ($ProfileName) { $script:CommonArgs += @('--profile', $ProfileName) }
    if ($RegionName)  { $script:CommonArgs += @('--region', $RegionName) }

    $versionText = (Invoke-Aws -Arguments @('--version') -Bare).Trim()
    $cliVersion  = $null
    if ($versionText -match 'aws-cli/(\d+\.\d+\.\d+)') { $cliVersion = [version]$Matches[1] }
    if (-not $cliVersion -or $cliVersion.Major -lt 2) {
        throw "AWS CLI v2 が必要です（検出: $versionText）。"
    }

    if ($cliVersion -ge [version]'2.24.14') {
        # 出力の文字コード指定は 2.24.14 以降で使える
        $script:AwsUtf8Output = $true
    }
    elseif ([Environment]::OSVersion.Platform -eq [PlatformID]::Win32NT) {
        # 古い CLI は Windows の既定の文字コード（日本語環境なら Shift_JIS）で出力する
        $encoding = [System.Text.Encoding]::Default
        if ($encoding.CodePage -eq 65001) {
            try { $encoding = [System.Text.Encoding]::GetEncoding([Globalization.CultureInfo]::CurrentCulture.TextInfo.ANSICodePage) } catch { }
        }
        $script:AwsEncoding = $encoding
    }
}

# 検索語 1 つを Logs Insights の条件式に変換する
function ConvertTo-LikeTerm {
    param([string]$Text)

    $hasDouble = $Text.Contains('"')
    $hasSingle = $Text.Contains("'")
    if ($Text.Contains('\') -or ($hasDouble -and $hasSingle)) {
        # 引用符でくくれない場合は、正規表現として全文字をエスケープして渡す
        $pattern = [regex]::Escape($Text).Replace('\ ', ' ').Replace('/', '\/')
        return "@message like /$pattern/"
    }
    if ($hasDouble) {
        return "@message like '$Text'"
    }
    return "@message like `"$Text`""
}

# 指定期間（epoch 秒、両端を含む）でクエリを 1 回実行し、完了した結果を返す
function Invoke-InsightsQuery {
    param([long]$StartEpoch, [long]$EndEpoch)

    $request = [ordered]@{
        logGroupNames = @($script:LogGroups)
        startTime     = $StartEpoch
        endTime       = $EndEpoch
        queryString   = $script:QueryString
        limit         = $MaxRows
    }
    # 引用符や日本語を壊さないよう、パラメーターは JSON ファイル経由で渡す
    [IO.File]::WriteAllText($script:RequestFile, ($request | ConvertTo-Json -Depth 5), $Utf8NoBom)

    $started = Invoke-Aws -Arguments @('logs', 'start-query', '--cli-input-json', "file://$($script:RequestFile)") | ConvertFrom-Json
    $script:ActiveQueryId = $started.queryId
    $script:QueryCount++

    while ($true) {
        Wait-Interval -Milliseconds 1000
        $result = Invoke-Aws -Arguments @('logs', 'get-query-results', '--query-id', $script:ActiveQueryId) | ConvertFrom-Json
        if ($result.status -eq 'Complete') { break }
        if (@('Scheduled', 'Running') -notcontains $result.status) {
            throw "クエリが完了しませんでした（status: $($result.status)）。"
        }
    }
    $script:ActiveQueryId = $null
    if ($result.statistics) {
        $script:BytesScanned += [double]$result.statistics.bytesScanned
    }
    return $result
}

# クエリ結果 1 行を変換する。Line = CSV の 1 行 / Second = 発生時刻の epoch 秒 / Key = 重複判定用の一意キー
function ConvertTo-Row {
    param($Fields)

    $map = @{}
    foreach ($f in @($Fields)) { $map[$f.field] = $f.value }

    # @timestamp は UTC の "yyyy-MM-dd HH:mm:ss.fff"。JST に直す
    $timeText = [string]$map['@timestamp']
    $second   = [long]-1
    $utc      = [datetime]::MinValue
    if ([datetime]::TryParseExact($timeText, 'yyyy-MM-dd HH:mm:ss.fff', $Invariant, [Globalization.DateTimeStyles]::None, [ref]$utc)) {
        $second   = [long][math]::Floor(($utc - $EpochUtc).TotalSeconds)
        $timeText = $utc.AddHours(9).ToString('yyyy-MM-dd HH:mm:ss.fff', $Invariant)
    }

    # @ptr はログイベントごとに一意。万一無い場合は内容で代用する
    $key = [string]$map['@ptr']
    if (-not $key) { $key = '{0}|{1}|{2}' -f $timeText, $map['@logStream'], $map['@message'] }

    # CSV の 1 行を組み立てる（全項目を引用符でくくり、項目内の " は "" にする。改行はそのまま残す）
    $group   = ([string]$map['@log']) -replace '^\d{12}:', ''
    $stream  = [string]$map['@logStream']
    $message = [string]$map['@message']
    $line = '"' + $timeText.Replace('"', '""') + '","' + $group.Replace('"', '""') + '","' +
            $stream.Replace('"', '""') + '","' + $message.Replace('"', '""') + '"'

    @{
        Second = $second
        Key    = $key
        Line   = $line
    }
}

# 期間内の全件を時刻順に取得し、CSV（$script:Writer）へ順次書き出す。
# 上限（limit 10000）に達したら、取得済みの行は捨てずに「最後に取得した秒」から続きを取得する。
# 続きの範囲は、残り件数（recordsMatched - 取得件数）から 1 回あたり約 $TargetRows 件になるよう分割する。
function Get-LogRows {
    param([long]$StartEpoch, [long]$EndEpoch)

    $result = Invoke-InsightsQuery -StartEpoch $StartEpoch -EndEpoch $EndEpoch
    $count  = @($result.results).Count

    # 直前のクエリと重なる 1 秒分の取得済みキー（今回のクエリにだけ適用する）
    $skipSecond = $script:OverlapSecond
    $skipKeys   = $script:OverlapKeys
    $script:OverlapSecond = [long]-1
    $script:OverlapKeys   = New-Object 'System.Collections.Generic.HashSet[string]'

    # 今回の結果のうち「最後の 1 秒」に含まれる行のキーを集めながら出力する
    $lastSecond = [long]-1
    $lastKeys   = New-Object 'System.Collections.Generic.HashSet[string]'
    $index      = 0
    foreach ($fields in $result.results) {
        $index++
        if ($index % 1000 -eq 0) { Test-Cancel }
        $item = ConvertTo-Row -Fields $fields
        if ($item.Second -ne $lastSecond) {
            $lastSecond = $item.Second
            $lastKeys.Clear()
        }
        [void]$lastKeys.Add($item.Key)
        if ($item.Second -eq $skipSecond -and $skipKeys.Contains($item.Key)) { continue }
        $script:Writer.WriteLine($item.Line)
        $script:RowCount++
    }

    # 件数が多いときは、5 秒に 1 回ほど途中経過を出す
    if ($script:QueryCount -gt 1 -and ([DateTime]::UtcNow - $script:LastProgress).TotalSeconds -ge 5) {
        Write-Status ('  {0} 件まで取得しました...' -f $script:RowCount)
        $script:LastProgress = [DateTime]::UtcNow
    }

    if ($count -lt $MaxRows) { return }

    # ---- ここから上限超過時の処理 ----
    if ($lastSecond -lt 0) {
        throw '上限を超えた分の続きを取得できません（@timestamp の形式が想定と異なります）。'
    }

    if ($lastSecond -le $StartEpoch) {
        # 開始の 1 秒だけで上限に達した。この秒の残りは取得する手段がないので、警告して次の秒へ進む
        $script:Truncated = $true
        $next = $StartEpoch + 1
    }
    else {
        # 最後の秒は途中で切れている可能性があるので、その秒から取り直し、取得済みの行は重複として除外する
        $next = $lastSecond
        $script:OverlapSecond = $lastSecond
        $script:OverlapKeys   = $lastKeys
    }
    if ($next -gt $EndEpoch) { return }

    # 残り件数から分割数を決める（統計が取れない場合は 2 分割）。偏りで再び上限に達しても同じ処理が再帰的に働く
    $parts   = [long]2
    $matched = [double]0
    if ($result.statistics) {
        $matchedProperty = $result.statistics.PSObject.Properties['recordsMatched']
        if ($matchedProperty -and $matchedProperty.Value) { $matched = [double]$matchedProperty.Value }
    }
    if ($matched -ge $count) {
        $parts = [long][math]::Ceiling(($matched - $count) / $TargetRows)
    }
    $seconds = $EndEpoch - $next + 1
    if ($parts -lt 1) { $parts = [long]1 }
    if ($parts -gt $seconds) { $parts = $seconds }
    $size = [long][math]::Ceiling($seconds / $parts)
    $result = $null   # 続きを取得している間、不要になった結果を抱えたままにしない

    Write-Status ('  上限 {0} 件に達しました。続きを {1} 回に分けて取得します...' -f $MaxRows, $parts)
    $windowStart = $next
    while ($windowStart -le $EndEpoch) {
        $windowEnd = [math]::Min($windowStart + $size - 1, $EndEpoch)
        Get-LogRows -StartEpoch $windowStart -EndEpoch $windowEnd
        $windowStart = $windowEnd + 1
    }
}

# 検索の本体。入力チェック → 認証確認 → 検索しながら CSV へ出力、までを行う。
# 検索条件が空なら期間内の全件を取得する。
# 出力した CSV のパスは $script:LastOutPath に入る（0 件のときは $null）
function Invoke-LogSearch {
    param(
        [string[]]$SearchLogGroup,
        [string[]]$SearchKeyword,
        [int]$SearchDays = 1,
        [string]$SearchFrom,
        [string]$SearchTo,
        [bool]$UseOr,
        [string]$OutputFile,
        [string]$ProfileName,
        [string]$RegionName
    )

    $script:LastOutPath   = $null
    $script:Writer        = $null
    $script:RowCount      = 0
    $script:LastProgress  = [DateTime]::UtcNow
    $script:ActiveQueryId = $null
    $script:BytesScanned  = [double]0
    $script:QueryCount    = 0
    $script:Truncated     = $false
    $script:IgnoreCancel  = $false
    $script:OverlapSecond = [long]-1
    $script:OverlapKeys   = New-Object 'System.Collections.Generic.HashSet[string]'

    # ---- 入力チェックと期間の計算 ----
    $groups   = @($SearchLogGroup | Where-Object { $_ } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $keywords = @($SearchKeyword | Where-Object { $_ -and $_.Trim() })
    if ($groups.Count -eq 0) { throw 'ロググループを 1 つ以上指定してください。' }
    if ($groups.Count -gt 50) { throw 'ロググループは一度に 50 個までです。' }
    if ($SearchDays -lt 1 -or $SearchDays -gt $MaxRangeDays) { throw "日数は 1 ～ $MaxRangeDays で指定してください。" }

    $nowJst = [DateTimeOffset]::UtcNow.ToOffset($JstOffset)
    $endCap = $nowJst.AddSeconds(1)

    if ($SearchFrom) {
        $startJst = ConvertFrom-JstText -Text $SearchFrom
        if ($SearchTo) { $endJst = ConvertFrom-JstText -Text $SearchTo -EndOfDay } else { $endJst = $endCap }
    }
    else {
        if ($SearchTo) { throw '終了日時（-To）は開始日時（-From）と一緒に指定してください。' }
        $todayJst = New-Object DateTimeOffset($nowJst.Year, $nowJst.Month, $nowJst.Day, 0, 0, 0, $JstOffset)
        $startJst = $todayJst.AddDays(-($SearchDays - 1))
        $endJst   = $endCap
    }
    if ($endJst -gt $endCap) { $endJst = $endCap }
    if ($startJst -ge $endJst) { throw '開始日時が終了日時（または現在）より後になっています。' }
    if (($endJst - $startJst).TotalDays -gt $MaxRangeDays) {
        throw "期間が長すぎます（最大 $MaxRangeDays 日）。スキャン量の課金を抑えるため、期間を絞ってください。"
    }

    # start-query の期間は両端を含むので、終了側は 1 秒手前にする
    $startEpoch = [long][math]::Floor(($startJst.UtcDateTime - $EpochUtc).TotalSeconds)
    $endEpoch   = [long][math]::Floor(($endJst.UtcDateTime - $EpochUtc).TotalSeconds) - 1

    # 検索条件があれば filter を付ける。なければ期間内の全件が対象になる
    $filterPart = ''
    $condition  = '（指定なし＝期間内の全件を取得）'
    if ($keywords.Count -gt 0) {
        $joiner = ' and '
        if ($UseOr) { $joiner = ' or ' }
        $condition  = (@($keywords | ForEach-Object { ConvertTo-LikeTerm -Text $_ }) -join $joiner)
        $filterPart = " | filter $condition"
    }
    $script:LogGroups   = $groups
    $script:QueryString = "fields @timestamp, @log, @logStream, @message$filterPart | sort @timestamp asc | limit $MaxRows"

    if (-not $OutputFile) {
        $OutputFile = 'cwlogs_{0}.csv' -f $nowJst.ToString('yyyyMMdd_HHmmss', $Invariant)
    }
    $outPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFile)
    $outFolder = Split-Path -Path $outPath -Parent
    if ($outFolder -and -not (Test-Path -LiteralPath $outFolder)) {
        throw "出力先のフォルダーが存在しません: $outFolder"
    }
    # 取得中は .partial に書き、最後まで取得できたら正式なファイル名にする
    $partialPath = $outPath + '.partial'
    $completed   = $false

    # ---- 実行 ----
    $script:RequestFile = Join-Path ([IO.Path]::GetTempPath()) ('cwlogs_request_{0}.json' -f [guid]::NewGuid().ToString('N'))
    try {
        Initialize-Aws -ProfileName $ProfileName -RegionName $RegionName

        # 認証の確認（MFA セッション切れをここで検出する）
        try {
            $identity = Invoke-Aws -Arguments @('sts', 'get-caller-identity') | ConvertFrom-Json
        }
        catch {
            if ($script:CancelRequested) { throw }
            throw "AWS の認証を確認できませんでした。MFA でログインし直してから再実行してください。`r`n$($_.Exception.Message)"
        }

        $displayEnd = $endJst
        if ($endJst -eq $endCap) { $displayEnd = $nowJst }
        Write-Status ('アカウント    : {0}' -f $identity.Account)
        Write-Status ('ロググループ  : {0}' -f ($groups -join ', '))
        Write-Status ('期間 (JST)    : {0} ～ {1}' -f $startJst.ToString('yyyy-MM-dd HH:mm:ss', $Invariant), $displayEnd.ToString('yyyy-MM-dd HH:mm:ss', $Invariant))
        Write-Status ('検索条件      : {0}' -f $condition)
        Write-Status '検索中...'

        # Excel でそのまま開けるよう BOM 付き UTF-8 で、取得した分から順に書き出す（件数が多くてもメモリにためない）
        $script:Writer = New-Object System.IO.StreamWriter($partialPath, $false, $Utf8Bom)
        $script:Writer.NewLine = "`r`n"
        $script:Writer.WriteLine('"timestamp_jst","log_group","log_stream","message"')

        Get-LogRows -StartEpoch $startEpoch -EndEpoch $endEpoch

        $script:Writer.Dispose()
        $script:Writer = $null

        $scannedMb = [math]::Round($script:BytesScanned / 1MB, 1)
        Write-Status ('ヒット {0} 件 / クエリ {1} 回 / スキャン量 {2} MB' -f $script:RowCount, $script:QueryCount, $scannedMb)
        if ($script:Truncated) {
            Write-Status -IsWarning "同じ 1 秒間に $MaxRows 件以上ヒットした箇所があり、その秒の一部が欠けている可能性があります。検索条件で絞り込んでください。"
        }

        if ($script:RowCount -gt 1048575) {
            Write-Status -IsWarning 'Excel で開ける行数（1,048,576 行）を超えています。CSV には全件入っていますが、Excel では途中までしか表示されません。'
        }

        if ($script:RowCount -eq 0) {
            Remove-Item -LiteralPath $partialPath -Force
            Write-Status '該当するログはありませんでした（CSV は作成していません）。'
        }
        else {
            if (Test-Path -LiteralPath $outPath) { Remove-Item -LiteralPath $outPath -Force }
            [IO.File]::Move($partialPath, $outPath)
            $script:LastOutPath = $outPath
            Write-Status "出力しました: $outPath"
        }
        $completed = $true
    }
    finally {
        # 中断（Ctrl+C や「中止」）・失敗した場合の後始末
        $script:IgnoreCancel = $true
        if ($script:Writer) {
            try { $script:Writer.Dispose() } catch { }
            $script:Writer = $null
        }
        if (-not $completed -and (Test-Path -LiteralPath $partialPath)) {
            if ($script:RowCount -gt 0) {
                Write-Status ('途中までの {0} 件を次のファイルに残しました: {1}' -f $script:RowCount, $partialPath)
            }
            else {
                Remove-Item -LiteralPath $partialPath -Force
            }
        }
        if ($script:ActiveQueryId) {
            try { [void](Invoke-Aws -Arguments @('logs', 'stop-query', '--query-id', $script:ActiveQueryId)) } catch { }
            $script:ActiveQueryId = $null
        }
        if ($script:RequestFile -and (Test-Path -LiteralPath $script:RequestFile)) {
            Remove-Item -LiteralPath $script:RequestFile -Force
        }
        $script:IgnoreCancel = $false
    }
}

# =============================================================================
#  画面（Windows Forms）
# =============================================================================

# 画面の部品を 1 つ作って親に載せる
function New-UiControl {
    param([string]$Type, $Parent, [int]$X, [int]$Y, [int]$Width, [int]$Height, [string]$Text)

    $control = New-Object ("System.Windows.Forms.$Type")
    $control.Location = New-Object System.Drawing.Point($X, $Y)
    $control.Size     = New-Object System.Drawing.Size($Width, $Height)
    if ($Text) { $control.Text = $Text }
    $Parent.Controls.Add($control)
    return $control
}

# 期間の指定方法（日数／日時）に合わせて、入力欄の有効・無効を切り替える
function Update-RangeControl {
    $ui = $script:Ui
    $byDays = $ui.RangeDays.Checked
    $ui.Days.Enabled = $byDays
    $ui.From.Enabled = -not $byDays
    $ui.To.Enabled   = -not $byDays
}

# 「検索」ボタンの処理。画面の入力値で検索を実行する
function Start-GuiSearch {
    if ($script:Running) { return }
    $ui = $script:Ui

    # 検索条件が空欄のときは全件取得になるので、押し間違いに備えて確認する
    if (@($ui.Keyword.Lines | Where-Object { $_ -and $_.Trim() }).Count -eq 0) {
        $answer = [System.Windows.Forms.MessageBox]::Show(
            $ui.Form, "検索条件が空欄です。期間内のログを全件取得します。`r`n件数が多いと時間がかかります。続けますか？", '全件取得の確認',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }

    $script:Running         = $true
    $script:CancelRequested = $false
    $ui.Search.Enabled = $false
    $ui.Cancel.Enabled = $true
    $ui.Open.Enabled   = $false
    $ui.Log.Clear()
    try {
        $arguments = @{
            SearchLogGroup = @($ui.LogGroup.Text -split '[\r\n,]+')
            SearchKeyword  = @($ui.Keyword.Lines)
            UseOr          = [bool]$ui.ModeOr.Checked
            OutputFile     = $ui.OutFile.Text.Trim()
            ProfileName    = $ui.AwsProfile.Text.Trim()
            RegionName     = $ui.Region.Text.Trim()
        }
        if ($ui.RangeDays.Checked) {
            $arguments.SearchDays = [int]$ui.Days.Value
        }
        else {
            $arguments.SearchFrom = $ui.From.Value.ToString('yyyy-MM-dd HH:mm', $Invariant)
            if ($ui.To.Checked) {
                $arguments.SearchTo = $ui.To.Value.ToString('yyyy-MM-dd HH:mm', $Invariant)
            }
        }

        Invoke-LogSearch @arguments

        if ($script:LastOutPath) { $ui.Open.Enabled = $true }
    }
    catch {
        if ($script:CancelRequested) {
            Write-Status '中止しました。'
        }
        else {
            $message = $_.Exception.Message
            Write-Status "エラー: $message"
            [void][System.Windows.Forms.MessageBox]::Show(
                $ui.Form, $message, 'エラー',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
    finally {
        $script:Running         = $false
        $script:CancelRequested = $false
        $ui.Search.Enabled = $true
        $ui.Cancel.Enabled = $false
    }
}

# 入力画面を組み立てて返す（スクリプトの引数が指定されていれば初期値にする）
function New-SearchForm {
    $anchorWide  = [System.Windows.Forms.AnchorStyles]'Top, Left, Right'
    $anchorRight = [System.Windows.Forms.AnchorStyles]'Top, Right'
    $anchorFill  = [System.Windows.Forms.AnchorStyles]'Top, Bottom, Left, Right'

    $form = New-Object System.Windows.Forms.Form
    $form.Text          = 'CloudWatch Logs 検索'
    $form.Font          = New-Object System.Drawing.Font('Meiryo UI', 9)
    $form.ClientSize    = New-Object System.Drawing.Size(640, 660)
    $form.MinimumSize   = New-Object System.Drawing.Size(656, 560)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    $ui = @{ Form = $form }

    # ロググループ
    [void](New-UiControl Label $form 12 14 112 36 "ロググループ`r`n（1 行に 1 つ）")
    $ui.LogGroup = New-UiControl TextBox $form 130 12 498 62
    $ui.LogGroup.Multiline     = $true
    $ui.LogGroup.AcceptsReturn = $true
    $ui.LogGroup.ScrollBars    = [System.Windows.Forms.ScrollBars]::Vertical
    $ui.LogGroup.Anchor        = $anchorWide

    # 検索条件
    [void](New-UiControl Label $form 12 84 112 36 "検索条件`r`n（1 行に 1 つ）")
    [void](New-UiControl Label $form 12 122 116 18 '空欄なら全件取得')
    $ui.Keyword = New-UiControl TextBox $form 130 82 498 62
    $ui.Keyword.Multiline     = $true
    $ui.Keyword.AcceptsReturn = $true
    $ui.Keyword.ScrollBars    = [System.Windows.Forms.ScrollBars]::Vertical
    $ui.Keyword.Anchor        = $anchorWide

    $modePanel  = New-UiControl Panel $form 130 148 498 26
    $ui.ModeAnd = New-UiControl RadioButton $modePanel 0 2 170 22 'すべて含む（AND）'
    $ui.ModeOr  = New-UiControl RadioButton $modePanel 180 2 190 22 'いずれかを含む（OR）'
    $ui.ModeAnd.Checked = $true

    # 期間
    [void](New-UiControl Label $form 12 184 112 20 '期間（JST）')
    $rangePanel   = New-UiControl Panel $form 130 180 498 62
    $ui.RangeDays = New-UiControl RadioButton $rangePanel 0 2 60 24 '直近'
    $ui.Days      = New-UiControl NumericUpDown $rangePanel 62 2 56 24
    $ui.Days.Minimum = 1
    $ui.Days.Maximum = $MaxRangeDays
    [void](New-UiControl Label $rangePanel 124 5 300 20 '日分（今日を含む。1 = 今日の 0:00 から）')
    $ui.RangeDate = New-UiControl RadioButton $rangePanel 0 32 60 24 '日時'
    $ui.From      = New-UiControl DateTimePicker $rangePanel 62 32 140 24
    [void](New-UiControl Label $rangePanel 206 36 20 20 '～')
    $ui.To        = New-UiControl DateTimePicker $rangePanel 228 32 160 24
    [void](New-UiControl Label $rangePanel 392 36 106 20 '✓なし＝現在まで')
    foreach ($picker in @($ui.From, $ui.To)) {
        $picker.Format       = [System.Windows.Forms.DateTimePickerFormat]::Custom
        $picker.CustomFormat = 'yyyy-MM-dd HH:mm'
    }
    $ui.To.ShowCheckBox = $true

    # 出力先
    [void](New-UiControl Label $form 12 254 112 20 '出力先 CSV')
    $ui.OutFile = New-UiControl TextBox $form 130 250 408 24
    $ui.OutFile.Anchor = $anchorWide
    $ui.Browse  = New-UiControl Button $form 544 249 84 26 '参照...'
    $ui.Browse.Anchor = $anchorRight
    [void](New-UiControl Label $form 130 278 498 18 '空欄の場合は、現在のフォルダーに cwlogs_日時.csv を作成します。')

    # AWS の設定
    [void](New-UiControl Label $form 12 308 112 20 'プロファイル')
    $ui.AwsProfile = New-UiControl TextBox $form 130 304 160 24
    [void](New-UiControl Label $form 306 308 70 20 'リージョン')
    $ui.Region     = New-UiControl TextBox $form 378 304 160 24
    [void](New-UiControl Label $form 130 332 498 18 'どちらも空欄なら AWS CLI の既定の設定を使います。')

    # ボタン
    $ui.Search = New-UiControl Button $form 130 360 110 30 '検索'
    $ui.Cancel = New-UiControl Button $form 248 360 110 30 '中止'
    $ui.Open   = New-UiControl Button $form 366 360 130 30 'CSV を開く'
    $ui.Cancel.Enabled = $false
    $ui.Open.Enabled   = $false

    # 進行状況
    $ui.Log = New-UiControl TextBox $form 12 400 616 248
    $ui.Log.Multiline  = $true
    $ui.Log.ReadOnly   = $true
    $ui.Log.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $ui.Log.Anchor     = $anchorFill

    # ---- 初期値（スクリプトの引数を反映）----
    $nowJst = [DateTimeOffset]::UtcNow.ToOffset($JstOffset)
    if ($LogGroup) { $ui.LogGroup.Text = (@($LogGroup) -join "`r`n") }
    if ($Keyword)  { $ui.Keyword.Text  = (@($Keyword) -join "`r`n") }
    if ($Or)       { $ui.ModeOr.Checked = $true }
    $ui.Days.Value   = $Days
    $ui.From.Value   = $nowJst.DateTime.Date
    $ui.To.Value     = $nowJst.DateTime
    $ui.To.Checked   = $false
    $ui.RangeDays.Checked = $true
    if ($From) {
        $ui.From.Value = (ConvertFrom-JstText -Text $From).DateTime
        if ($To) {
            $ui.To.Value   = (ConvertFrom-JstText -Text $To -EndOfDay).DateTime
            $ui.To.Checked = $true
        }
        $ui.RangeDate.Checked = $true
    }
    if ($OutFile)    { $ui.OutFile.Text    = $OutFile }
    if ($AwsProfile) { $ui.AwsProfile.Text = $AwsProfile }
    if ($Region)     { $ui.Region.Text     = $Region }

    $ui.LogGroup.Select(0, 0)

    $script:Ui = $ui
    Update-RangeControl

    # ---- 操作時の処理 ----
    $ui.RangeDays.Add_CheckedChanged({ Update-RangeControl })
    $ui.Search.Add_Click({ Start-GuiSearch })
    $ui.Cancel.Add_Click({
        $script:CancelRequested = $true
        $script:Ui.Cancel.Enabled = $false
    })
    $ui.Browse.Add_Click({
        $dialog = New-Object System.Windows.Forms.SaveFileDialog
        $dialog.Filter     = 'CSV ファイル (*.csv)|*.csv|すべてのファイル (*.*)|*.*'
        $dialog.DefaultExt = 'csv'
        $dialog.FileName   = 'cwlogs_{0}.csv' -f [DateTimeOffset]::UtcNow.ToOffset($JstOffset).ToString('yyyyMMdd_HHmmss', $Invariant)
        if ($dialog.ShowDialog($script:Ui.Form) -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:Ui.OutFile.Text = $dialog.FileName
        }
        $dialog.Dispose()
    })
    $ui.Open.Add_Click({
        if ($script:LastOutPath -and (Test-Path -LiteralPath $script:LastOutPath)) {
            Invoke-Item -LiteralPath $script:LastOutPath
        }
    })
    $form.Add_FormClosing({
        # 検索中に閉じようとした場合は、まず検索を中止する（止まってからもう一度閉じる）
        if ($script:Running) {
            $script:CancelRequested = $true
            $_.Cancel = $true
        }
    })

    return $form
}

# 入力画面を表示する
function Show-SearchForm {
    try {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
    }
    catch {
        throw '画面を表示できません（Windows Forms が使えない環境です）。-LogGroup を指定して画面なしで実行してください。'
    }
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $form = New-SearchForm
    try {
        [void]$form.ShowDialog()
    }
    finally {
        $form.Dispose()
        $script:Ui = $null
    }
}

# =============================================================================
#  ここから実行
# =============================================================================

if ($Gui -or (-not $LogGroup -and -not $Keyword)) {
    # 画面あり
    Show-SearchForm
}
else {
    # 画面なし
    if (-not $LogGroup) {
        throw '-LogGroup を指定してください（引数なしで実行すると入力画面が開きます）。'
    }
    Invoke-LogSearch -SearchLogGroup $LogGroup -SearchKeyword $Keyword -SearchDays $Days `
        -SearchFrom $From -SearchTo $To -UseOr $Or.IsPresent `
        -OutputFile $OutFile -ProfileName $AwsProfile -RegionName $Region
}
