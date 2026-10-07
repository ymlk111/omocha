Option Explicit

'==============================================================================
' CloudWatch Logs の CSV（@timestamp, @message の2列）からエビデンスを作成する
'
'   実行するマクロ : CreateCwLogEvidence
'   出力先         : 実行時にアクティブなブック（ログの日付ごとに新規シートを追加）
'   既存シート     : 削除・変更しない（同名シートがあれば _2, _3 ... の名前で追加）
'
'   後から変えたい値（列幅・行の高さ・文言・判定条件）は、すぐ下の「設定」に
'   まとめてあります。処理本体を触る必要はありません。
'==============================================================================

'------------------------------------------------------------------------------
' 設定 1/4 : レイアウト
'------------------------------------------------------------------------------
Private Const HEADER_ROW As Long = 15            ' ヘッダー行
Private Const DATA_START_ROW As Long = 16        ' 転記開始行

Private Const COL_WIDTH_A As Double = 10         ' A列（判定）の幅
Private Const COL_WIDTH_B As Double = 16         ' B列（理由）の幅
Private Const COL_WIDTH_C As Double = 24         ' C列（日付）の幅
Private Const COL_WIDTH_D As Double = 120        ' D列（内容）の幅

Private Const HEADER_ROW_HEIGHT As Double = 18   ' ヘッダー行の高さ（0 = 変更しない）
Private Const DATA_ROW_HEIGHT As Double = 18     ' データ行の高さ  （0 = 変更しない）
Private Const WRAP_TEXT As Boolean = False       ' データ行を折り返して表示するか

Private Const COLOR_UNCONFIRMED As Long = vbRed  ' 「未確認」の文字色

'------------------------------------------------------------------------------
' 設定 2/4 : シートに出力する文言
'------------------------------------------------------------------------------
Private Const TXT_HEADER_A As String = "判定"
Private Const TXT_HEADER_B As String = "理由"
Private Const TXT_HEADER_C As String = "日付"
Private Const TXT_HEADER_D As String = "内容"

Private Const TXT_UNCONFIRMED As String = "未確認"        ' error / warn を含む行のA列
Private Const TXT_EXCERPT As String = "前後5行抜粋"       ' スタックトレースを抜粋した行のB列
Private Const TXT_CONTEXT_ROW As String = ""              ' 抜粋で追加された前後行のB列（空 = 何も入れない）
Private Const TXT_TRUNCATED As String = "...(以降省略)"   ' セル上限を超えた内容の末尾

Private Const SHEET_NAME_PREFIX As String = ""            ' シート名の接頭辞（例 "LOG_"）
Private Const SHEET_DATE_SEP As String = "-"              ' シート名の日付区切り（/ は使用不可）
Private Const TXT_UNKNOWN_DATE As String = "日付不明"     ' 日付を読み取れない行のシート名

'------------------------------------------------------------------------------
' 設定 3/4 : 判定条件・CSV
'------------------------------------------------------------------------------
Private Const CONTEXT_LINES As Long = 5                   ' 前後に抜粋する行数
Private Const PTN_ERROR As String = "error|warn"          ' 未確認とする条件（大文字小文字は区別しない）

' スタックトレース行の条件 : メッセージが「at 」または「日付(時刻) at 」で始まる
Private Const PTN_STACKTRACE As String = _
    "^\s*(\d{4}[-/]\d{1,2}[-/]\d{1,2}([T ]\d{1,2}:\d{2}(:\d{2})?([.,]\d+)?)?(Z|[+-]\d{2}:?\d{2})?\s+)?at\s"

Private Const SEARCH_IGNORE_CASE As Boolean = True        ' 検索条件で大文字小文字を区別しない

Private Const CSV_CHARSET As String = "UTF-8"             ' CSVの文字コード
Private Const CSV_COL_TIMESTAMP As String = "@timestamp"  ' CSVの列名（日付）
Private Const CSV_COL_MESSAGE As String = "@message"      ' CSVの列名（内容）

' @timestamp に加算する時間。CSVがUTCで、JSTの日付でシートを分けたい場合は 9 にする。
' 0 のときは @timestamp を一切加工せずそのまま転記する。
Private Const TZ_OFFSET_HOURS As Double = 0

Private Const MAX_CELL_CHARS As Long = 32767              ' Excelの1セルの文字数上限
Private Const BULK_WRITE_MAX_CHARS As Long = 8000         ' これを超える内容は一括転記せず1セルずつ書き込む

'------------------------------------------------------------------------------
' 設定 4/4 : ダイアログの文言
'------------------------------------------------------------------------------
Private Const TXT_TITLE As String = "CloudWatchログ エビデンス作成"
Private Const TXT_DLG_FILE_TITLE As String = "CloudWatch Logs のCSVファイルを選択"
Private Const TXT_FILE_FILTER As String = "CSVファイル (*.csv),*.csv,すべてのファイル (*.*),*.*"
Private Const TXT_DLG_SEARCH_PROMPT As String = "検索条件を入力してください（正規表現可・空欄で全件）"
Private Const TXT_ERR_NO_BOOK As String = "出力先のブックが開かれていません。"
Private Const TXT_ERR_REGEX As String = "検索条件を正規表現として解釈できません。"
Private Const TXT_ERR_HEADER As String = "CSVの1行目に必要な列が見つかりません。"
Private Const TXT_MSG_NO_DATA As String = "CSVにデータ行がありません。"
Private Const TXT_MSG_NO_HIT As String = "検索条件に該当するログはありませんでした。"
Private Const TXT_MSG_DONE As String = "エビデンスを作成しました。"
Private Const TXT_MSG_ROWS As String = "行"
Private Const TXT_ERR_UNEXPECTED As String = "処理中にエラーが発生しました。"

'------------------------------------------------------------------------------
' 以下、処理本体
'------------------------------------------------------------------------------
Private mNextComma As Long
Private mNextLf As Long

Public Sub CreateCwLogEvidence()
    Dim wb As Workbook, ws As Worksheet
    Dim csvPath As Variant, keyword As String
    Dim ts() As String, msg() As String, total As Long
    Dim reSearch As Object, reErr As Object, reStack As Object
    Dim reDate As Object, reTs As Object
    Dim matchAll As Boolean, hit As Boolean, found As Boolean
    Dim inc() As Boolean, isStack() As Boolean, isHit() As Boolean
    Dim judge() As String, reason() As String, dkey() As String, owner() As Long
    Dim dict As Object, k As Variant
    Dim buf() As Variant
    Dim i As Long, j As Long, lo As Long, hi As Long, r As Long, cnt As Long
    Dim summary As String

    On Error GoTo ErrHandler

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox TXT_ERR_NO_BOOK, vbExclamation, TXT_TITLE
        Exit Sub
    End If

    '--- 入力 1 : CSVファイル
    csvPath = Application.GetOpenFilename(FileFilter:=TXT_FILE_FILTER, Title:=TXT_DLG_FILE_TITLE)
    If VarType(csvPath) = vbBoolean Then Exit Sub            ' キャンセル

    '--- 入力 2 : 検索条件（正規表現）
    keyword = InputBox(TXT_DLG_SEARCH_PROMPT, TXT_TITLE)
    If StrPtr(keyword) = 0 Then Exit Sub                     ' キャンセル
    matchAll = (Len(keyword) = 0)

    If Not matchAll Then
        Set reSearch = NewRegExp(keyword, SEARCH_IGNORE_CASE)
        If Not IsValidPattern(reSearch) Then
            MsgBox TXT_ERR_REGEX & vbCrLf & keyword, vbExclamation, TXT_TITLE
            Exit Sub
        End If
    End If
    Set reErr = NewRegExp(PTN_ERROR, True)
    Set reStack = NewRegExp(PTN_STACKTRACE, False)
    Set reDate = NewRegExp("(\d{4})[-/](\d{1,2})[-/](\d{1,2})", False)
    Set reTs = NewRegExp("^\s*(\d{4})[-/](\d{1,2})[-/](\d{1,2})[T ](\d{1,2}):(\d{2}):(\d{2})([.,]\d+)?", False)

    '--- CSV読み込み
    total = LoadCsv(CStr(csvPath), ts, msg)
    If total < 0 Then
        MsgBox TXT_ERR_HEADER & vbCrLf & CSV_COL_TIMESTAMP & " / " & CSV_COL_MESSAGE, vbExclamation, TXT_TITLE
        Exit Sub
    End If
    If total = 0 Then
        MsgBox TXT_MSG_NO_DATA, vbInformation, TXT_TITLE
        Exit Sub
    End If

    ReDim inc(1 To total)
    ReDim isStack(1 To total)
    ReDim isHit(1 To total)
    ReDim judge(1 To total)
    ReDim reason(1 To total)
    ReDim dkey(1 To total)
    ReDim owner(1 To total)

    '--- スタックトレース行の判定（全行）
    For i = 1 To total
        isStack(i) = reStack.Test(msg(i))
    Next i

    '--- 検索条件・error/warn・前後抜粋の判定
    For i = 1 To total
        If matchAll Then
            hit = True
        Else
            hit = reSearch.Test(msg(i))
        End If

        If hit Then
            inc(i) = True
            isHit(i) = True
            owner(i) = i
            If reErr.Test(msg(i)) Then
                judge(i) = TXT_UNCONFIRMED                   ' (1) error / warn を含む

                lo = i - CONTEXT_LINES
                If lo < 1 Then lo = 1
                hi = i + CONTEXT_LINES
                If hi > total Then hi = total

                found = False
                For j = lo To hi
                    If j <> i Then
                        If isStack(j) Then
                            found = True
                            Exit For
                        End If
                    End If
                Next j

                If found Then                                ' (2) 前後にスタックトレースあり
                    reason(i) = TXT_EXCERPT
                    For j = lo To hi
                        inc(j) = True
                        If owner(j) = 0 Then owner(j) = i    ' 抜粋行は元のログと同じシートに出す
                    Next j
                End If
            End If
        End If
    Next i

    '--- 転記対象を日付ごとに集計（CSVに現れた順）
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 1 To total
        If inc(i) Then ts(i) = ShiftTimestamp(ts(i), reTs)
    Next i
    For i = 1 To total
        If inc(i) Then
            dkey(i) = DateKeyOf(ts(owner(i)), reDate)
            If dict.Exists(dkey(i)) Then
                dict(dkey(i)) = dict(dkey(i)) + 1
            Else
                dict.Add dkey(i), 1
            End If
        End If
    Next i

    If dict.Count = 0 Then
        MsgBox TXT_MSG_NO_HIT, vbInformation, TXT_TITLE
        Exit Sub
    End If

    '--- (3) 日付ごとに新規シートを作成して転記
    Application.ScreenUpdating = False
    For Each k In dict.Keys
        cnt = dict(k)
        ReDim buf(1 To cnt, 1 To 4)
        r = 0
        For i = 1 To total
            If inc(i) Then
                If dkey(i) = k Then
                    r = r + 1
                    buf(r, 1) = judge(i)
                    If isHit(i) Then
                        buf(r, 2) = reason(i)
                    Else
                        buf(r, 2) = TXT_CONTEXT_ROW          ' 抜粋で追加された前後行
                    End If
                    buf(r, 3) = ts(i)
                    buf(r, 4) = CellText(msg(i))
                End If
            End If
        Next i

        Set ws = AddEvidenceSheet(wb, CStr(k))
        WriteSheet ws, buf, cnt
        summary = summary & vbCrLf & ws.Name & " : " & cnt & TXT_MSG_ROWS
    Next k
    Application.ScreenUpdating = True

    MsgBox TXT_MSG_DONE & vbCrLf & summary, vbInformation, TXT_TITLE
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox TXT_ERR_UNEXPECTED & vbCrLf & Err.Number & " : " & Err.Description, vbCritical, TXT_TITLE
End Sub

'------------------------------------------------------------------------------
' シートへの書き込み（ヘッダー・データ・列幅・行の高さ・文字色）
'------------------------------------------------------------------------------
Private Sub WriteSheet(ByVal ws As Worksheet, ByRef buf() As Variant, ByVal cnt As Long)
    Dim rng As Range, r As Long
    Dim longTxt() As String, hasLong As Boolean

    ' 長い内容は一括転記で失敗することがあるため、別に1セルずつ書き込む
    ReDim longTxt(1 To cnt)
    For r = 1 To cnt
        If Len(buf(r, 4)) > BULK_WRITE_MAX_CHARS Then
            longTxt(r) = buf(r, 4)
            buf(r, 4) = vbNullString
            hasLong = True
        End If
    Next r

    With ws
        .Columns(1).ColumnWidth = COL_WIDTH_A
        .Columns(2).ColumnWidth = COL_WIDTH_B
        .Columns(3).ColumnWidth = COL_WIDTH_C
        .Columns(4).ColumnWidth = COL_WIDTH_D

        .Cells(HEADER_ROW, 1).Value = TXT_HEADER_A
        .Cells(HEADER_ROW, 2).Value = TXT_HEADER_B
        .Cells(HEADER_ROW, 3).Value = TXT_HEADER_C
        .Cells(HEADER_ROW, 4).Value = TXT_HEADER_D

        Set rng = .Range(.Cells(DATA_START_ROW, 1), .Cells(DATA_START_ROW + cnt - 1, 4))
        rng.NumberFormat = "@"                   ' 文字列として転記（日付の自動変換・数式化を防ぐ）
        rng.Value = buf
        If hasLong Then
            For r = 1 To cnt
                If Len(longTxt(r)) > 0 Then .Cells(DATA_START_ROW + r - 1, 4).Value = longTxt(r)
            Next r
        End If
        rng.WrapText = WRAP_TEXT

        If HEADER_ROW_HEIGHT > 0 Then .Rows(HEADER_ROW).RowHeight = HEADER_ROW_HEIGHT
        If DATA_ROW_HEIGHT > 0 Then rng.EntireRow.RowHeight = DATA_ROW_HEIGHT

        For r = 1 To cnt
            If Len(buf(r, 1)) > 0 Then
                .Cells(DATA_START_ROW + r - 1, 1).Font.Color = COLOR_UNCONFIRMED
            End If
        Next r
    End With
End Sub

'------------------------------------------------------------------------------
' 新規シートを末尾に追加する。同名シートがある場合は _2, _3 ... を付ける
' （既存シートには触れない）
'------------------------------------------------------------------------------
Private Function AddEvidenceSheet(ByVal wb As Workbook, ByVal baseName As String) As Worksheet
    Dim nm As String, n As Long
    Dim ws As Worksheet

    nm = baseName
    n = 1
    Do While SheetExists(wb, nm)
        n = n + 1
        nm = baseName & "_" & n
    Loop

    Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = nm
    Set AddEvidenceSheet = ws
End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim sh As Object
    On Error Resume Next
    Set sh = wb.Sheets(nm)
    On Error GoTo 0
    SheetExists = Not sh Is Nothing
End Function

'------------------------------------------------------------------------------
' @timestamp からシート名（日付）を作る
'------------------------------------------------------------------------------
Private Function DateKeyOf(ByVal tsText As String, ByVal reDate As Object) As String
    Dim m As Object

    Set m = reDate.Execute(tsText)
    If m.Count = 0 Then
        DateKeyOf = TXT_UNKNOWN_DATE
    Else
        With m.Item(0)
            DateKeyOf = SHEET_NAME_PREFIX & .SubMatches(0) & SHEET_DATE_SEP & _
                        Format$(CLng(.SubMatches(1)), "00") & SHEET_DATE_SEP & _
                        Format$(CLng(.SubMatches(2)), "00")
        End With
    End If
End Function

'------------------------------------------------------------------------------
' TZ_OFFSET_HOURS が 0 以外のとき、@timestamp に時間を加算する
'------------------------------------------------------------------------------
Private Function ShiftTimestamp(ByVal tsText As String, ByVal reTs As Object) As String
    Dim m As Object, d As Date, frac As String

    ShiftTimestamp = tsText
    If TZ_OFFSET_HOURS = 0 Then Exit Function

    Set m = reTs.Execute(tsText)
    If m.Count = 0 Then Exit Function

    With m.Item(0)
        d = DateSerial(CLng(.SubMatches(0)), CLng(.SubMatches(1)), CLng(.SubMatches(2))) + _
            TimeSerial(CLng(.SubMatches(3)), CLng(.SubMatches(4)), CLng(.SubMatches(5)))
        frac = .SubMatches(6) & ""               ' 小数秒（".123" など。無ければ空）
    End With
    d = DateAdd("n", CLng(TZ_OFFSET_HOURS * 60), d)
    ShiftTimestamp = Format$(d, "yyyy-mm-dd hh:nn:ss") & frac
End Function

'------------------------------------------------------------------------------
' セルの文字数上限を超える内容を切り詰める
'------------------------------------------------------------------------------
Private Function CellText(ByRef s As String) As String
    If Len(s) > MAX_CELL_CHARS Then
        CellText = Left$(s, MAX_CELL_CHARS - Len(TXT_TRUNCATED)) & TXT_TRUNCATED
    Else
        CellText = s
    End If
End Function

'------------------------------------------------------------------------------
' 正規表現
'------------------------------------------------------------------------------
Private Function NewRegExp(ByVal ptn As String, ByVal bIgnoreCase As Boolean) As Object
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = ptn
    re.IgnoreCase = bIgnoreCase
    re.Global = False
    re.MultiLine = False
    Set NewRegExp = re
End Function

Private Function IsValidPattern(ByVal re As Object) As Boolean
    Dim dummy As Boolean
    On Error Resume Next
    Err.Clear
    dummy = re.Test("")
    IsValidPattern = (Err.Number = 0)
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' CSV読み込み
'   戻り値 : データ行数（ヘッダーに必要な列が無い場合は -1）
'   "" で囲まれた項目、項目内のカンマ・改行・"" に対応
'------------------------------------------------------------------------------
Private Function LoadCsv(ByVal path As String, ByRef ts() As String, ByRef msg() As String) As Long
    Dim s As String, n As Long, pos As Long
    Dim flds() As String, cnt As Long, i As Long
    Dim tsCol As Long, msgCol As Long
    Dim total As Long, cap As Long

    LoadCsv = -1

    With CreateObject("ADODB.Stream")
        .Type = 2                                ' テキスト
        .Charset = CSV_CHARSET
        .Open
        .LoadFromFile path
        s = .ReadText(-1)
        .Close
    End With

    If Left$(s, 1) = ChrW$(&HFEFF) Then s = Mid$(s, 2)       ' BOM除去
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    n = Len(s)
    If n = 0 Then Exit Function

    mNextComma = 0
    mNextLf = 0
    pos = 1
    ReDim flds(0 To 7)

    '--- ヘッダー行から列位置を決める
    cnt = ReadCsvRecord(s, pos, n, flds)
    If Left$(flds(0), 5) = "#TYPE" And pos <= n Then         ' PowerShell の Export-Csv が付ける型情報行
        cnt = ReadCsvRecord(s, pos, n, flds)
    End If
    tsCol = -1
    msgCol = -1
    For i = 0 To cnt - 1
        Select Case LCase$(Trim$(flds(i)))
            Case LCase$(CSV_COL_TIMESTAMP): tsCol = i
            Case LCase$(CSV_COL_MESSAGE): msgCol = i
        End Select
    Next i
    If tsCol < 0 Or msgCol < 0 Then Exit Function

    '--- データ行
    cap = 1024
    ReDim ts(1 To cap)
    ReDim msg(1 To cap)

    Do While pos <= n
        cnt = ReadCsvRecord(s, pos, n, flds)
        If cnt = 1 And Len(flds(0)) = 0 Then
            ' 空行は読み飛ばす
        Else
            total = total + 1
            If total > cap Then
                cap = cap * 2
                ReDim Preserve ts(1 To cap)
                ReDim Preserve msg(1 To cap)
            End If
            If tsCol < cnt Then ts(total) = flds(tsCol)
            If msgCol < cnt Then msg(total) = flds(msgCol)
        End If
    Loop

    LoadCsv = total
End Function

' 1レコード（複数行にまたがる場合あり）を読み、項目数を返す
Private Function ReadCsvRecord(ByRef s As String, ByRef pos As Long, ByVal n As Long, _
                               ByRef flds() As String) As Long
    Dim cnt As Long, fld As String, q As Long, e As Long

    Do
        fld = vbNullString

        If Mid$(s, pos, 1) = """" Then           ' "" で囲まれた項目
            pos = pos + 1
            Do
                q = InStr(pos, s, """", vbBinaryCompare)
                If q = 0 Then                    ' 閉じ引用符なし（ファイル末尾まで）
                    fld = fld & Mid$(s, pos)
                    pos = n + 1
                    Exit Do
                End If
                fld = fld & Mid$(s, pos, q - pos)
                If Mid$(s, q + 1, 1) = """" Then ' "" は " 1文字
                    fld = fld & """"
                    pos = q + 2
                Else
                    pos = q + 1
                    Exit Do
                End If
            Loop
        End If

        e = NextDelim(s, pos, n)                 ' 次のカンマまたは改行
        If e > pos Then fld = fld & Mid$(s, pos, e - pos)
        pos = e

        If cnt > UBound(flds) Then ReDim Preserve flds(0 To cnt * 2 + 1)
        flds(cnt) = fld
        cnt = cnt + 1

        If pos > n Then Exit Do                  ' ファイル末尾
        If Mid$(s, pos, 1) = vbLf Then           ' レコード終端
            pos = pos + 1
            Exit Do
        End If
        pos = pos + 1                            ' カンマを読み飛ばして次の項目へ
    Loop

    ReadCsvRecord = cnt
End Function

' pos 以降で最初に現れるカンマまたは改行の位置（無ければ n + 1）
Private Function NextDelim(ByRef s As String, ByVal pos As Long, ByVal n As Long) As Long
    If mNextComma < pos Then
        mNextComma = InStr(pos, s, ",", vbBinaryCompare)
        If mNextComma = 0 Then mNextComma = n + 1
    End If
    If mNextLf < pos Then
        mNextLf = InStr(pos, s, vbLf, vbBinaryCompare)
        If mNextLf = 0 Then mNextLf = n + 1
    End If
    If mNextComma < mNextLf Then
        NextDelim = mNextComma
    Else
        NextDelim = mNextLf
    End If
End Function
