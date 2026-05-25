# Excel 行分割スクリプト (Split-ExcelRowsToSheets)

Excelの特定シートの特定行以降を1行ずつ取り出し、**元シートのコピーを行数分作成**して、それぞれの新シートの指定行に1行だけ貼り付けるPowerShellスクリプトです。

書式・列幅・数式・条件付き書式・印刷設定など、元シートのレイアウトはそのまま保持されます。帳票テンプレートの一括展開などに使えます。

---

## ファイル構成

| ファイル | 用途 |
|---|---|
| `Split-ExcelRowsToSheets.ps1` | メインスクリプト |
| `Run-Split.bat` | 対話形式の起動バッチ(ダブルクリック実行用) |
| `README.md` | このファイル |

---

## 動作要件

- Windows
- Microsoft Excel がインストール済み (COM経由で操作するため)
- PowerShell 5.1 以上 (Windows 10/11標準)

---

## 動作イメージ

元ファイル(変更前):

```
[明細シート]
 1行目: ヘッダー
 2行目: タイトル等
 3行目: 共通項目
 4行目: 表ヘッダ
 5行目: データ1  ← StartRow
 6行目: データ2
 7行目: データ3
```

実行後(`-StartRow 5 -PasteRow 5` の場合):

```
[明細シート]         ← 元のまま
[明細No1シート]      ← 5行目に データ1 のみ
[明細No2シート]      ← 5行目に データ2 のみ
[明細No3シート]      ← 5行目に データ3 のみ
```

各新シートは「元シートを丸ごとコピー → データ行をクリア → 該当1行だけ貼り付け」という流れで作られるため、書式や周辺セルの内容はそのまま残ります。

---

## 使い方

### A. PowerShellから直接実行(推奨)

```powershell
# 最小構成: 5行目以降を1行ずつ分割し、新シートの1行目に貼り付け
.\Split-ExcelRowsToSheets.ps1 `
    -FilePath "C:\work\data.xlsx" `
    -SheetName "明細" `
    -StartRow 5
```

```powershell
# 貼り付け先行を5行目にし、書式ごと貼り付け、別ファイルに保存
.\Split-ExcelRowsToSheets.ps1 `
    -FilePath "C:\work\data.xlsx" `
    -SheetName "明細" `
    -StartRow 5 `
    -PasteRow 5 `
    -PasteMode AllContents `
    -OutputPath "C:\work\data_split.xlsx"
```

```powershell
# 動作確認用にExcelを開いたまま終了
.\Split-ExcelRowsToSheets.ps1 -FilePath ... -SheetName ... -StartRow 5 -KeepOpen
```

実行ポリシーで止まる場合は次のいずれか:

```powershell
# 一時的にこのプロセスのみ許可
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# またはコマンドラインから一発実行
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Split-ExcelRowsToSheets.ps1" -FilePath "..." -SheetName "..." -StartRow 5
```

### B. バッチからの対話実行

`Run-Split.bat` をダブルクリックすると、パラメータをコンソールで順に聞かれます。
PowerShellに不慣れな利用者へ配布する場合はこちらが手軽です。

---

## パラメータ一覧

| パラメータ | 必須 | 既定値 | 説明 |
|---|---|---|---|
| `-FilePath` | ○ | - | 対象Excelファイルのフルパス |
| `-SheetName` | ○ | - | 元になるシートの名前 |
| `-StartRow` | ○ | - | データ開始行(この行以降を分割) |
| `-PasteRow` | | 1 | 新シートで貼り付け先となる行番号 |
| `-PasteMode` | | `ValuesOnly` | `ValuesOnly` / `AllContents` / `Formulas` |
| `-OutputPath` | | (上書き) | 出力先ファイルパス |
| `-KeepOpen` | | OFF | 処理後にExcelを表示したままにする(検証用) |

### PasteModeの違い

| 値 | 動作 | 用途 |
|---|---|---|
| `ValuesOnly` | 値だけ貼り付け。書式は新シート側(=テンプレート由来)が使われる | 帳票テンプレートにデータを差し込みたい場合 |
| `AllContents` | 値・書式・罫線などすべて貼り付け | データ行ごとに書式が違う場合 |
| `Formulas` | 数式を保持して貼り付け | 元シートに計算式があり、その式を残したい場合 |

---

## 動作の詳細

1. `UsedRange` から元シートの最終行・最終列を自動判定
2. データ行数(= `LastRow - StartRow + 1`)分だけループ:
    1. 元シートを `Worksheet.Copy` で複製(末尾に配置)
    2. シート名を `{元シート名}No{連番}` に変更
    3. 新シートの `StartRow`以降のデータ部分を `ClearContents` で値だけクリア(書式は保持)
    4. 元シートの該当1行を `PasteSpecial` で `PasteRow` に貼り付け
3. すべて完了後に保存

### 内部最適化

- `ScreenUpdating = false` で画面再描画を抑制
- `Calculation = xlCalculationManual` で自動再計算を停止し、保存直前に元に戻す
- これにより数百シート規模でも比較的高速に動作します

---

## よくある状況への対応

### Q. シート名に31文字制限があるが大丈夫か
スクリプト側で31文字を超える場合は元シート名を自動で切り詰めます (`{切詰めた元名}No{連番}`)。

### Q. 再実行したらどうなる
`{元シート名}No{連番}` という名前で既に存在するシートがあれば削除してから作り直します。元シート自体には触れません。

### Q. データが空の行も処理されるのか
`UsedRange` の最終行までを処理対象とします。途中に空行があってもその行で空のデータが貼り付けられるだけです。完全な空行を除外したい場合はスクリプトのループ内に判定を追加してください。

### Q. 数式の参照がずれないか
`AllContents` または `Formulas` で貼り付けた場合、相対参照は貼り付け先位置に応じて自動調整されます。絶対参照(`$A$1`形式)は変わりません。

### Q. .xlsm (マクロ付き) でも動くか
動きます。`OutputPath` の拡張子が `.xlsm` であれば、マクロを保持したまま保存します。

---

## トラブルシューティング

### Excelのプロセスが残る
スクリプトは `finally` 句でExcelを終了しますが、エラー時にプロセスが残った場合は次で強制終了できます:

```powershell
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
```

### 「シートが見つかりません」エラー
シート名の前後に半角/全角スペースが入っていないか確認してください。エラーメッセージに実在するシート名一覧が表示されるので、コピペでそのまま使えます。

### 文字化け
バッチ実行(`Run-Split.bat`)で文字化けする場合、コマンドプロンプトのフォントを「MS ゴシック」または「Consolas」以外の日本語対応フォントに変更してください。バッチ冒頭で `chcp 65001` (UTF-8) に設定済みです。

### 動作確認したい
`-KeepOpen` スイッチを付けて実行すると、処理後にExcelが画面に表示されたままになります。中身を目視確認できます(その場合、保存はされていてもExcelは閉じられません)。

---

## 注意

- 元ファイルを直接書き換える場合は事前にバックアップを取ってください
- 大量シート(数百以上)生成時はファイルサイズが大きくなります
- Excel COM を使うため、サーバー環境やExcel未インストール環境では動きません
