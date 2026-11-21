# omocha

ちょっとした作業を楽にするためのスクリプトと、遊び用の HTML の置き場です。
ほとんどが Windows 向けです。

```
omocha/
├─ excel/
│  ├─ split-rows-to-sheets/   行を1行ずつ別シートに展開
│  └─ log-to-excel/           ログの追記を Excel に自動転記
├─ cloudwatch-logs/           CloudWatch Logs の検索とエビデンス作成
├─ tools/
│  ├─ screenshot/             スクショ君
│  ├─ remove-whitespace/      空白・改行の除去
│  └─ sakura-md-preview/      サクラエディタ用 Markdown プレビュー
├─ games/                     飲み会・遊び用の HTML
└─ notes/                     プロンプトや調べ物のメモ
```

## Excel

| 場所 | できること | 起動方法 |
|---|---|---|
| [excel/split-rows-to-sheets](excel/split-rows-to-sheets/) | 指定した行以降を1行ずつ取り出し、元シートのコピーに貼り付けて行数分のシートを作る | `Run-Split.bat` をダブルクリック。詳しくは[フォルダ内の README](excel/split-rows-to-sheets/README.md) |
| [excel/log-to-excel](excel/log-to-excel/) | ログファイルへの追記を監視し、開いている Excel のアクティブシートへ1行1セルで転記する常駐ツール。`Ctrl+Alt+L` で ON/OFF | `powershell.exe -ExecutionPolicy Bypass -File .\LogToExcel.ps1`（Windows PowerShell 5.1 専用。PowerShell 7 では動きません） |

## CloudWatch Logs

| 場所 | できること | 起動方法 |
|---|---|---|
| [cloudwatch-logs/Search-CWLogs.ps1](cloudwatch-logs/Search-CWLogs.ps1) | ロググループをキーワードと期間（JST）で検索し、結果を CSV に出力する。AWS CLI の認証が済んでいる前提 | 引数なしで実行すると入力画面が開く。`-LogGroup` を付けると画面なしで実行 |
| [cloudwatch-logs/CwLogEvidence.bas](cloudwatch-logs/CwLogEvidence.bas) | `@timestamp`・`@message` 列を持つ CSV から、ログの日付ごとにエビデンス用シートを作る Excel マクロ | VBE でモジュールとして取り込み、`CreateCwLogEvidence` を実行 |

## ツール

| 場所 | できること | 起動方法 |
|---|---|---|
| [tools/screenshot](tools/screenshot/) | スクショ君。ホットキー（初期設定 `Ctrl+Shift+1`）でアクティブなウィンドウを撮影し、フォルダとクリップボードに保存 | `run.bat` をダブルクリック。[使い方.txt](tools/screenshot/使い方.txt) |
| [tools/remove-whitespace](tools/remove-whitespace/) | ドロップしたファイルから半角・全角スペース、タブ、改行を取り除いた別ファイルを作る（元ファイルは変更しない） | ファイルを `RemoveWhitespace.bat` にドラッグ＆ドロップ |
| [tools/sakura-md-preview](tools/sakura-md-preview/) | サクラエディタで編集中の Markdown をブラウザでプレビュー・編集するマクロ | マクロフォルダに置いて登録。手順は `mdview.js` の冒頭 |

## ゲーム

ブラウザで開くだけで遊べます。

| 場所 | 内容 |
|---|---|
| [games/ちんちろ.html](games/ちんちろ.html) | チンチロ酒場。丼ひとつで遊べる飲みゲー用チンチロリン |
| [games/どすこい.html](games/どすこい.html) | どすこい親。親の代わりに文字を出して読み上げる |

## メモ

| 場所 | 内容 |
|---|---|
| [notes/API設計レビュー_プロンプト.md](notes/API設計レビュー_プロンプト.md) | API 設計書を生成 AI にレビューさせるためのプロンプト |
| [notes/レビュー根拠.txt](notes/レビュー根拠.txt) | 設計書レビューの指摘（追跡可能性・API 設計）の根拠にできる公的文書・標準のまとめ |
