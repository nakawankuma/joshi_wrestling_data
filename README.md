# 女子プロレスラー一覧・試合カードツール

女子プロレスラーの一覧表示と、試合カード作成に使うHTMLツールです。
すべてブラウザで開いて使用できます。

## 選手一覧

[選手一覧を開く](./index.html)

`index.html` は、現役選手をデビュー年・団体別に確認できる一覧ページです。
団体での絞り込みや選手名検索ができます。

## 試合カード組み立てツール

[試合カード組み立てツールを開く](./match_card_planner.html)

`match_card_planner.html` は、団体別の選手ラベルをドラッグ＆ドロップして試合を組むツールです。

- 試合番号と試合名の自由入力
- 選手ラベルのクリック追加とドラッグ移動
- 3WAY、4WAY、5WAY以上に対応した対戦枠の追加・削除
- 通常対戦とイリミネーションの切り替え
- イリミネーションの入場順並べ替え
- 使用済み選手の薄色表示（繰り返し使用可能）
- 試合順の並べ替え
- 自動保存、テキストコピー、印刷
- 編集用JSONの表示とインポート
- 右上の「？ 使い方」から開ける操作ヘルプ

編集用JSONを保存しておくと、試合番号、対戦枠、選手配置、入場順、備考などを再現できます。

## VSカード画像生成ツール

[VSカード画像生成ツールを開く](./match_vs_generator_1.html)

`match_vs_generator_1.html` は、試合結果のテキストから試合ごとのVSカード画像を作るツールです。
個別PNGまたはZIPで保存できます。

## 選手データの更新

選手データの正本は `woman-excel.xlsx` です。
Excelの内容を `index.html` と `match_card_planner.html` の両方へ反映するときは、PowerShellで次を実行します。

```powershell
python -m pip install pandas openpyxl
python .\xlsx_to_html_converter_all_in_one.py
```

既定の設定は `converter_config.json` にあります。
1回の変換で、選手一覧と試合カード組み立てツールの団体・選手データが同時に更新されます。

## 主なファイル

- `index.html`：女子プロレスラー一覧
- `match_card_planner.html`：試合カード組み立て
- `match_vs_generator_1.html`：VSカード画像生成
- `woman-excel.xlsx`：選手データ
- `xlsx_to_html_converter_all_in_one.py`：Excelデータ反映スクリプト
- `converter_config.json`：変換設定

## ライセンス

MIT License
