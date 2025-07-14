# XLSX to HTML Converter

XLSXファイルのデータをHTMLファイルのJavaScript配列に変換する再利用可能なスクリプト

## 概要

このスクリプトは、ExcelファイルのデータをHTMLファイル内のJavaScript配列として埋め込むためのツールです。女子プロレス選手データベースのように、構造化されたデータをWebページで表示する際に活用できます。

## 必要な環境

- Python 3.6以上
- pandas ライブラリ

## インストール

```bash
pip install pandas openpyxl
```

## 基本的な使用方法

### 1. 簡単な実行

```bash
python3 xlsx_to_html_converter.py
```

デフォルトで以下のファイルを処理します：
- `woman-excel.xlsx` → `wrestling_data_chart.html`

### 2. ファイルを指定して実行

```bash
python3 xlsx_to_html_converter.py --xlsx data.xlsx --html output.html
```

### 3. 設定ファイルを使用

```bash
# 設定ファイルのサンプルを作成
python3 xlsx_to_html_converter.py --create-config

# 設定ファイルを使用して実行
python3 xlsx_to_html_converter.py --config converter_config.json
```

## 設定ファイル

`converter_config.json` で詳細な設定が可能：

```json
{
    "xlsx_file": "woman-excel.xlsx",
    "html_file": "wrestling_data_chart.html", 
    "sheet_name": 0,
    "data_start_row": 0,
    "year_column": 0,
    "data_columns_start": 1,
    "js_variable_name": "wrestlerData",
    "encoding": "utf-8"
}
```

### 設定項目の説明

- `xlsx_file`: 入力XLSXファイルのパス
- `html_file`: 出力HTMLファイルのパス
- `sheet_name`: 処理するシート（番号または名前）
- `data_start_row`: データ開始行（0ベース）
- `year_column`: 年度列のインデックス
- `data_columns_start`: データ列の開始インデックス
- `js_variable_name`: JavaScript変数名
- `encoding`: ファイルエンコーディング

## データフォーマット

### 入力XLSXフォーマット

```
| 年度 | 団体1 | 団体2 | 団体3 | ... |
|------|-------|-------|-------|-----|
| 2025 | 選手A | 選手B | 選手C | ... |
| 2024 | 選手D | 選手E |       | ... |
| ...  | ...   | ...   | ...   | ... |
```

### 出力JavaScript配列

```javascript
const wrestlerData = [
    [2025, "選手A", "選手B", "選手C", ...],
    [2024, "選手D", "選手E", "", ...],
    ...
];
```

## 使用例

### 女子プロレス選手データベース

```bash
# 現在のプロジェクト構成で実行
python3 xlsx_to_html_converter.py

# 結果：woman-excel.xlsx のデータが wrestling_data_chart.html に反映される
```

### カスタムデータ

```bash
# 売上データをチャートHTMLに反映
python3 xlsx_to_html_converter.py --xlsx sales_data.xlsx --html sales_chart.html
```

## エラー対処

### よくあるエラー

1. **ファイルが見つからない**
   ```
   XLSXファイル読み込みエラー: [Errno 2] No such file or directory
   ```
   → ファイルパスを確認してください

2. **JavaScript配列が見つからない**
   ```
   警告: wrestlerData配列が見つかりませんでした
   ```
   → HTMLファイル内の変数名を確認してください

3. **pandas がインストールされていない**
   ```
   ModuleNotFoundError: No module named 'pandas'
   ```
   → `pip install pandas openpyxl` を実行してください

## 拡張方法

### 新しいデータ形式への対応

1. `convert_to_js_array()` メソッドを修正
2. 必要に応じて設定項目を追加
3. データ処理ロジックをカスタマイズ

### HTMLテンプレート生成

スクリプトを拡張してHTMLテンプレート全体を生成することも可能です。

## ライセンス

このプロジェクトはMITライセンスのもとで公開されています。

## 貢献

バグ報告や機能要望、プルリクエストを歓迎します。

---

**注意**: このスクリプトはWebアプリケーションのデータを更新する際に使用してください。本番環境での使用前には必ずバックアップを取ることをお勧めします。