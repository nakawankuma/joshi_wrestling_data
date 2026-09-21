#!/usr/bin/env python3
"""
XLSX to HTML Converter - All-in-One Script
XLSXファイルのデータをHTMLファイルのJavaScript配列に変換する完全統合スクリプト
改行文字の問題も自動修正し、完全な再現性を保証します。

使用方法:
    python3 xlsx_to_html_converter_all_in_one.py
    python3 xlsx_to_html_converter_all_in_one.py --xlsx data.xlsx --html output.html
    python3 xlsx_to_html_converter_all_in_one.py --config config.json
"""

import pandas as pd
import json
import re
import argparse
import os
import sys
import html
from collections import Counter
from numbers import Real


WRESTLER_DATA_START = "// GENERATED WRESTLER DATA START"
WRESTLER_DATA_END = "// GENERATED WRESTLER DATA END"
PROMOTION_NAMES_START = "// GENERATED PROMOTION NAMES START"
PROMOTION_NAMES_END = "// GENERATED PROMOTION NAMES END"
ROSTER_DATA_START = "// GENERATED ROSTER DATA START"
ROSTER_DATA_END = "// GENERATED ROSTER DATA END"
COMPLETION_ICON = "✅"

class XlsxToHtmlConverter:
    def __init__(self, config_file=None):
        """コンバーターを初期化"""
        self.config = self.load_config(config_file)
    
    def load_config(self, config_file):
        """設定ファイルを読み込み - デフォルトはconverter_config.json"""
        # 設定ファイルが指定されていない場合はデフォルトを使用
        if not config_file:
            config_file = "converter_config.json"
        
        # 設定ファイルが存在しない場合は処理を中止
        if not os.path.exists(config_file):
            raise Exception(f"設定ファイルが見つかりません: {config_file}")
        
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            # 必須項目をチェック
            required_fields = ["xlsx_file", "html_file", "js_variable_name"]
            for field in required_fields:
                if field not in config:
                    raise Exception(f"設定ファイルに必須項目 '{field}' がありません")
            
            # デフォルト値を設定
            config.setdefault("sheet_name", 0)
            config.setdefault("encoding", "utf-8")
            config.setdefault("planner_html_file", "match_card_planner.html")
            
            return config
            
        except json.JSONDecodeError as e:
            raise Exception(f"設定ファイルのJSON形式が正しくありません: {e}")
        except Exception as e:
            raise Exception(f"設定ファイル読み込みエラー: {e}")
    
    @staticmethod
    def replace_generated_block(content, start_marker, end_marker, replacement, label):
        """生成マーカー間を一意に特定して置換する"""
        if content.count(start_marker) != 1 or content.count(end_marker) != 1:
            raise Exception(f"{label}の生成マーカーが一意に見つかりません")

        start = content.index(start_marker) + len(start_marker)
        end = content.index(end_marker, start)
        if start >= end:
            raise Exception(f"{label}の生成マーカー順序が不正です")

        return content[:start] + "\n" + replacement.rstrip() + "\n" + content[end:]

    @staticmethod
    def extract_generated_json(content, start_marker, end_marker, variable_name, label):
        """生成マーカー間のconst代入からJSON値を取り出す"""
        if content.count(start_marker) != 1 or content.count(end_marker) != 1:
            raise Exception(f"{label}の生成マーカーが一意に見つかりません")

        start = content.index(start_marker) + len(start_marker)
        end = content.index(end_marker, start)
        block = content[start:end].strip()
        prefix = f"const {variable_name} = "
        if not block.startswith(prefix) or not block.endswith(";"):
            raise Exception(f"{label}のconst宣言形式が不正です")

        try:
            return json.loads(block[len(prefix):-1])
        except json.JSONDecodeError as e:
            raise Exception(f"{label}のJSON形式が不正です: {e}") from e

    @staticmethod
    def find_header_row(df):
        """デビュー年ヘッダーを一意に特定する"""
        matches = [
            index
            for index, row in df.iterrows()
            if len(row) > 1
            and pd.notna(row.iloc[1])
            and str(row.iloc[1]).strip() == "デビュー年"
        ]
        if len(matches) != 1:
            raise Exception(
                f"デビュー年ヘッダーは1行必要です（検出数: {len(matches)}）"
            )
        return matches[0]

    @staticmethod
    def normalize_parentheses(value):
        """全角の丸括弧を半角へ統一する"""
        return str(value).replace("（", "(").replace("）", ")")

    @staticmethod
    def normalize_year(year, excel_row_number):
        """年度欄を検証して文字列へ正規化する"""
        if isinstance(year, Real) and not isinstance(year, bool):
            numeric_year = float(year)
            if not numeric_year.is_integer():
                raise Exception(
                    f"Excel {excel_row_number}行目の年度が整数ではありません: {year}"
                )
            return str(int(numeric_year))

        year_text = XlsxToHtmlConverter.normalize_parentheses(year).strip()
        if not year_text:
            raise Exception(f"Excel {excel_row_number}行目の年度が空です")
        return year_text

    def read_xlsx_data(self, xlsx_file=None):
        """XLSXファイルからデータを読み込み"""
        file_path = xlsx_file or self.config["xlsx_file"]
        
        try:
            df = pd.read_excel(file_path, sheet_name=self.config["sheet_name"], header=None)
            print(f"XLSXファイル読み込み完了: {file_path}")
            print(f"データサイズ: {df.shape[0]}行 x {df.shape[1]}列")
            return df
        except Exception as e:
            raise Exception(f"XLSXファイル読み込みエラー: {e}")
    
    def extract_promotion_headers(self, df):
        """Excelファイルから団体名のヘッダーを抽出"""
        header_row = self.find_header_row(df)
        
        # ヘッダー行から団体名を抽出（2列目以降）
        promotion_names = []
        header_data = df.iloc[header_row]
        
        for col_idx in range(2, len(header_data)):
            cell_value = header_data.iloc[col_idx]
            promotion_names.append(
                self.normalize_parentheses(cell_value).strip()
                if pd.notna(cell_value)
                else ""
            )

        blank_columns = [
            index + 3 for index, name in enumerate(promotion_names) if not name
        ]
        if blank_columns:
            raise Exception(
                "団体名が空の列があります（Excel列番号: "
                + ", ".join(map(str, blank_columns))
                + "）"
            )

        duplicate_names = sorted(
            name for name, count in Counter(promotion_names).items() if count > 1
        )
        if duplicate_names:
            raise Exception(
                "団体名が重複しています: " + ", ".join(duplicate_names)
            )
        
        print(f"抽出された団体名: {promotion_names}")
        return promotion_names

    def convert_to_js_array(self, df):
        """
        DataFrameをJavaScript配列形式に変換（完全なJSON処理）

        - 年度欄は文字列として扱い、「練習生」などの特殊値にも対応
        - HTML特殊文字(<, >, &, ", ')を自動的にエスケープ処理
          例: <選手名> → &lt;選手名&gt;
        """
        js_array = []
        
        header_row = self.find_header_row(df)
        
        print(f"ヘッダー行: {header_row}")
        
        # データ行を処理
        for index, row in df.iterrows():
            if index <= header_row:
                continue

            year = row.iloc[1]
            if pd.isna(year):
                continue

            year_str = self.normalize_year(year, index + 1)

            # 年度欄もHTMLエスケープ処理を実行
            year_escaped = html.escape(year_str, quote=True)
            row_data = [year_escaped]
            
            # データ列を処理（3列目以降）
            for col_idx in range(2, len(row)):
                cell_value = row.iloc[col_idx]
                
                if pd.isna(cell_value):
                    row_data.append("")
                else:
                    # セルの値を文字列に変換し、HTMLエスケープ処理を実行
                    # <, >, &, ", ' などの特殊文字を &lt;, &gt;, &amp;, &quot;, &#x27; に変換
                    cell_str = self.normalize_parentheses(cell_value).strip()
                    escaped_str = html.escape(cell_str, quote=True)
                    row_data.append(escaped_str)
            
            js_array.append(row_data)
        
        print(f"処理されたデータ行数: {len(js_array)}")
        
        # JavaScript配列を完全なJSONとして生成
        js_array_json = json.dumps(js_array, ensure_ascii=False, indent=12)
        
        # const宣言とセミコロンを追加
        js_string = f"const {self.config['js_variable_name']} = {js_array_json};"
        
        return js_string

    def convert_to_roster_data(self, df, promotion_names):
        """カード検討ツール用の団体別選手データを生成"""
        header_row = self.find_header_row(df)

        roster_data = {name: [] for name in promotion_names if name.strip()}

        for index, row in df.iterrows():
            if index <= header_row or pd.isna(row.iloc[1]):
                continue

            for promotion_index, promotion_name in enumerate(promotion_names):
                if not promotion_name.strip():
                    continue

                column_index = promotion_index + 2
                if column_index >= len(row):
                    continue

                cell_value = row.iloc[column_index]
                if pd.notna(cell_value):
                    # Excelの1セルに改行区切りで複数選手が入っていても、
                    # rosterDataでは「1選手 = 1要素」に正規化する。
                    wrestler_names = (
                        self.normalize_parentheses(name).strip()
                        for name in str(cell_value).splitlines()
                    )
                    roster_data[promotion_name].extend(
                        html.escape(name, quote=True)
                        for name in wrestler_names
                        if name
                    )

        return roster_data

    def validate_source_data(self, df, promotion_names, roster_data):
        """変換前にExcelの構造と重複を検証する"""
        header_row = self.find_header_row(df)
        years = []

        for index, row in df.iterrows():
            if index <= header_row:
                continue

            promotion_cells = row.iloc[2:2 + len(promotion_names)]
            has_wrestler_data = any(
                pd.notna(value) and str(value).strip()
                for value in promotion_cells
            )
            year = row.iloc[1]
            if pd.isna(year):
                if has_wrestler_data:
                    raise Exception(
                        f"Excel {index + 1}行目に年度なしの選手データがあります"
                    )
                continue

            years.append(self.normalize_year(year, index + 1))

        duplicate_years = sorted(
            year for year, count in Counter(years).items() if count > 1
        )
        if duplicate_years:
            raise Exception(
                "年度ラベルが重複しています: " + ", ".join(duplicate_years)
            )

        if not years:
            raise Exception("変換対象の選手データ行がありません")

        duplicate_wrestlers = []
        for promotion_name, wrestlers in roster_data.items():
            duplicates = sorted(
                name for name, count in Counter(wrestlers).items() if count > 1
            )
            duplicate_wrestlers.extend(
                f"{promotion_name}: {name}" for name in duplicates
            )

        if duplicate_wrestlers:
            raise Exception(
                "同一団体内で選手名が重複しています: "
                + ", ".join(duplicate_wrestlers)
            )

    def update_match_card_planner(self, roster_data, planner_file=None):
        """カード検討ツール内の団体別選手データを更新"""
        file_path = planner_file or self.config["planner_html_file"]

        if not os.path.exists(file_path):
            raise Exception(f"カード検討ツールが見つかりません: {file_path}")

        with open(file_path, 'r', encoding=self.config["encoding"]) as f:
            content = f.read()

        # 団体・選手単位でGit差分を確認できるよう、読みやすく整形して出力する
        roster_json = json.dumps(roster_data, ensure_ascii=False, indent=2)
        replacement = f"const rosterData = {roster_json};"
        updated = self.replace_generated_block(
            content,
            ROSTER_DATA_START,
            ROSTER_DATA_END,
            replacement,
            "カード検討ツールのrosterData",
        )

        parsed_roster = self.extract_generated_json(
            updated,
            ROSTER_DATA_START,
            ROSTER_DATA_END,
            "rosterData",
            "カード検討ツールのrosterData",
        )
        if parsed_roster != roster_data:
            raise Exception("カード検討ツールの選手データ生成に失敗しました")

        with open(file_path, 'w', encoding=self.config["encoding"]) as f:
            f.write(updated)

        wrestler_count = sum(len(wrestlers) for wrestlers in roster_data.values())
        print(f"カード検討ツール更新完了: {len(roster_data)}団体 / {wrestler_count}選手")
        return True
    
    def update_html_headers(self, promotion_names, html_file=None):
        """HTMLファイルのヘッダー部分を更新"""
        file_path = html_file or self.config["html_file"]
        
        try:
            # HTMLファイルを読み込み
            with open(file_path, 'r', encoding=self.config["encoding"]) as f:
                html_content = f.read()
            
            # ヘッダー行のパターンを検索
            header_pattern = r'(<tr>\s*<th class="sortable desc"[^>]*>デビュー年</th>\s*)(.*?)(\s*</tr>)'
            
            header_matches = list(re.finditer(header_pattern, html_content, re.DOTALL))
            if len(header_matches) != 1:
                raise Exception(
                    "テーブルヘッダーは1件必要です"
                    f"（検出数: {len(header_matches)}）"
                )
            match = header_matches[0]
            
            # 新しいヘッダーを生成
            new_headers = []
            for i, promotion_name in enumerate(promotion_names):
                escaped_name = html.escape(promotion_name, quote=True)
                header_html = (
                    '<th class="promotion-header" '
                    f'data-promotion-index="{i + 1}" '
                    'role="button" tabindex="0">'
                    f'{escaped_name}</th>'
                )
                new_headers.append(header_html)
            
            # ヘッダー行を置換
            new_header_content = match.group(1) + '\n                                '.join(new_headers) + match.group(3)
            html_content = re.sub(
                header_pattern,
                lambda _: new_header_content,
                html_content,
                count=1,
                flags=re.DOTALL,
            )
            
            # JavaScript配列の団体名も更新
            promotion_names_js = json.dumps(promotion_names, ensure_ascii=False)
            new_promotion_js = f'const promotionNames = {promotion_names_js};'
            html_content = self.replace_generated_block(
                html_content,
                PROMOTION_NAMES_START,
                PROMOTION_NAMES_END,
                new_promotion_js,
                "index.htmlのpromotionNames",
            )

            parsed_names = self.extract_generated_json(
                html_content,
                PROMOTION_NAMES_START,
                PROMOTION_NAMES_END,
                "promotionNames",
                "index.htmlのpromotionNames",
            )
            if parsed_names != promotion_names:
                raise Exception("団体名データの生成に失敗しました")
            
            # HTMLファイルに書き戻し
            with open(file_path, 'w', encoding=self.config["encoding"]) as f:
                f.write(html_content)
            
            print(f"HTMLヘッダー更新完了: {len(promotion_names)}個の団体")
            return html_content
            
        except Exception as e:
            raise Exception(f"HTMLヘッダー更新エラー: {e}")

    def update_html_file(self, js_array_string, html_file=None):
        """HTMLファイルのJavaScript配列を更新"""
        file_path = html_file or self.config["html_file"]
        
        try:
            # HTMLファイルを読み込み
            with open(file_path, 'r', encoding=self.config["encoding"]) as f:
                html_content = f.read()
            
            html_content = self.replace_generated_block(
                html_content,
                WRESTLER_DATA_START,
                WRESTLER_DATA_END,
                js_array_string,
                f"index.htmlの{self.config['js_variable_name']}",
            )
            
            expected_json = json.loads(
                js_array_string.split("=", 1)[1].strip().removesuffix(";")
            )
            parsed_json = self.extract_generated_json(
                html_content,
                WRESTLER_DATA_START,
                WRESTLER_DATA_END,
                self.config["js_variable_name"],
                f"index.htmlの{self.config['js_variable_name']}",
            )
            if parsed_json != expected_json:
                raise Exception(
                    f"{self.config['js_variable_name']}データの生成に失敗しました"
                )

            print(f"既存の{self.config['js_variable_name']}配列を更新しました")
            
            # HTMLファイルに書き戻し
            with open(file_path, 'w', encoding=self.config["encoding"]) as f:
                f.write(html_content)
            
            print(f"HTMLファイル更新完了: {file_path}")
            return html_content
            
        except Exception as e:
            raise Exception(f"HTMLファイル更新エラー: {e}")
    
    def verify_outputs(
        self,
        html_file,
        planner_file,
        expected_js_array,
        promotion_names,
        roster_data,
    ):
        """生成した2つのHTMLをJSONとして再読込し、相互整合性を検証する"""
        with open(html_file, 'r', encoding=self.config["encoding"]) as f:
            html_content = f.read()
        with open(planner_file, 'r', encoding=self.config["encoding"]) as f:
            planner_content = f.read()

        wrestler_data = self.extract_generated_json(
            html_content,
            WRESTLER_DATA_START,
            WRESTLER_DATA_END,
            self.config["js_variable_name"],
            f"index.htmlの{self.config['js_variable_name']}",
        )
        saved_promotion_names = self.extract_generated_json(
            html_content,
            PROMOTION_NAMES_START,
            PROMOTION_NAMES_END,
            "promotionNames",
            "index.htmlのpromotionNames",
        )
        saved_roster_data = self.extract_generated_json(
            planner_content,
            ROSTER_DATA_START,
            ROSTER_DATA_END,
            "rosterData",
            "カード検討ツールのrosterData",
        )
        expected_wrestler_data = json.loads(
            expected_js_array.split("=", 1)[1].strip().removesuffix(";")
        )

        if wrestler_data != expected_wrestler_data:
            raise Exception("index.htmlの選手データがExcel変換結果と一致しません")
        if saved_promotion_names != promotion_names:
            raise Exception("index.htmlの団体名がExcel変換結果と一致しません")
        if saved_roster_data != roster_data:
            raise Exception("カード検討ツールの選手データがExcel変換結果と一致しません")

        derived_roster = {
            promotion_name: [
                wrestler.strip()
                for row in wrestler_data
                for wrestler in str(row[index + 1] or "").splitlines()
                if wrestler.strip()
            ]
            for index, promotion_name in enumerate(promotion_names)
        }
        if derived_roster != saved_roster_data:
            raise Exception("2つのHTML間で団体別選手データが一致しません")

        header_indexes = [
            int(value)
            for value in re.findall(
                r'<th class="promotion-header" data-promotion-index="(\d+)"',
                html_content,
            )
        ]
        expected_indexes = list(range(1, len(promotion_names) + 1))
        if header_indexes != expected_indexes:
            raise Exception("団体ヘッダーの列番号がExcelの列順と一致しません")

        print(
            "生成結果検証完了: "
            f"{len(promotion_names)}団体 / "
            f"{sum(len(names) for names in roster_data.values())}選手"
        )
    
    def convert(self, xlsx_file=None, html_file=None):
        """XLSXからHTMLへの完全変換ワークフローを実行"""
        try:
            print("=== XLSX to HTML 完全変換ワークフロー開始 ===")
            print()
            
            target_file = html_file or self.config["html_file"]
            planner_file = self.config["planner_html_file"]
            for label, file_path in (
                ("一覧HTML", target_file),
                ("カード検討ツール", planner_file),
            ):
                if not os.path.exists(file_path):
                    raise Exception(f"{label}が見つかりません: {file_path}")

            # 書き込み前に全置換対象が揃っていることを検証する
            with open(target_file, 'r', encoding=self.config["encoding"]) as f:
                html_template = f.read()
            with open(planner_file, 'r', encoding=self.config["encoding"]) as f:
                planner_template = f.read()

            self.extract_generated_json(
                html_template,
                WRESTLER_DATA_START,
                WRESTLER_DATA_END,
                self.config["js_variable_name"],
                f"index.htmlの{self.config['js_variable_name']}",
            )
            self.extract_generated_json(
                html_template,
                PROMOTION_NAMES_START,
                PROMOTION_NAMES_END,
                "promotionNames",
                "index.htmlのpromotionNames",
            )
            self.extract_generated_json(
                planner_template,
                ROSTER_DATA_START,
                ROSTER_DATA_END,
                "rosterData",
                "カード検討ツールのrosterData",
            )
            header_pattern = (
                r'(<tr>\s*<th class="sortable desc"[^>]*>'
                r'デビュー年</th>\s*)(.*?)(\s*</tr>)'
            )
            if len(re.findall(header_pattern, html_template, re.DOTALL)) != 1:
                raise Exception("一覧HTMLのテーブルヘッダーが一意に見つかりません")

            # ステップ1: XLSXデータを読み込み
            print("ステップ1: XLSXデータを検証・変換")
            df = self.read_xlsx_data(xlsx_file)
            promotion_names = self.extract_promotion_headers(df)
            js_array = self.convert_to_js_array(df)
            roster_data = self.convert_to_roster_data(df, promotion_names)
            self.validate_source_data(df, promotion_names, roster_data)
            
            # HTMLファイルを更新
            self.update_html_file(js_array, html_file)
            print("データ変換完了")
            print()
            
            # ステップ1.5: HTMLヘッダーを更新
            print("ステップ2: HTMLヘッダーをExcelに合わせて更新")
            self.update_html_headers(promotion_names, html_file)
            print("ヘッダー更新完了")
            print()
            
            # カード検討ツールも同じ選手データへ更新
            print("ステップ3: カード検討ツールの選手データ更新")
            self.update_match_card_planner(roster_data, planner_file)
            print("カード検討ツール更新完了")
            print()
            
            print("ステップ4: 生成結果の完全比較")
            self.verify_outputs(
                target_file,
                planner_file,
                js_array,
                promotion_names,
                roster_data,
            )
            print()

            print(f"{COMPLETION_ICON} 完全変換ワークフロー完了")
            print()
            print(f"出力ファイル: {target_file}")
            print(f"カード検討ツール: {planner_file}")
            
            return True
            
        except Exception as e:
            print(f"変換エラー: {e}")
            return False

def create_sample_config():
    """設定ファイルのサンプルを作成"""
    sample_config = {
        "xlsx_file": "woman-excel.xlsx",
        "html_file": "index.html",
        "planner_html_file": "match_card_planner.html",
        "sheet_name": 0,
        "js_variable_name": "wrestlerData",
        "encoding": "utf-8"
    }
    
    config_file = "converter_config.json"
    with open(config_file, 'w', encoding='utf-8') as f:
        json.dump(sample_config, f, indent=4, ensure_ascii=False)
    
    print(f"設定ファイルのサンプルを作成しました: {config_file}")

def main():
    """メイン関数"""
    parser = argparse.ArgumentParser(
        description="XLSX to HTML Converter (All-in-One)",
        epilog="""
使用例:
  python3 xlsx_to_html_converter_all_in_one.py
  python3 xlsx_to_html_converter_all_in_one.py --xlsx data.xlsx --html output.html
  python3 xlsx_to_html_converter_all_in_one.py --config config.json
  python3 xlsx_to_html_converter_all_in_one.py --create-config
        """,
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument("--xlsx", help="XLSXファイルのパス")
    parser.add_argument("--html", help="HTMLファイルのパス")
    parser.add_argument("--config", help="設定ファイルのパス")
    parser.add_argument("--create-config", action="store_true", help="設定ファイルのサンプルを作成")
    
    args = parser.parse_args()
    
    if args.create_config:
        create_sample_config()
        return
    
    try:
        converter = XlsxToHtmlConverter(args.config)
        success = converter.convert(args.xlsx, args.html)
        
        if not success:
            print("変換に失敗しました")
            sys.exit(1)
        
    except Exception as e:
        print(f"致命的エラー: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
