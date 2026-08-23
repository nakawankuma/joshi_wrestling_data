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
        header_row = None
        for index, row in df.iterrows():
            if pd.notna(row.iloc[1]) and str(row.iloc[1]).strip() == "デビュー年":
                header_row = index
                break
        
        if header_row is None:
            print("ヘッダー行が見つかりません")
            return []
        
        # ヘッダー行から団体名を抽出（2列目以降）
        promotion_names = []
        header_data = df.iloc[header_row]
        
        for col_idx in range(2, len(header_data)):
            cell_value = header_data.iloc[col_idx]
            if pd.notna(cell_value):
                promotion_name = str(cell_value).strip()
                promotion_names.append(promotion_name)
            else:
                promotion_names.append("")
        
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
        
        # ヘッダー行を見つける
        header_row = None
        for index, row in df.iterrows():
            if pd.notna(row.iloc[1]) and str(row.iloc[1]).strip() == "デビュー年":
                header_row = index
                break
        
        if header_row is None:
            print("ヘッダー行が見つかりません")
            return f"const {self.config['js_variable_name']} = [];"
        
        print(f"ヘッダー行: {header_row}")
        
        # データ行を処理
        for index, row in df.iterrows():
            if index <= header_row:
                continue

            year = row.iloc[1]
            if pd.isna(year):
                continue

            # 行データを作成（年度は文字列として扱う - 練習生などの特殊値にも対応）
            # 数値の場合は文字列に変換、既に文字列の場合はそのまま使用
            if isinstance(year, (int, float)):
                year_str = str(int(year))
            else:
                year_str = str(year).strip()

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
                    cell_str = str(cell_value).strip()
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
        header_row = None
        for index, row in df.iterrows():
            if pd.notna(row.iloc[1]) and str(row.iloc[1]).strip() == "デビュー年":
                header_row = index
                break

        if header_row is None:
            raise Exception("カード検討ツール用データのヘッダー行が見つかりません")

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
                    roster_data[promotion_name].append(
                        html.escape(str(cell_value).strip(), quote=True)
                    )

        return roster_data

    def update_match_card_planner(self, roster_data, planner_file=None):
        """カード検討ツール内の団体別選手データを更新"""
        file_path = planner_file or self.config["planner_html_file"]

        if not os.path.exists(file_path):
            raise Exception(f"カード検討ツールが見つかりません: {file_path}")

        with open(file_path, 'r', encoding=self.config["encoding"]) as f:
            content = f.read()

        pattern = r'const\s+rosterData\s*=\s*\{[\s\S]*?\};'
        if len(re.findall(pattern, content)) != 1:
            raise Exception("カード検討ツールのrosterDataが一意に見つかりません")

        roster_json = json.dumps(roster_data, ensure_ascii=False, separators=(',', ':'))
        replacement = f"const rosterData = {roster_json};"
        # 置換文字列内の \n やバックスラッシュをre.subに解釈させない
        updated = re.sub(pattern, lambda _: replacement, content, count=1)

        with open(file_path, 'w', encoding=self.config["encoding"]) as f:
            f.write(updated)

        # 書き戻したJSONを再読込し、欠落や破損がないことを確認
        with open(file_path, 'r', encoding=self.config["encoding"]) as f:
            saved = f.read()
        match = re.search(r'const\s+rosterData\s*=\s*(\{[\s\S]*?\});', saved)
        if not match or json.loads(match.group(1)) != roster_data:
            raise Exception("カード検討ツールの選手データ検証に失敗しました")

        wrestler_count = sum(
            len(cell.splitlines()) for cells in roster_data.values() for cell in cells
        )
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
            
            match = re.search(header_pattern, html_content, re.DOTALL)
            if not match:
                print("警告: ヘッダー行が見つかりませんでした")
                return html_content
            
            # 新しいヘッダーを生成
            new_headers = []
            for i, promotion_name in enumerate(promotion_names):
                if promotion_name.strip():
                    header_html = f'<th class="promotion-header" onclick="filterByPromotion(\'{promotion_name}\', {i+1})">{promotion_name}</th>'
                    new_headers.append(header_html)
            
            # ヘッダー行を置換
            new_header_content = match.group(1) + '\n                                '.join(new_headers) + match.group(3)
            html_content = re.sub(header_pattern, new_header_content, html_content, flags=re.DOTALL)
            
            # JavaScript配列の団体名も更新
            promotion_names_js = json.dumps(promotion_names, ensure_ascii=False)
            promotion_pattern = r'const promotionNames = \[.*?\];'
            new_promotion_js = f'const promotionNames = {promotion_names_js};'
            html_content = re.sub(promotion_pattern, new_promotion_js, html_content, flags=re.DOTALL)
            
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
            
            # 既存のデータ配列を検索して置換
            pattern = r'const\s+' + re.escape(self.config["js_variable_name"]) + r'\s*=\s*\[[\s\S]*?\];'
            
            if re.search(pattern, html_content):
                # 既存の配列を置換
                html_content = re.sub(pattern, js_array_string, html_content)
                print(f"既存の{self.config['js_variable_name']}配列を更新しました")
            else:
                print(f"警告: {self.config['js_variable_name']}配列が見つかりませんでした")
                print("生成されたJavaScriptコード:")
                print(js_array_string)
                return js_array_string
            
            # HTMLファイルに書き戻し
            with open(file_path, 'w', encoding=self.config["encoding"]) as f:
                f.write(html_content)
            
            print(f"HTMLファイル更新完了: {file_path}")
            return html_content
            
        except Exception as e:
            raise Exception(f"HTMLファイル更新エラー: {e}")
    
    def fix_newlines_in_html(self, file_path=None):
        """HTMLファイル内の改行文字を\\nにエスケープ"""
        target_file = file_path or self.config["html_file"]
        
        print("🔧 改行文字の修正を開始します...")
        
        with open(target_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # wrestlerData配列内の文字列を修正
        def fix_string_newlines(match):
            full_match = match.group(0)
            # クォート内の改行文字を\\nに置換
            fixed = full_match.replace('\n', '\\n').replace('\r', '\\r')
            return fixed
        
        # 文字列パターンを検索（改行を含む可能性のある文字列）
        pattern = r'"[^"]*(?:\n[^"]*)*"'
        
        # wrestlerData配列の部分を特定
        start_marker = f'const {self.config["js_variable_name"]} = ['
        end_marker = '];'
        
        start_idx = content.find(start_marker)
        if start_idx == -1:
            print(f"❌ {self.config['js_variable_name']}配列が見つかりません")
            return False
        
        end_idx = content.find(end_marker, start_idx) + len(end_marker)
        if end_idx == len(end_marker) - 1:
            print(f"❌ {self.config['js_variable_name']}配列の終了が見つかりません")
            return False
        
        # 配列部分を抽出
        before_array = content[:start_idx]
        array_part = content[start_idx:end_idx]
        after_array = content[end_idx:]
        
        print(f"配列部分を抽出: {len(array_part)}文字")
        
        # 配列部分の改行文字を修正
        fixed_array = re.sub(pattern, fix_string_newlines, array_part)
        
        # 修正されたコンテンツを作成
        fixed_content = before_array + fixed_array + after_array
        
        # ファイルに書き戻し
        with open(target_file, 'w', encoding='utf-8') as f:
            f.write(fixed_content)
        
        print(f"✅ 改行文字の修正完了: {target_file}")
        
        # 修正数をカウント
        original_newlines = array_part.count('\n"') + array_part.count('"\n')
        fixed_newlines = fixed_array.count('\\n')
        
        print(f"修正された改行文字: {fixed_newlines}個")
        
        return True
    
    def verify_fix(self, file_path=None):
        """修正結果を検証"""
        target_file = file_path or self.config["html_file"]
        
        try:
            with open(target_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # JavaScript配列部分をチェック
            start_idx = content.find(f'const {self.config["js_variable_name"]} = [')
            end_idx = content.find('];', start_idx) + 2
            
            if start_idx != -1 and end_idx > start_idx:
                array_section = content[start_idx:end_idx]
                
                # 生の改行文字がまだ残っているかチェック
                problem_lines = []
                lines = array_section.split('\n')
                for i, line in enumerate(lines):
                    if '"' in line and line.count('"') % 2 == 1:  # 開始クォートのみ
                        problem_lines.append(i)
                
                if problem_lines:
                    print(f"❌ まだ問題のある行があります: {len(problem_lines)}行")
                    return False
                else:
                    print("✅ 改行文字の問題は解決されました")
                    return True
            else:
                print("❌ 配列が見つかりません")
                return False
                
        except Exception as e:
            print(f"❌ 検証エラー: {e}")
            return False
    
    def convert(self, xlsx_file=None, html_file=None):
        """XLSXからHTMLへの完全変換ワークフローを実行"""
        try:
            print("=== XLSX to HTML 完全変換ワークフロー開始 ===")
            print()
            
            # ステップ1: XLSXデータを読み込み
            print("📋 ステップ1: XLSXデータをHTMLに変換")
            df = self.read_xlsx_data(xlsx_file)
            
            # 団体名ヘッダーを抽出
            promotion_names = self.extract_promotion_headers(df)
            
            # JavaScript配列に変換
            js_array = self.convert_to_js_array(df)

            # カード検討ツール用データも同じExcelから生成
            roster_data = self.convert_to_roster_data(df, promotion_names)
            planner_file = self.config["planner_html_file"]
            if not os.path.exists(planner_file):
                raise Exception(f"カード検討ツールが見つかりません: {planner_file}")
            with open(planner_file, 'r', encoding=self.config["encoding"]) as f:
                planner_content = f.read()
            planner_pattern = r'const\s+rosterData\s*=\s*\{[\s\S]*?\};'
            if len(re.findall(planner_pattern, planner_content)) != 1:
                raise Exception("カード検討ツールのrosterDataが一意に見つかりません")
            
            # HTMLファイルを更新
            self.update_html_file(js_array, html_file)
            print("✅ データ変換完了")
            print()
            
            # ステップ1.5: HTMLヘッダーを更新
            print("🏷️ ステップ1.5: HTMLヘッダーをExcelに合わせて更新")
            self.update_html_headers(promotion_names, html_file)
            print("✅ ヘッダー更新完了")
            print()
            
            # ステップ2: 改行文字を修正
            print("🔧 ステップ2: 改行文字の修正")
            if self.fix_newlines_in_html(html_file):
                print("✅ 改行文字修正完了")
            else:
                print("❌ 改行文字修正失敗")
                return False
            print()
            
            # ステップ3: 検証
            print("🔍 ステップ3: 結果検証")
            if self.verify_fix(html_file):
                print("✅ 検証成功")
            else:
                print("❌ 検証失敗")
                return False
            print()

            # ステップ4: カード検討ツールも同じ選手データへ更新
            print("🃏 ステップ4: カード検討ツールの選手データ更新")
            self.update_match_card_planner(roster_data, planner_file)
            print("✅ カード検討ツール更新・検証完了")
            print()
            
            print("🎉 完全変換ワークフロー完了！")
            print()
            target_file = html_file or self.config["html_file"]
            print(f"📁 出力ファイル: {target_file}")
            print(f"📁 カード検討ツール: {planner_file}")
            print("🔁 このスクリプトは完全に再現性があります")
            
            return True
            
        except Exception as e:
            print(f"❌ 変換エラー: {e}")
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
            print("💥 変換に失敗しました")
            sys.exit(1)
        
    except Exception as e:
        print(f"💥 致命的エラー: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
