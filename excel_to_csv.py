#!/usr/bin/env python3
"""
Excel → CSV変換スクリプト
woman-excel.xlsx のピボットテーブル形式データを
1選手1行のフラットCSVに変換する

出力列:
  選手名, 呼びがな姓名, 呼びがな姓のみ, デビュー年, デビュー団体,
  所属団体, 引退フラグ, 補足, twitter, wikipedia, instagram, tiktok
"""

import pandas as pd
import csv
import sys

INPUT_FILE = "woman-excel.xlsx"
OUTPUT_FILE = "wrestlers.csv"

CSV_COLUMNS = [
    "選手名",
    "呼びがな姓名",
    "呼びがな姓のみ",
    "デビュー年",
    "デビュー団体",
    "所属団体",
    "引退フラグ",
    "補足",
    "twitter",
    "wikipedia",
    "instagram",
    "tiktok",
]


def find_header_row(df):
    """'デビュー年'が含まれるヘッダー行のインデックスを返す"""
    for index, row in df.iterrows():
        if pd.notna(row.iloc[1]) and str(row.iloc[1]).strip() == "デビュー年":
            return index
    return None


def extract_promotion_names(df, header_row):
    """ヘッダー行から団体名リストを抽出（列2以降）"""
    promotions = []
    header_data = df.iloc[header_row]
    for col_idx in range(2, len(header_data)):
        cell = header_data.iloc[col_idx]
        promotions.append(str(cell).strip() if pd.notna(cell) else "")
    return promotions


def parse_year(value):
    """年度値を文字列に変換（数値・文字列・練習生などに対応）"""
    if isinstance(value, (int, float)):
        return str(int(value))
    return str(value).strip()


def parse_wrestler_name(raw_name, column_promotion):
    """
    括弧内の情報を列の種類に応じて振り分ける

    「その他」列の場合:
      '堀田祐美子(T-HEARTS)' → 選手名='堀田祐美子', 所属団体='T-HEARTS', 補足=''
    それ以外の列の場合:
      '選手名（休）' → 選手名='選手名', 所属団体=列の団体名, 補足='休'
    括弧がない場合はそのまま返す
    """
    import re
    # 全角・半角括弧どちらにも対応
    match = re.search(r'(.+?)[（(](.+?)[）)]\s*$', raw_name)
    if not match:
        return raw_name, column_promotion, ""

    name = match.group(1).strip()
    paren_content = match.group(2).strip()

    if column_promotion == "その他":
        # 括弧内は団体名
        return name, paren_content, ""
    else:
        # 括弧内は補足情報
        return name, column_promotion, paren_content


def extract_wrestlers(df, header_row, promotion_names):
    """
    データ行を走査し、選手データのリストを返す
    各要素: dict（CSV列名→値）
    """
    wrestlers = []
    warnings = []

    for index, row in df.iterrows():
        # ヘッダー行以前はスキップ
        if index <= header_row:
            continue

        year_raw = row.iloc[1]
        if pd.isna(year_raw):
            continue

        year = parse_year(year_raw)

        # 団体列（列2以降）を走査
        for col_offset, promotion in enumerate(promotion_names):
            col_idx = 2 + col_offset
            if col_idx >= len(row):
                break

            cell = row.iloc[col_idx]
            if pd.isna(cell) or str(cell).strip() == "":
                continue

            # セル内の選手名を改行で分割
            names = [n.strip() for n in str(cell).split("\n") if n.strip()]

            for name in names:
                # 括弧内を解析:「その他」列は団体名、他の列は補足情報として扱う
                parsed_name, actual_promotion, note = parse_wrestler_name(name, promotion)
                wrestlers.append({
                    "選手名": parsed_name,
                    "呼びがな姓名": "",
                    "呼びがな姓のみ": "",
                    "デビュー年": year,
                    "デビュー団体": "",
                    "所属団体": actual_promotion,
                    "引退フラグ": "false",
                    "補足": note,
                    "twitter": "",
                    "wikipedia": "",
                    "instagram": "",
                    "tiktok": "",
                })

    return wrestlers, warnings


def main():
    print(f"読み込み中: {INPUT_FILE}")

    try:
        df = pd.read_excel(INPUT_FILE, sheet_name=0, header=None)
    except Exception as e:
        print(f"ERROR: Excelファイルの読み込みに失敗しました: {e}")
        sys.exit(1)

    print(f"データサイズ: {df.shape[0]}行 x {df.shape[1]}列")

    # ヘッダー行を特定
    header_row = find_header_row(df)
    if header_row is None:
        print("ERROR: ヘッダー行（'デビュー年'）が見つかりません")
        sys.exit(1)
    print(f"ヘッダー行: {header_row}行目")

    # 団体名を抽出
    promotion_names = extract_promotion_names(df, header_row)
    print(f"団体数: {len(promotion_names)} → {promotion_names}")

    # 選手データを抽出
    wrestlers, warnings = extract_wrestlers(df, header_row, promotion_names)

    # ワーニング出力
    for w in warnings:
        print(f"WARNING: {w}")

    print(f"\n抽出選手数: {len(wrestlers)}人")

    # CSV出力
    with open(OUTPUT_FILE, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS)
        writer.writeheader()
        writer.writerows(wrestlers)

    print(f"出力完了: {OUTPUT_FILE}")
    print("\n※ 以下の列は手動で補完が必要です:")
    print("  呼びがな姓名, 呼びがな姓のみ, デビュー団体, 引退フラグ, 補足, twitter, wikipedia, instagram, tiktok")


if __name__ == "__main__":
    main()
