#!/usr/bin/env python3
"""
CSV → HTML コンバーター
wrestlers.csv のフラットデータから index.html の JS データを更新する

処理内容:
  1. wrestlers.csv を読み込み
  2. promotion_state.json から前回の団体別選手数を取得
  3. 現在の現役選手数を集計し、団体表示順を決定
     - 4名以上の団体を独立列として表示
     - 3名以下の団体はその他列に「選手名(団体名)」形式で表示
     - 減少中の団体（例:5→4）を左寄り、増加中の団体（例:3→4）を右寄りに配置
  4. 年×団体のマトリクス形式で wrestlerData を生成
  5. 選手ごとのSNS情報を wrestlerMeta として生成
  6. index.html の該当箇所を更新
  7. promotion_state.json を更新
"""

import csv
import json
import re
import sys
from datetime import datetime, timezone
from collections import defaultdict

INPUT_CSV = "wrestlers.csv"
STATE_FILE = "promotion_state.json"
HTML_FILE = "index.html"

# 独立列として表示する最低選手数
MIN_WRESTLERS_FOR_COLUMN = 4

# その他列のキー名（CSVの所属団体値とも一致させる）
SONOTA_COLUMN = "その他"


# ──────────────────────────────────────────
# CSV 読み込み
# ──────────────────────────────────────────

def load_csv(path):
    """wrestlers.csv を読み込み、辞書のリストを返す"""
    try:
        with open(path, encoding="utf-8-sig") as f:
            return list(csv.DictReader(f))
    except FileNotFoundError:
        print(f"ERROR: {path} が見つかりません")
        sys.exit(1)


def parse_retired_flag(value, row_num, name):
    """
    引退フラグを bool に変換する
    大文字小文字・前後空白を正規化した上で true/false のみ受け付ける
    それ以外はワーニングを出力し None を返す
    """
    normalized = str(value).strip().lower()
    if normalized == "true":
        return True
    if normalized == "false":
        return False
    print(f"WARNING: 行{row_num} '{name}' の引退フラグ '{value}' は true/false 以外の値です")
    return None


# ──────────────────────────────────────────
# 団体表示順の計算
# ──────────────────────────────────────────

def count_active_per_promotion(wrestlers):
    """現役選手（引退フラグ=false）の団体別人数を集計"""
    counts = defaultdict(int)
    for w in wrestlers:
        if not w["_retired"]:
            counts[w["所属団体"]] += 1
    return dict(counts)


def load_promotion_state(path):
    """promotion_state.json を読み込む。存在しない場合は空の状態を返す"""
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return {"promotions": {}}


def save_promotion_state(path, promotions_order, current_counts):
    """promotion_state.json を更新する"""
    state = {
        "updated_at": datetime.now(timezone.utc).isoformat(),
        "promotions": {}
    }
    for order_idx, name in enumerate(promotions_order):
        state["promotions"][name] = {
            "count": current_counts.get(name, 0),
            "order": order_idx + 1
        }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)
    print(f"promotion_state.json を更新しました")


# 末尾に固定配置する団体（この順序で末尾に並ぶ）
FIXED_TAIL_COLUMNS = ["海外", SONOTA_COLUMN, "フリー"]


def determine_promotion_order(current_counts, prev_state):
    """
    団体の表示列順を決定する

    ルール:
      1. 4名以上の現役選手がいる団体のみ独立列として表示
      2. 基本順: 現役選手数の多い順（左→右）
      3. 同数内の優先順位:
           5→4（減少して最低ラインに）は左寄り（trend=0）
           変化なし                    は中央  （trend=1）
           3→4（増加して新規参入）     は右寄り（trend=2）
      4. 海外・その他・フリーは常に末尾に固定（この順序）

    戻り値: 表示する団体名のリスト（末尾固定列を含む）
    """
    prev_promotions = prev_state.get("promotions", {})

    # 末尾固定列と通常列を分離
    variable_counts = {
        p: c for p, c in current_counts.items()
        if c >= MIN_WRESTLERS_FOR_COLUMN and p not in FIXED_TAIL_COLUMNS
    }

    def sort_key(promotion):
        """
        主キー: 現役選手数の多い順（降順なので負値）
        副キー: トレンド（0=減少→左, 1=安定→中, 2=増加→右）
        """
        curr = current_counts.get(promotion, 0)
        prev_info = prev_promotions.get(promotion, {})
        prev_count = prev_info.get("count", curr)

        if curr < prev_count:
            trend = 0   # 減少中（例: 5→4）→ 同数内で左
        elif curr > prev_count:
            trend = 2   # 増加中（例: 3→4）→ 同数内で右
        else:
            trend = 1   # 安定

        return (-curr, trend)

    ordered = sorted(variable_counts.keys(), key=sort_key)

    # 末尾固定列を追加（固定順序を維持）
    # その他列は小規模団体の受け皿になるため選手が直接所属していなくても常に追加する
    # 海外・フリーは選手がいる場合のみ追加する
    for col in FIXED_TAIL_COLUMNS:
        if col == SONOTA_COLUMN or col in current_counts:
            ordered.append(col)

    return ordered


# ──────────────────────────────────────────
# マトリクス生成
# ──────────────────────────────────────────

def build_matrix(wrestlers, promotions_order):
    """
    年×団体のマトリクスを生成する

    - promotions_order に含まれない団体（3名以下）の選手は
      「選手名(団体名)」形式でその他列に移動
    - 引退選手も含める（表示側で区別可能なように引退フラグを名前に付与しない）
    - 戻り値: {year: {promotion: [name, ...]}}
    """
    matrix = defaultdict(lambda: defaultdict(list))
    promotion_set = set(promotions_order)

    for w in wrestlers:
        year = w["デビュー年"]
        promotion = w["所属団体"]
        name = w["選手名"]

        if promotion in promotion_set:
            matrix[year][promotion].append(name)
        else:
            # 独立列に満たない団体（3名以下）→ その他列に「選手名(団体名)」で表示
            # 元々その他列の選手はそのまま表示
            if promotion == SONOTA_COLUMN:
                matrix[year][SONOTA_COLUMN].append(name)
            else:
                matrix[year][SONOTA_COLUMN].append(f"{name}({promotion})")

    return matrix


def matrix_to_js_array(matrix, promotions_order):
    """
    マトリクスを HTML の wrestlerData 形式（JavaScript配列）に変換する

    形式: [[year, org1_names, org2_names, ...], ...]
    promotions_order には末尾固定列（海外・その他・フリー）も含まれる
    セル内の複数選手は \\n 区切り
    """
    rows = []
    for year in sorted(matrix.keys(), key=lambda y: (y.isdigit(), int(y) if y.isdigit() else y)):
        row = [year]
        for col in promotions_order:
            names = matrix[year].get(col, [])
            row.append("\\n".join(names) if names else "")
        rows.append(row)

    return rows


# ──────────────────────────────────────────
# wrestlerMeta 生成（SNSリンク用辞書）
# ──────────────────────────────────────────

def build_wrestler_meta(wrestlers):
    """
    選手名 → SNS情報 の辞書を生成する
    同名選手が複数いる場合は後勝ち（警告を出す）
    """
    meta = {}
    for w in wrestlers:
        name = w["選手名"]
        if name in meta:
            print(f"WARNING: 選手名 '{name}' が重複しています")
        entry = {}
        if w.get("twitter"):
            entry["twitter"] = w["twitter"]
        if w.get("wikipedia"):
            entry["wikipedia"] = w["wikipedia"]
        if w.get("instagram"):
            entry["instagram"] = w["instagram"]
        if w.get("tiktok"):
            entry["tiktok"] = w["tiktok"]
        if entry:
            meta[name] = entry
    return meta


# ──────────────────────────────────────────
# HTML 更新
# ──────────────────────────────────────────


def update_html(html_path, js_array_rows, promotions_order, wrestler_meta):
    """
    index.html の以下の箇所を更新する:
      - const wrestlerData = [...]
      - const promotionNames = [...]
      - const wrestlerMeta = {...}   ← 新規追加または更新
      - テーブルの団体ヘッダー行
    """
    with open(html_path, encoding="utf-8") as f:
        content = f.read()

    # --- wrestlerData を更新（ブラケット対応で確実に範囲を特定）---
    data_json = json.dumps(js_array_rows, ensure_ascii=False, indent=12)
    MARKER = "const wrestlerData = "
    start_pos = content.find(MARKER)
    if start_pos == -1:
        print("ERROR: wrestlerData の更新に失敗しました（'const wrestlerData = ' が見つかりません）")
        sys.exit(1)
    # '[' の位置を特定し、対応する ']' をブラケット数で探す
    bracket_start = content.index("[", start_pos)
    depth = 0
    bracket_end = bracket_start
    for i in range(bracket_start, len(content)):
        if content[i] == "[":
            depth += 1
        elif content[i] == "]":
            depth -= 1
            if depth == 0:
                bracket_end = i
                break
    # '[' から ']' の次の ';' までを置換
    semicolon_pos = content.index(";", bracket_end)
    content = content[:start_pos] + MARKER + data_json + ";" + content[semicolon_pos + 1:]
    print("    wrestlerData 更新OK")

    # --- promotionNames を更新 ---
    # 1行に収まる前提で re.sub を使用（DOTALL 不要）
    names_json = json.dumps(promotions_order, ensure_ascii=False)
    new_content = re.sub(
        r'const promotionNames = \[[^\]]*\];',
        f'const promotionNames = {names_json};',
        content
    )
    if new_content == content:
        print("ERROR: promotionNames の更新に失敗しました（パターンが見つかりません）")
        sys.exit(1)
    content = new_content
    print("    promotionNames 更新OK")

    # --- wrestlerMeta を更新または挿入 ---
    meta_json = json.dumps(wrestler_meta, ensure_ascii=False, indent=8)
    meta_js = f'const wrestlerMeta = {meta_json};'
    if 'const wrestlerMeta = {' in content:
        # 既存の wrestlerMeta を置換（END マーカーとして "\n};" を使用）
        new_content = re.sub(
            r'const wrestlerMeta = \{[\s\S]*?\n\};',
            meta_js,
            content
        )
        if new_content == content:
            print("WARNING: wrestlerMeta の更新に失敗しました（パターン不一致）")
        else:
            content = new_content
            print("    wrestlerMeta 更新OK")
    else:
        # 存在しない場合は promotionNames の直後に挿入
        new_content = re.sub(
            r'(const promotionNames = \[[^\]]*\];)',
            r'\1\n\n        ' + meta_js,
            content
        )
        if new_content == content:
            print("WARNING: wrestlerMeta の挿入に失敗しました")
        else:
            content = new_content
            print("    wrestlerMeta 挿入OK")

    # --- テーブルの団体ヘッダー行を更新 ---
    new_headers = "\n".join(
        f'                                <th class="promotion-header" '
        f'onclick="filterByPromotion(\'{p}\', {i+1})">{p}</th>'
        for i, p in enumerate(promotions_order)
    )
    header_pattern = (
        r'(<tr>\s*<th class="sortable[^"]*"[^>]*>デビュー年</th>\s*)'
        r'([\s\S]*?)'
        r'(\s*</tr>)'
    )
    match = re.search(header_pattern, content)
    if match:
        content = content[:match.start()] \
            + match.group(1) \
            + "\n" + new_headers + "\n                            " \
            + match.group(3) \
            + content[match.end():]
        print("    テーブルヘッダー更新OK")
    else:
        print("ERROR: テーブルヘッダー行が見つかりませんでした")
        sys.exit(1)

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(content)

    print(f"    index.html 書き込み完了")


# ──────────────────────────────────────────
# wrestlerNameWithTwitter の更新
# （wrestlerMeta を使うよう HTML の JS 関数を置き換え）
# ──────────────────────────────────────────

NEW_WRESTLER_NAME_WITH_TWITTER = '''\
        // 選手名にSNSリンクアイコンを付与するヘルパー
        // wrestlerMeta に実際のアカウント情報があればそれを使用し、
        // なければ選手名でX/Wikipedia検索URLを生成する
        function wrestlerNameWithTwitter(name) {
            const trimmed = name.trim();
            if (!trimmed) return \\'\\';
            // 括弧とその中身を除去して純粋な選手名を取得
            // 例: '堀田祐美子(T-HEARTS)' → '堀田祐美子'
            const plainName = trimmed
                .replace(/<[^>]*>/g, \\'\\')
                .replace(/[（(][^）)]*[）)]/g, \\'\\')
                .replace(/&lt;|&gt;|[<>]/g, \\'\\')
                .trim();

            const meta = wrestlerMeta[plainName] || {};

            // X(Twitter)リンク: アカウントがあればプロフィールへ、なければ検索
            const xUrl = meta.twitter
                ? `https://x.com/${meta.twitter}`
                : `https://x.com/search?q=${encodeURIComponent(plainName)}&src=typed_query`;

            // Wikipediaリンク: 記事名があれば記事へ、なければ検索
            const wikiUrl = meta.wikipedia
                ? `https://ja.wikipedia.org/wiki/${encodeURIComponent(meta.wikipedia)}`
                : `https://ja.wikipedia.org/wiki/Special:Search?search=${encodeURIComponent(plainName)}`;

            let html = trimmed
                + `<a href="${xUrl}" target="_blank" rel="noopener noreferrer" class="twitter-link" title="${plainName}をXで検索">𝕏</a>`
                + `<a href="${wikiUrl}" target="_blank" rel="noopener noreferrer" class="wikipedia-link" title="${plainName}をWikipediaで検索">Ｗ</a>`;

            // Instagram リンク（アカウントがある場合のみ表示）
            if (meta.instagram) {
                const igUrl = `https://www.instagram.com/${meta.instagram}`;
                html += `<a href="${igUrl}" target="_blank" rel="noopener noreferrer" class="instagram-link" title="${plainName}のInstagram">ig</a>`;
            }

            // TikTok リンク（アカウントがある場合のみ表示）
            if (meta.tiktok) {
                const ttUrl = `https://www.tiktok.com/@${meta.tiktok}`;
                html += `<a href="${ttUrl}" target="_blank" rel="noopener noreferrer" class="tiktok-link" title="${plainName}のTikTok">tt</a>`;
            }

            return html;
        }\
'''


def update_wrestler_name_function(html_path):
    """
    index.html の wrestlerNameWithTwitter 関数を
    wrestlerMeta 対応版に置き換える
    """
    with open(html_path, encoding="utf-8") as f:
        content = f.read()

    pattern = r'// 選手名にX\(Twitter\)検索リンクアイコンを付与するヘルパー[\s\S]*?return [^}]+;\n        \}'
    new_func = NEW_WRESTLER_NAME_WITH_TWITTER.replace("\\'", "'")

    if re.search(pattern, content):
        content = re.sub(pattern, new_func, content)
        print("wrestlerNameWithTwitter 関数を更新しました")
    else:
        print("WARNING: wrestlerNameWithTwitter 関数が見つかりませんでした")

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(content)


# ──────────────────────────────────────────
# メイン処理
# ──────────────────────────────────────────

def main():
    print("=== CSV to HTML コンバーター開始 ===\n")

    # ステップ1: CSV 読み込み
    print(f"[1] {INPUT_CSV} を読み込み中...")
    raw_wrestlers = load_csv(INPUT_CSV)
    print(f"    {len(raw_wrestlers)} 件読み込み完了")

    # 引退フラグをパース
    wrestlers = []
    for i, w in enumerate(raw_wrestlers, start=2):  # 2行目から（1行目はヘッダー）
        retired = parse_retired_flag(w.get("引退フラグ", "false"), i, w.get("選手名", ""))
        w["_retired"] = retired if retired is not None else False
        wrestlers.append(w)

    # ステップ2: 前回の状態を読み込み
    print(f"\n[2] {STATE_FILE} を読み込み中...")
    prev_state = load_promotion_state(STATE_FILE)
    prev_count = len(prev_state.get("promotions", {}))
    print(f"    前回登録団体数: {prev_count}")

    # ステップ3: 現在の選手数を集計し表示順を決定
    print("\n[3] 団体別現役選手数を集計中...")
    current_counts = count_active_per_promotion(wrestlers)
    for p, c in sorted(current_counts.items(), key=lambda x: -x[1]):
        if p in FIXED_TAIL_COLUMNS:
            mark = "[末尾固定]"
        elif c >= MIN_WRESTLERS_FOR_COLUMN:
            mark = "○ 独立列"
        else:
            mark = f"△ その他列へ振り分け（{c}名≤3）"
        print(f"    {p}: {c}名 {mark}")

    # determine_promotion_order の戻り値には末尾固定列（海外・その他・フリー）も含まれる
    promotions_order = determine_promotion_order(current_counts, prev_state)
    print(f"\n    表示列順 ({len(promotions_order)}列):")
    for i, p in enumerate(promotions_order):
        count = current_counts.get(p, 0)
        tail_mark = " [末尾固定]" if p in FIXED_TAIL_COLUMNS else ""
        print(f"      {i+1:2d}. {p}({count}名){tail_mark}")

    # ステップ4: マトリクス生成
    print("\n[4] 年×団体マトリクスを生成中...")
    matrix = build_matrix(wrestlers, promotions_order)
    js_rows = matrix_to_js_array(matrix, promotions_order)
    print(f"    {len(js_rows)} 年分のデータを生成")

    # その他列の実際の人数を集計してログに出力（3名以下から振り分けられた分を含む）
    sonota_total = sum(
        len(names) for year_data in matrix.values()
        for col, names in year_data.items()
        if col == SONOTA_COLUMN
    )
    print(f"    その他列の実人数: {sonota_total}名（3名以下の団体から振り分けを含む）")

    # ステップ5: wrestlerMeta 生成
    print("\n[5] 選手SNSメタ情報を生成中...")
    wrestler_meta = build_wrestler_meta(wrestlers)
    print(f"    SNS情報あり: {len(wrestler_meta)} 名")

    # ステップ6: HTML 更新
    print(f"\n[6] {HTML_FILE} を更新中...")
    update_html(HTML_FILE, js_rows, promotions_order, wrestler_meta)
    update_wrestler_name_function(HTML_FILE)

    # ステップ7: promotion_state.json 更新
    print(f"\n[7] {STATE_FILE} を更新中...")
    save_promotion_state(STATE_FILE, promotions_order, current_counts)

    print("\n=== 変換完了 ===")


if __name__ == "__main__":
    main()
