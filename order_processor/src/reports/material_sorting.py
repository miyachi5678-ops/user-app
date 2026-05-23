"""材料仕分けリスト生成モジュール

オーエフから材料が届いた際に仕分けするためのリストを生成する。
ホンダ / 相地 / 直 / 直鏡面 の4グループに分類し、
材料名称と合計数量を集計して横並びExcelを出力する。
"""
import re
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from ..db import get_connection
from ..order_parser import Order


# ── グループ定義 ──────────────────────────────────────────────
GROUP_HONDA       = "ホンダ"
GROUP_AICHI       = "相地"
GROUP_CHOKU       = "直"
GROUP_CHOKU_KAGAMI = "直 鏡面"

# グループの表示順
GROUP_ORDER = [GROUP_HONDA, GROUP_AICHI, GROUP_CHOKU, GROUP_CHOKU_KAGAMI]

# 列幅設定
COL_WIDTHS = {
    "担当": 8,
    "記号": 12,
    "数量": 6,
}


def build_sorting_data(orders: list[Order], db_path: str) -> dict[str, dict[str, float]]:
    """
    注文リストをBOM展開し、材料仕分けデータを集計する。

    Returns:
        {
            "ホンダ":  {"材料名称": 合計数量, ...},
            "相地":    {"材料名称": 合計数量, ...},
            "直":      {"材料名称": 合計数量, ...},
            "直 鏡面": {"材料名称": 合計数量, ...},
        }
    """
    conn = get_connection(db_path)
    result: dict[str, dict[str, float]] = {g: {} for g in GROUP_ORDER}

    for order in orders:
        bom_rows = conn.execute(
            "SELECT * FROM bom WHERE 親品番 = ?",
            (order.品番,)
        ).fetchall()

        if not bom_rows:
            continue

        # この親品番のどこかにトメがあるか確認
        parent_has_tome = any("トメ" in (r["形状"] or "") for r in bom_rows)

        for row in bom_rows:
            形状     = row["形状"] or ""
            材料名称 = row["材料名称"]
            数量     = order.発注数 * (row["員数"] or 1)

            if not 材料名称:
                continue

            # グループ判定
            group = _classify_group(形状, parent_has_tome)

            result[group][材料名称] = result[group].get(材料名称, 0.0) + 数量

    conn.close()
    return result


def _classify_group(形状: str, parent_has_tome: bool) -> str:
    """形状文字列とトメ有無からグループを判定する。

    「直」系と判定するルール:
      - 形状から「鏡面」「図XXXX（図面参照番号）」「外蓋/内蓋」を正規化・除去し
      - 残った文字列が「直」と「蓋」だけで構成される場合 → 直 or 直鏡面
      - 例: '直', '蓋直', '直蓋', '蓋直蓋', '蓋直図2401.5', '直蓋　鏡面' など
    """
    has_鏡面 = "鏡面" in 形状

    # 「鏡面」「図XXXX / 図はXXXX（図面参照番号）」「外蓋/内蓋」を除去し空白も除去
    shape_clean = 形状.replace("鏡面", "")
    shape_clean = re.sub(r'図は?[\d.]+', '', shape_clean)   # 図2401.5 / 図は390.7
    shape_clean = shape_clean.replace("外蓋", "蓋").replace("内蓋", "蓋")
    shape_clean = shape_clean.replace("　", "").replace(" ", "").strip()

    # 「直」と「蓋」だけで構成され、かつ「直」を1つ以上含む → 直系（ストレートカット）
    if re.fullmatch(r'[直蓋]*直[直蓋]*', shape_clean):
        return GROUP_CHOKU_KAGAMI if has_鏡面 else GROUP_CHOKU
    elif parent_has_tome:
        return GROUP_AICHI
    else:
        return GROUP_HONDA


def _sort_key(name: str) -> float:
    """材料名称を数値ソート用キーに変換する（例: '①525-1' → 525.0）"""
    # 丸付き数字プレフィックスを除去
    s = re.sub(r'^[①②③④⑤⑥⑦⑧⑨⑩]+', '', str(name)).strip()
    # 先頭の数値部分を取り出す
    m = re.match(r'^(\d+\.?\d*)', s)
    return float(m.group(1)) if m else float('inf')


def write_material_sorting_list(
    orders: list[Order],
    db_path: str,
    output_path: str,
    month_label: str,
) -> dict[str, int]:
    """
    材料仕分けリストExcelを出力する。

    Returns:
        {"ホンダ": 件数, "相地": 件数, "直": 件数, "直 鏡面": 件数}
    """
    data = build_sorting_data(orders, db_path)

    # 各グループを数値順ソート
    sorted_groups: dict[str, list[tuple[str, float]]] = {}
    for group in GROUP_ORDER:
        sorted_groups[group] = sorted(
            data[group].items(),
            key=lambda x: _sort_key(x[0])
        )

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = month_label

    _write_sheet(ws, sorted_groups)
    _set_column_widths(ws)

    wb.save(output_path)

    return {g: len(sorted_groups[g]) for g in GROUP_ORDER}


def _write_sheet(ws, sorted_groups: dict[str, list[tuple[str, float]]]):
    """シートにデータを書き込む"""
    # ── スタイル定義 ──────────────────────────────────
    header_fills = {
        GROUP_HONDA:        PatternFill("solid", fgColor="DDEBF7"),  # 青系
        GROUP_AICHI:        PatternFill("solid", fgColor="E2EFDA"),  # 緑系
        GROUP_CHOKU:        PatternFill("solid", fgColor="FFF2CC"),  # 黄系
        GROUP_CHOKU_KAGAMI: PatternFill("solid", fgColor="FCE4D6"),  # オレンジ系
    }
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    bold = Font(bold=True)
    center = Alignment(horizontal="center")

    # 各グループのデータ行数
    max_rows = max((len(rows) for rows in sorted_groups.values()), default=0)

    # ── グループごとに列位置を計算（担当/記号/数量 + 空白列） ──
    # 列番号: ホンダ=1,2,3 | 相地=5,6,7 | 直=9,10,11 | 直鏡面=13,14,15
    group_cols = {
        GROUP_HONDA:        (1, 2, 3),
        GROUP_AICHI:        (5, 6, 7),
        GROUP_CHOKU:        (9, 10, 11),
        GROUP_CHOKU_KAGAMI: (13, 14, 15),
    }

    # ── ヘッダー行 ──────────────────────────────────────
    for group, (c_担当, c_記号, c_数量) in group_cols.items():
        fill = header_fills[group]
        for col, label in [(c_担当, "担当"), (c_記号, "記号"), (c_数量, "数量")]:
            cell = ws.cell(row=1, column=col, value=label)
            cell.font = bold
            cell.fill = fill
            cell.border = border
            cell.alignment = center

    # ── データ行 ──────────────────────────────────────
    for group, (c_担当, c_記号, c_数量) in group_cols.items():
        rows = sorted_groups[group]
        for i, (材料名称, 数量) in enumerate(rows):
            row_num = i + 2  # 1行目=ヘッダーなので2行目から

            # 担当列: 1行目だけグループ名を書く（2行目以降は省略 or 繰り返し）
            cell_担当 = ws.cell(row=row_num, column=c_担当, value=group)
            cell_担当.alignment = center

            cell_記号 = ws.cell(row=row_num, column=c_記号, value=材料名称)
            cell_数量 = ws.cell(row=row_num, column=c_数量, value=int(数量) if 数量 == int(数量) else 数量)
            cell_数量.alignment = Alignment(horizontal="right")

            for cell in [cell_担当, cell_記号, cell_数量]:
                cell.border = border


def _set_column_widths(ws):
    """列幅を設定する"""
    # 担当列: 1, 5, 9, 13
    for col in [1, 5, 9, 13]:
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 9
    # 記号列: 2, 6, 10, 14
    for col in [2, 6, 10, 14]:
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 13
    # 数量列: 3, 7, 11, 15
    for col in [3, 7, 11, 15]:
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 6
    # 空白列: 4, 8, 12
    for col in [4, 8, 12]:
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 2
