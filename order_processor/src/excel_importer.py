"""構成一覧ExcelをSQLiteにインポートする（Windows対応）"""
import re
import openpyxl
from .db import get_connection


def import_from_excel(excel_path: str, db_path: str, replace: bool = True) -> dict:
    """構成一覧ExcelをSQLiteにインポートする。
    replace=True の場合は既存データを全て置き換える。
    戻り値: {"単品": 件数, "ASSY": 件数}
    """
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    if "構成一覧" not in wb.sheetnames:
        raise ValueError(f"「構成一覧」シートが見つかりません: {excel_path}")

    ws = wb["構成一覧"]
    tanpin_rows = []
    assy_rows = []
    mode = "tanpin"  # "tanpin" → "assy" に切り替わる

    for row in ws.iter_rows(min_row=6, values_only=True):
        # ASSYセクションのヘッダー行を検出（B列='品番', C列='員数'）
        if row[1] == "品番" and row[2] == "員数":
            mode = "assy"
            continue

        if mode == "tanpin":
            _parse_tanpin_row(row, tanpin_rows)
        else:
            _parse_assy_row(row, assy_rows)

    all_rows = tanpin_rows + assy_rows
    conn = get_connection(db_path)
    if replace:
        conn.execute("DELETE FROM bom")

    conn.executemany("""
        INSERT INTO bom (親品番, 子品番, 員数, 長さ, 長さ表示, 長さ記号, 材料名称, 形状, R側, L側, 形状ラベル, 備考)
        VALUES (:親品番, :子品番, :員数, :長さ, :長さ表示, :長さ記号, :材料名称, :形状, :R側, :L側, :形状ラベル, :備考)
    """, all_rows)
    conn.commit()
    conn.close()
    return {"単品": len(tanpin_rows), "ASSY": len(assy_rows)}


def _parse_tanpin_row(row, out: list):
    """単品セクション（〜行146）: D列=品番（親品番=子品番）"""
    # D=品番, E=長さ, F=長さ記号, G=バーリング数, H=ピッチ, I=形状
    hinban = row[3]
    if not hinban:
        return
    hinban = str(hinban).strip()
    # 「鏡面品番」など漢字・ひらがな・カタカナを含むラベル行を除外
    if not _is_valid_hinban(hinban):
        return
    nagasa_raw = _str(row[4])
    kigo       = _str(row[5])
    out.append({
        "親品番":     hinban,
        "子品番":     hinban,   # 単品は自己参照
        "員数":       1,
        "長さ":       _parse_nagasa_value(nagasa_raw),
        "長さ表示":   nagasa_raw,
        "長さ記号":   kigo,
        "材料名称":   _build_material_name(nagasa_raw, kigo),
        "形状":       _str(row[8]),
        "R側":        None,
        "L側":        None,
        "形状ラベル": _str(row[7]),
        "備考":       None,
    })


def _parse_assy_row(row, out: list):
    """ASSYセクション（行148〜）: B列=親品番, C列=員数, D列=子品番"""
    # B=親品番, C=員数, D=子品番, E=長さ, F=長さ記号, G=バーリング数, H=ピッチ, I=形状
    oyahinban = row[1]
    kohinban  = row[3]
    if not oyahinban or not kohinban:
        return
    nagasa_raw = _str(row[4])
    kigo       = _str(row[5])
    out.append({
        "親品番":     str(oyahinban).strip(),
        "子品番":     str(kohinban).strip(),
        "員数":       _to_float(row[2]) or 1,
        "長さ":       _parse_nagasa_value(nagasa_raw),
        "長さ表示":   nagasa_raw,
        "長さ記号":   kigo,
        "材料名称":   _build_material_name(nagasa_raw, kigo),
        "形状":       _str(row[8]),
        "R側":        None,
        "L側":        None,
        "形状ラベル": _str(row[7]),
        "備考":       None,
    })


# ①②③ などの丸付き数字（Unicode）を検出するパターン
_PREFIX_PATTERN = re.compile(r'^[①②③④⑤⑥⑦⑧⑨⑩]+')


def _parse_nagasa_value(raw: str | None) -> float | None:
    """長さ表示から数値を抽出する。
    例: '①525' → 525.0 / '2615' → 2615.0 / None → None
    """
    if raw is None:
        return None
    # 丸付き数字プレフィックスを除去して数値化
    stripped = _PREFIX_PATTERN.sub("", raw).strip()
    try:
        return float(stripped)
    except (ValueError, TypeError):
        return None


def _build_material_name(nagasa_raw: str | None, kigo: str | None) -> str | None:
    """材料名称を組み立てる。
    例: nagasa_raw='①525', kigo='-1' → '①525-1'
        nagasa_raw='2615',  kigo=None → '2615'
        nagasa_raw='155',   kigo='-1' → '155-1'
    """
    if nagasa_raw is None:
        return None
    if kigo:
        return f"{nagasa_raw}{kigo}"
    return nagasa_raw


def _is_valid_hinban(hinban: str) -> bool:
    """品番として有効かどうか判定する。
    漢字・ひらがな・カタカナを含む場合はラベル行と判定して除外。
    例: '鏡面品番' → False / '0F035600741' → True
    """
    for ch in hinban:
        cp = ord(ch)
        # CJK統合漢字・ひらがな・カタカナ範囲
        if (0x3040 <= cp <= 0x309F or   # ひらがな
            0x30A0 <= cp <= 0x30FF or   # カタカナ
            0x4E00 <= cp <= 0x9FFF):    # 漢字
            return False
    return True


def _to_float(value) -> float | None:
    if value is None:
        return None
    try:
        return float(value)
    except (ValueError, TypeError):
        return None


def _str(value) -> str | None:
    if value is None:
        return None
    s = str(value).strip()
    return s if s else None
