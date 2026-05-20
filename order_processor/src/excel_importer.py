"""構成一覧ExcelをSQLiteにインポートする（Windows対応）"""
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
        INSERT INTO bom (親品番, 子品番, 員数, 長さ, 長さ記号, 形状, R側, L側, 形状ラベル, 備考)
        VALUES (:親品番, :子品番, :員数, :長さ, :長さ記号, :形状, :R側, :L側, :形状ラベル, :備考)
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
    out.append({
        "親品番":     hinban,
        "子品番":     hinban,   # 単品は自己参照
        "員数":       _to_float(row[6]) or 1,
        "長さ":       _to_float(row[4]),
        "長さ記号":   _str(row[5]),
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
    out.append({
        "親品番":     str(oyahinban).strip(),
        "子品番":     str(kohinban).strip(),
        "員数":       _to_float(row[2]) or 1,
        "長さ":       _to_float(row[4]),
        "長さ記号":   _str(row[5]),
        "形状":       _str(row[8]),
        "R側":        None,
        "L側":        None,
        "形状ラベル": _str(row[7]),
        "備考":       None,
    })


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
