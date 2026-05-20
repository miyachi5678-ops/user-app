"""構成一覧ExcelをSQLiteにインポートする（Windows対応）"""
import openpyxl
from datetime import datetime
from .db import get_connection


# 構成一覧シートの列マッピング（0始まりのインデックス）
COL_NOTE     = 0   # A: メモ（類似品番注意など）
COL_HINBAN   = 3   # D: 品番
COL_NAGASA   = 4   # E: 長さ
COL_KIGO     = 5   # F: 長さ（記号）: -1, -2, L, R など
COL_BRNUM    = 6   # G: BR数
COL_BRSUN    = 7   # H: BR寸法
COL_SHAPE    = 8   # I: 形状
COL_CORNER   = 9   # J: コーナー数


def import_from_excel(excel_path: str, db_path: str, replace: bool = True) -> int:
    """構成一覧ExcelをSQLiteにインポートする。
    replace=True の場合は既存データを全て置き換える。
    戻り値はインポート行数。
    """
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    if "構成一覧" not in wb.sheetnames:
        raise ValueError(f"「構成一覧」シートが見つかりません: {excel_path}")

    ws = wb["構成一覧"]
    rows = []

    for row in ws.iter_rows(min_row=6, values_only=True):
        hinban = row[COL_HINBAN]
        if not hinban:
            continue

        hinban = str(hinban).strip()
        nagasa  = _to_float(row[COL_NAGASA])
        kigo    = str(row[COL_KIGO]).strip() if row[COL_KIGO] is not None else None
        brnum   = _to_float(row[COL_BRNUM])
        shape   = str(row[COL_SHAPE]).strip() if row[COL_SHAPE] is not None else None
        brsun   = str(row[COL_BRSUN]).strip() if row[COL_BRSUN] is not None else None

        # 構成一覧Excelでは 品番 = 親品番 = 子品番（自己参照）
        rows.append({
            "親品番":     hinban,
            "子品番":     hinban,
            "員数":       brnum if brnum is not None else 1,
            "長さ":       nagasa,
            "長さ記号":   kigo,
            "形状":       shape,
            "R側":        None,
            "L側":        None,
            "形状ラベル": brsun,
            "備考":       None,
        })

    conn = get_connection(db_path)
    if replace:
        conn.execute("DELETE FROM bom")

    conn.executemany("""
        INSERT INTO bom (親品番, 子品番, 員数, 長さ, 長さ記号, 形状, R側, L側, 形状ラベル, 備考)
        VALUES (:親品番, :子品番, :員数, :長さ, :長さ記号, :形状, :R側, :L側, :形状ラベル, :備考)
    """, rows)
    conn.commit()
    conn.close()
    return len(rows)


def _to_float(value) -> float | None:
    if value is None:
        return None
    try:
        return float(value)
    except (ValueError, TypeError):
        return None
