"""構成一覧ExcelをSQLiteにインポートする（Windows対応）"""
import re
import openpyxl
from .db import get_connection


# ── 形状表記の正規化テーブル ────────────────────────────────────────
# ジュケンのExcelで表記ゆれが発生しやすい文字を統一する
_SHAPE_NORMALIZE = [
    ("U", "∩"),    # 半角Uをアーチ記号に統一
]


def _normalize_shape(shape: str | None) -> str | None:
    """形状文字列の表記ゆれを正規化する"""
    if shape is None:
        return None
    for before, after in _SHAPE_NORMALIZE:
        shape = shape.replace(before, after)
    return shape if shape.strip() else None


def parse_excel_to_rows(excel_path: str) -> list[dict]:
    """
    構成一覧ExcelをパースしてBOM行のリストを返す（DBへの書き込みは行わない）。

    差分確認などで事前にデータを取得したいときに使う。
    """
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    if "構成一覧" not in wb.sheetnames:
        raise ValueError(f"「構成一覧」シートが見つかりません: {excel_path}")

    ws = wb["構成一覧"]
    tanpin_rows = []
    assy_rows = []
    mode = "tanpin"

    for row in ws.iter_rows(min_row=6, values_only=True):
        if row[1] == "品番" and row[2] == "員数":
            mode = "assy"
            continue
        if mode == "tanpin":
            _parse_tanpin_row(row, tanpin_rows)
        else:
            _parse_assy_row(row, assy_rows)

    return tanpin_rows + assy_rows


def import_from_excel(excel_path: str, db_path: str, replace: bool = True) -> dict:
    """構成一覧ExcelをSQLiteにインポートする。
    replace=True の場合は既存データを全て置き換える。
    戻り値: {"単品": 件数, "ASSY": 件数}
    """
    all_rows = parse_excel_to_rows(excel_path)
    tanpin_count = sum(1 for r in all_rows if r["親品番"] == r["子品番"])
    assy_count   = len(all_rows) - tanpin_count

    conn = get_connection(db_path)
    if replace:
        conn.execute("DELETE FROM bom")

    conn.executemany("""
        INSERT INTO bom (親品番, 子品番, 員数, 長さ, 長さ表示, 長さ記号, 材料名称, 形状, R側, L側, 形状ラベル, 備考)
        VALUES (:親品番, :子品番, :員数, :長さ, :長さ表示, :長さ記号, :材料名称, :形状, :R側, :L側, :形状ラベル, :備考)
    """, all_rows)
    conn.commit()
    conn.close()
    return {"単品": tanpin_count, "ASSY": assy_count}


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
        "形状":       _normalize_shape(_str(row[8])),
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
        "形状":       _normalize_shape(_str(row[8])),
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
        nagasa_raw='830',   kigo='0'  → '830-0'  （数字のみ記号はハイフン付加）
        nagasa_raw='1670',  kigo='L'  → '1670L'  （L/R/RAMはハイフンなし）
        nagasa_raw='2615',  kigo=None → '2615'
    """
    if nagasa_raw is None:
        return None
    if not kigo:
        return nagasa_raw
    # 記号が純粋な数字（'0', '1' など）の場合はハイフンを付加
    # '-1', '-2' などすでにハイフンを含む場合や L/R/RAM はそのまま結合
    if kigo.isdigit():
        return f"{nagasa_raw}-{kigo}"
    return f"{nagasa_raw}{kigo}"


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
