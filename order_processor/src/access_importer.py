"""AccessデータベースのT30_構成一覧マスタをSQLiteにインポートする（Linux/Mac用）
Windows環境ではexcel_importerを使用してください"""
import subprocess
import csv
import io
from .db import get_connection


def import_from_access(accdb_path: str, db_path: str) -> int:
    """AccessのT30_構成一覧マスタをSQLiteにインポートする。戻り値はインポート行数"""
    result = subprocess.run(
        ["mdb-export", accdb_path, "T30_構成一覧マスタ"],
        capture_output=True,
    )
    if result.returncode != 0:
        raise RuntimeError(f"mdb-export失敗: {result.stderr.decode()}")

    text = result.stdout.decode("utf-8")
    reader = csv.DictReader(io.StringIO(text))

    conn = get_connection(db_path)
    conn.execute("DELETE FROM bom")

    rows = []
    for row in reader:
        rows.append({
            "親品番":     row.get("親品番", "").strip('"'),
            "子品番":     row.get("子品番", "").strip('"') or None,
            "員数":       _to_float(row.get("員数")),
            "長さ":       _to_float(row.get("直（長さ）")),
            "長さ記号":   row.get("直（記号）", "").strip('"') or None,
            "形状":       row.get("形状", "").strip('"') or None,
            "R側":        row.get("R側", "").strip('"') or None,
            "L側":        row.get("L側", "").strip('"') or None,
            "形状ラベル": row.get("形状（ラベル用）", "").strip('"') or None,
            "備考":       row.get("備考２", "").strip('"') or None,
        })

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
        return float(str(value).strip('"'))
    except (ValueError, TypeError):
        return None
