"""SQLiteデータベースの初期化・接続管理"""
import sqlite3
import os


def get_connection(db_path: str) -> sqlite3.Connection:
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def initialize_db(db_path: str):
    """テーブルを作成する（存在しない場合のみ）"""
    conn = get_connection(db_path)
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS bom (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            親品番      TEXT NOT NULL,
            子品番      TEXT,
            員数        REAL DEFAULT 1,
            長さ        REAL,       -- 数値（計算用）例: 525.0
            長さ表示    TEXT,       -- 表示用プレフィックス付き 例: ①525
            長さ記号    TEXT,       -- F列の記号 例: -1, -2, L, R
            材料名称    TEXT,       -- 発注時の材料名 例: ①525-1
            形状        TEXT,
            R側         TEXT,
            L側         TEXT,
            形状ラベル  TEXT,
            備考        TEXT,
            更新日時    TEXT DEFAULT (datetime('now', 'localtime'))
        );

        CREATE INDEX IF NOT EXISTS idx_bom_親品番 ON bom(親品番);

        CREATE TABLE IF NOT EXISTS orders (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            取込日時    TEXT DEFAULT (datetime('now', 'localtime')),
            ファイル名  TEXT,
            オーダーNo  TEXT NOT NULL,
            品番        TEXT NOT NULL,
            発注数      REAL NOT NULL,
            納期        TEXT,
            最新納期    TEXT,
            工場        TEXT,
            発注日      TEXT
        );

        CREATE TABLE IF NOT EXISTS order_details (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            order_id    INTEGER REFERENCES orders(id) ON DELETE CASCADE,
            オーダーNo  TEXT NOT NULL,
            親品番      TEXT NOT NULL,
            子品番      TEXT,
            数量        REAL,
            最新納期    TEXT,
            工場        TEXT,
            形状        TEXT,
            R側         TEXT,
            直_長さ     REAL,
            L側         TEXT,
            形状ラベル  TEXT
        );

        CREATE INDEX IF NOT EXISTS idx_details_order ON order_details(オーダーNo);
    """)
    conn.commit()
    conn.close()
