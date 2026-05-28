"""BOM差分比較・対話的更新モジュール

ジュケンから新しい構成一覧Excelが届いたとき、
現在のマスタDB との差分を確認しながら更新する。

誰が触っても迷わないよう、すべての操作を日本語で案内します。
"""
from .db import get_connection


# 比較対象とするフィールド（これらのいずれかが違えば「変更あり」と判定）
_COMPARE_FIELDS = ["子品番", "員数", "長さ", "形状", "材料名称"]

# DBへの挿入に使うフィールド
_INSERT_FIELDS = [
    "親品番", "子品番", "員数", "長さ", "長さ表示",
    "長さ記号", "材料名称", "形状", "R側", "L側", "形状ラベル", "備考",
]


def compute_diff(new_rows: list[dict], db_path: str) -> dict:
    """
    新しい構成一覧Excelのデータと現在のDBを比較し、差分を返す。

    Args:
        new_rows: excel_importer.parse_excel_to_rows() が返した行リスト
        db_path:  SQLiteデータベースパス

    Returns:
        {
            "added":     [(親品番, rows), ...],            # DBにない新しい品番
            "deleted":   [(親品番, rows), ...],            # 新Excelにない既存品番
            "changed":   [(親品番, db_rows, new_rows), ...], # 内容が変わった品番
            "unchanged": [親品番, ...],                    # 変更なし
        }
    """
    conn = get_connection(db_path)
    db_all = conn.execute("SELECT * FROM bom").fetchall()
    conn.close()

    # DB側を親品番でグループ化
    db_by_oya: dict[str, list[dict]] = {}
    for r in db_all:
        key = r["親品番"]
        db_by_oya.setdefault(key, []).append(dict(r))

    # 新Excel側を親品番でグループ化
    new_by_oya: dict[str, list[dict]] = {}
    for r in new_rows:
        key = r["親品番"]
        new_by_oya.setdefault(key, []).append(r)

    db_keys  = set(db_by_oya.keys())
    new_keys = set(new_by_oya.keys())

    added   = [(k, new_by_oya[k]) for k in sorted(new_keys - db_keys)]
    deleted = [(k, db_by_oya[k])  for k in sorted(db_keys - new_keys)]

    changed   = []
    unchanged = []
    for k in sorted(db_keys & new_keys):
        if _rows_differ(db_by_oya[k], new_by_oya[k]):
            changed.append((k, db_by_oya[k], new_by_oya[k]))
        else:
            unchanged.append(k)

    return {
        "added":     added,
        "deleted":   deleted,
        "changed":   changed,
        "unchanged": unchanged,
    }


def _rows_differ(db_rows: list[dict], new_rows: list[dict]) -> bool:
    """2つのBOM行セットに実質的な差異があるか判定する"""
    def sig(r):
        return tuple(str(r.get(f) or "") for f in _COMPARE_FIELDS)

    return sorted(sig(r) for r in db_rows) != sorted(sig(r) for r in new_rows)


def apply_diff_interactively(diff: dict, db_path: str, ask_func=None) -> dict:
    """
    差分を画面に表示し、1件ずつ確認しながらDBを更新する。

    ask_func(label1, label2, context) -> '1' | '2'
        指定しない場合はターミナル入力を使う。
        GUIモードではtkinterダイアログを渡す。
    """
    if ask_func is None:
        ask_func = _cli_ask
    added     = diff["added"]
    deleted   = diff["deleted"]
    changed   = diff["changed"]
    unchanged = diff["unchanged"]

    # ── サマリー表示 ──────────────────────────────────────────
    print()
    print("━" * 62)
    print("  構成一覧  差分サマリー")
    print("━" * 62)
    print(f"  変更なし:          {len(unchanged):4d} 品番")
    print(f"  新規追加:          {len(added):4d} 品番  ← 自動でDBに追加します")
    print(f"  削除予定:          {len(deleted):4d} 品番  ← 1件ずつ確認します")
    print(f"  内容が変わった:    {len(changed):4d} 品番  ← 1件ずつ確認します")
    print("━" * 62)

    if not added and not deleted and not changed:
        print("\n  差分はありません。DBはすでに最新の状態です。\n")
        return {"added": 0, "deleted": 0, "changed": 0}

    conn = get_connection(db_path)

    # ── 新規追加（自動） ──────────────────────────────────────
    if added:
        print(f"\n【新規追加】 {len(added)} 件をDBに追加します\n")
        for oya, rows in added:
            print(f"  + {oya}  （{len(rows)} 行）")
            _insert_rows(conn, rows)

    # ── 削除予定（1件ずつ確認） ───────────────────────────────
    n_deleted = 0
    if deleted:
        print(f"\n【削除予定】 {len(deleted)} 件  ─  新しい構成一覧にない品番です\n")
        print("  ※ 手動で追加した品番や、ジュケンが削除した品番が含まれます。")
        print("     1件ずつ「残す」か「削除する」かを選んでください。\n")

        for i, (oya, rows) in enumerate(deleted, 1):
            print(f"  ─── [{i} / {len(deleted)}]  品番: {oya} ───")
            print()
            _print_rows_summary(rows)
            print()
            print("  [1] 残す  （削除しない）")
            print("  [2] 削除する")
            choice = ask_func("残す（削除しない）", "削除する",
                              f"品番: {oya}\n新しい構成一覧にこの品番がありません。")

            if choice == "2":
                conn.execute("DELETE FROM bom WHERE 親品番 = ?", (oya,))
                n_deleted += 1
                print(f"  → {oya} を削除しました。")
            else:
                print(f"  → {oya} を残します。")
            print()

    # ── 内容変更（1件ずつ確認） ───────────────────────────────
    n_changed = 0
    if changed:
        print(f"\n【内容変更】 {len(changed)} 件  ─  データが変わっている品番です\n")
        print("  ※ 1件ずつ現在の内容と新しい内容を見比べて選んでください。\n")

        for i, (oya, db_rows, new_rows_list) in enumerate(changed, 1):
            print(f"  ─── [{i} / {len(changed)}]  品番: {oya} ───")
            print()
            print("  ■ 現在のDB:")
            _print_rows_summary(db_rows)
            print()
            print("  ■ 新しい構成一覧:")
            _print_rows_summary(new_rows_list)
            print()
            print("  [1] 新しい内容で更新する")
            print("  [2] 現在の内容をそのまま残す（変更しない）")
            choice = ask_func("新しい内容で更新する", "現在の内容をそのまま残す",
                              f"品番: {oya}")

            if choice == "1":
                conn.execute("DELETE FROM bom WHERE 親品番 = ?", (oya,))
                _insert_rows(conn, new_rows_list)
                n_changed += 1
                print(f"  → {oya} を更新しました。")
            else:
                print(f"  → {oya} の現在の内容を残します。")
            print()

    conn.commit()
    conn.close()

    # ── 完了サマリー ─────────────────────────────────────────
    print("━" * 62)
    print("  更新完了")
    print(f"  追加: {len(added)} 件  /  削除: {n_deleted} 件  /  "
          f"更新: {n_changed} 件  /  変更なし: {len(unchanged)} 件")
    print("━" * 62)
    print()

    return {"added": len(added), "deleted": n_deleted, "changed": n_changed}


# ── ヘルパー ──────────────────────────────────────────────────────

def _print_rows_summary(rows: list[dict]):
    """BOM行を見やすく表示する"""
    for r in rows:
        parts = []
        if r.get("形状"):
            parts.append(f"形状: {r['形状']}")
        if r.get("材料名称"):
            parts.append(f"材料: {r['材料名称']}")
        if r.get("員数") and float(r.get("員数") or 1) != 1.0:
            parts.append(f"員数: {r['員数']}")
        print("    " + "  /  ".join(parts) if parts else "    （データなし）")


def _insert_rows(conn, rows: list[dict]):
    """BOM行をDBに挿入する"""
    placeholders = ", ".join(f":{f}" for f in _INSERT_FIELDS)
    cols = ", ".join(_INSERT_FIELDS)
    conn.executemany(
        f"INSERT INTO bom ({cols}) VALUES ({placeholders})",
        rows,
    )


def _cli_ask(label1: str, label2: str, context: str = "") -> str:
    """ターミナル版の選択入力（CLIモード用）"""
    while True:
        try:
            choice = input("  選択 [1/2]: ").strip()
        except EOFError:
            return "1"
        if choice in ("1", "2"):
            return choice
        print("  1 か 2 を入力してください。")
