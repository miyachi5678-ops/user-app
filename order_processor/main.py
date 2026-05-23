"""
注文解析システム - メインスクリプト

使い方:
    # 初回セットアップ（BOMをAccessからインポート）
    python main.py --setup-access path/to/database.accdb

    # 構成一覧ExcelからBOMをインポート（更新時）
    python main.py --update-master path/to/構成一覧.xlsx

    # 注文Excelを処理してレポート生成
    python main.py --input path/to/EDI注文.xlsx

    # 注文Excelを処理してレポート生成 + メール送信
    python main.py --input path/to/EDI注文.xlsx --send-email
"""
import argparse
import os
import sys
from datetime import datetime

# プロジェクトルートをパスに追加
sys.path.insert(0, os.path.dirname(__file__))

from config import EMAIL_CONFIG, DB_PATH, OUTPUT_DIR
from src.db import initialize_db, get_connection
from src.order_parser import parse_order_excel
from src.order_processor import process_orders
from src.excel_writer import write_report
from src.email_sender import send_report


# ── BOM未登録品番の確認フロー ──────────────────────────────────────

def _find_similar_in_bom(hinban: str, db_path: str) -> list[dict]:
    """
    先頭1桁を0〜9に変えた品番がBOMに存在するか検索する。

    Returns:
        [{"品番": "3D264400130", "形状例": "左トメ直片∩"}, ...]
    """
    if not hinban:
        return []
    conn = get_connection(db_path)
    rest = hinban[1:]
    similar = []
    for digit in "0123456789":
        candidate = digit + rest
        if candidate == hinban:
            continue
        rows = conn.execute(
            "SELECT 形状 FROM bom WHERE 親品番 = ? LIMIT 3",
            (candidate,)
        ).fetchall()
        if rows:
            shapes = [r["形状"] for r in rows if r["形状"]]
            similar.append({
                "品番": candidate,
                "形状例": "、".join(shapes[:2]) if shapes else "（形状なし）",
            })
    conn.close()
    return similar


def _resolve_unknowns_interactively(
    unknowns: list[tuple[str, float]],
    db_path: str,
    skip_list_path: str,
) -> None:
    """
    BOM未登録品番を1件ずつ表示し、スキップリスト追加かBOM登録かを選択させる。
    スキップを選択した品番は bom_skip_list.txt に自動で追記する。
    """
    print()
    print("━" * 62)
    print(f"  BOM未登録品番が {len(unknowns)} 件あります。1件ずつ確認してください。")
    print("━" * 62)

    for i, (hinban, qty) in enumerate(unknowns, 1):
        print()
        print(f"  [{i}/{len(unknowns)}]  品番: {hinban}  （発注数: {qty:g}）")
        print()

        # 先頭1桁違いの類似品番を検索
        similar = _find_similar_in_bom(hinban, db_path)
        if similar:
            print("  BOMに先頭1桁が異なる類似品番が見つかりました:")
            for s in similar:
                print(f"    → {s['品番']}  （形状例: {s['形状例']}）")
        else:
            print("  類似品番（先頭1桁違い）はBOMに見つかりませんでした。")

        print()
        print("  どうしますか？")
        print("  [1] スキップリストに追加（対象外品番として以降はスルー）")
        print("  [2] BOM登録が必要（構成一覧Excelに追加してください）")
        print()

        while True:
            try:
                choice = input("  選択 [1/2]: ").strip()
            except EOFError:
                # 非対話環境（パイプ等）では選択できないのでスキップ
                choice = ""
            if choice in ("1", "2"):
                break
            print("  1 か 2 を入力してください。")

        if choice == "1":
            with open(skip_list_path, "a", encoding="utf-8") as f:
                f.write(f"\n{hinban}  # 確認済み（対象外品番）")
            print(f"  → {hinban} をスキップリストに追加しました。")
        else:
            print(f"  → {hinban} は構成一覧Excelへの登録が必要です。")
            print("     登録後、--update-master で BOM を更新してください。")

        print()

    print("━" * 62)
    print("  確認完了。レポートを生成します。")
    print("━" * 62)
    print()


# ── コマンド処理 ──────────────────────────────────────────────────

def cmd_setup_access(accdb_path: str):
    """AccessデータベースからBOMをインポートする"""
    from src.access_importer import import_from_access

    print(f"Accessからインポート中: {accdb_path}")
    initialize_db(DB_PATH)
    count = import_from_access(accdb_path, DB_PATH)
    print(f"完了: {count}件のBOMデータをインポートしました")


def cmd_update_master(excel_path: str):
    """構成一覧ExcelからBOMを更新する"""
    from src.excel_importer import import_from_excel

    print(f"構成一覧Excelを読み込み中: {excel_path}")
    initialize_db(DB_PATH)
    count = import_from_excel(excel_path, DB_PATH, replace=True)
    print(f"完了: {count}件のBOMデータを更新しました")


def cmd_process(input_path: str, send_email: bool = False):
    """注文Excelを解析してレポートを生成する"""
    from src.reports.material_sorting import (
        write_material_sorting_list,
        load_skip_list,
        find_unknowns,
    )

    initialize_db(DB_PATH)

    # 注文を読み込む
    print(f"注文ファイルを読み込み中: {input_path}")
    orders = parse_order_excel(input_path)
    print(f"  → {len(orders)}件の注文を読み込みました")

    # ── BOM未登録品番の確認（レポート生成前）─────────────────────
    skip_list_path = os.path.join(os.path.dirname(__file__), "bom_skip_list.txt")
    skip_set = load_skip_list(skip_list_path)

    unknowns = find_unknowns(orders, DB_PATH, skip_set)
    if unknowns:
        _resolve_unknowns_interactively(unknowns, DB_PATH, skip_list_path)
        # 選択結果を反映するためスキップリストを再読み込み
        skip_set = load_skip_list(skip_list_path)

    # ── BOM展開（メインレポート用）───────────────────────────────
    print("BOM展開中...")
    details, unknown = process_orders(orders, DB_PATH)
    print(f"  → {len(details)}件の明細を生成しました")

    # スキップリストに載っていない未登録品番だけ表示
    unregistered = [h for h in unknown if h not in skip_set]
    if unregistered:
        print(f"  ⚠ BOM未登録品番（要対応）: {', '.join(unregistered)}")

    # ── Excel出力 ─────────────────────────────────────────────
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = os.path.basename(input_path).replace(".xlsx", "")

    # メインレポート
    output_path = os.path.join(OUTPUT_DIR, f"解析結果_{filename}_{timestamp}.xlsx")
    write_report(orders, details, unknown, output_path)
    print(f"レポートを保存しました: {output_path}")

    # 材料仕分けリスト
    month_label = datetime.now().strftime("%Y年%m月")
    sorting_path = os.path.join(OUTPUT_DIR, f"材料仕分けリスト_{filename}_{timestamp}.xlsx")
    counts, _ = write_material_sorting_list(
        orders, DB_PATH, sorting_path, month_label, skip_set
    )
    print(f"材料仕分けリストを保存しました: {sorting_path}")
    print(
        f"  ホンダ:{counts['ホンダ']} 相地:{counts['相地']}"
        f" 直:{counts['直']} 直鏡面:{counts['直 鏡面']}"
    )

    # メール送信
    if send_email:
        print("メール送信中...")
        recipients = send_report(output_path, EMAIL_CONFIG, len(orders), len(unknown))
        print(f"  → 送信完了: {', '.join(recipients)}")

    return output_path


def main():
    parser = argparse.ArgumentParser(description="注文解析システム")
    parser.add_argument("--setup-access", metavar="ACCDB", help="AccessDBからBOMをインポート")
    parser.add_argument("--update-master", metavar="EXCEL", help="構成一覧ExcelからBOMを更新")
    parser.add_argument("--input",         metavar="EXCEL", help="注文Excelを処理してレポート生成")
    parser.add_argument("--send-email",    action="store_true", help="レポートをメール送信する")
    args = parser.parse_args()

    if args.setup_access:
        cmd_setup_access(args.setup_access)
    elif args.update_master:
        cmd_update_master(args.update_master)
    elif args.input:
        cmd_process(args.input, send_email=args.send_email)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
