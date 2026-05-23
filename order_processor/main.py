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


def cmd_update_master(excel_path: str, force: bool = False):
    """
    構成一覧ExcelをもとにBOMマスタを更新する。

    通常は差分を確認しながら更新します。
    --force オプションをつけると確認なしで全件置き換えます（初回セットアップ時向け）。
    """
    from src.excel_importer import parse_excel_to_rows, import_from_excel
    from src.bom_updater import compute_diff, apply_diff_interactively

    initialize_db(DB_PATH)
    print(f"構成一覧Excelを読み込み中: {excel_path}")
    new_rows = parse_excel_to_rows(excel_path)
    print(f"  → {len(new_rows)} 行を読み込みました")

    if force:
        # 確認なしで全件置き換え（初回セットアップ時など）
        from src.db import get_connection
        conn = get_connection(DB_PATH)
        conn.execute("DELETE FROM bom")
        conn.executemany("""
            INSERT INTO bom (親品番, 子品番, 員数, 長さ, 長さ表示, 長さ記号, 材料名称, 形状, R側, L側, 形状ラベル, 備考)
            VALUES (:親品番, :子品番, :員数, :長さ, :長さ表示, :長さ記号, :材料名称, :形状, :R側, :L側, :形状ラベル, :備考)
        """, new_rows)
        conn.commit()
        conn.close()
        print(f"完了: {len(new_rows)} 件のBOMデータを全件置き換えました（--force）")
        return

    # DBが空なら差分なし → そのままインポート
    from src.db import get_connection
    count = get_connection(DB_PATH).execute("SELECT COUNT(*) FROM bom").fetchone()[0]
    if count == 0:
        print("  DBにデータがありません。全件インポートします。")
        from src.bom_updater import _insert_rows
        conn = get_connection(DB_PATH)
        _insert_rows(conn, new_rows)
        conn.commit()
        conn.close()
        print(f"完了: {len(new_rows)} 件をインポートしました")
        return

    # 差分確認フロー
    print(f"  現在のDB: {count} 行")
    print("  差分を確認します...\n")
    diff = compute_diff(new_rows, DB_PATH)
    apply_diff_interactively(diff, DB_PATH)


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


# ══════════════════════════════════════════════════════════════
#  UI 層 ― テキストメニュー版
#
#  将来 GUI に切り替える場合はこのブロックだけ差し替えればOK。
#  cmd_process() / cmd_update_master() などの処理層は変えない。
# ══════════════════════════════════════════════════════════════

_VERSION = "1.0"
_LINE    = "═" * 52


def _ask_file(prompt: str) -> str | None:
    """
    ファイルパスの入力を求める。
    ・Windowsのドラッグ&ドロップで付く引用符を自動で除去
    ・空欄またはキャンセル（q）でメニューに戻る
    """
    print(f"\n  {prompt}")
    print("  （ファイルをこのウィンドウにドラッグ＆ドロップしてもOK）")
    print("  （キャンセルする場合は q を入力）")
    print()
    try:
        raw = input("  > ").strip()
    except EOFError:
        return None

    if raw.lower() in ("q", "quit", "exit", ""):
        return None

    # Windows ドラッグ＆ドロップで付く引用符を除去
    path = raw.strip('"').strip("'")
    if not os.path.isfile(path):
        print(f"\n  ⚠ ファイルが見つかりません: {path}")
        print("  パスを確認してもう一度試してください。")
        return None
    return path


def _menu_process_order():
    """メニュー操作: 注文書を処理する"""
    path = _ask_file("注文書 Excel のパスを入力してください")
    if path is None:
        return
    print()
    try:
        cmd_process(path)
    except Exception as e:
        print(f"\n  ⚠ エラーが発生しました: {e}")
    _pause()


def _menu_update_master():
    """メニュー操作: 構成一覧を更新する"""
    path = _ask_file("構成一覧 Excel のパスを入力してください")
    if path is None:
        return
    print()
    try:
        cmd_update_master(path)
    except Exception as e:
        print(f"\n  ⚠ エラーが発生しました: {e}")
    _pause()


def _pause():
    """処理完了後に一時停止する"""
    print()
    try:
        input("  ── Enterキーを押すとメニューに戻ります ──")
    except EOFError:
        pass


def main_menu():
    """
    テキストメニュー版のメイン画面。

    GUI に切り替える場合はこの関数を差し替えることを想定している。
    処理の実体（cmd_process 等）はそのまま再利用できる。
    """
    initialize_db(DB_PATH)

    while True:
        # 画面クリア（Windows: cls / Mac・Linux: clear）
        os.system("cls" if os.name == "nt" else "clear")

        print(_LINE)
        print(f"  注文処理システム  ver.{_VERSION}")
        print(_LINE)
        print()
        print("  [1]  注文書を処理する")
        print("         → レポート・材料仕分けリストを生成します")
        print()
        print("  [2]  構成一覧を更新する")
        print("         → ジュケンから届いたExcelと差分確認しながら更新します")
        print()
        print("  [3]  終了")
        print()
        print(_LINE)

        try:
            choice = input("  選択してください [1-3]: ").strip()
        except EOFError:
            break

        if choice == "1":
            _menu_process_order()
        elif choice == "2":
            _menu_update_master()
        elif choice == "3":
            print("\n  終了します。\n")
            break
        else:
            print("\n  1 ～ 3 の数字を入力してください。")
            _pause()


# ══════════════════════════════════════════════════════════════
#  エントリーポイント
# ══════════════════════════════════════════════════════════════

def main():
    """
    引数なしで起動 → メニュー画面を表示。
    引数あり（--input など）→ 従来のコマンドライン動作。
    """
    # 引数が渡されている場合はコマンドライン動作（バッチ処理などに使う）
    if len(sys.argv) > 1:
        parser = argparse.ArgumentParser(description="注文解析システム")
        parser.add_argument("--setup-access", metavar="ACCDB", help="AccessDBからBOMをインポート")
        parser.add_argument("--update-master", metavar="EXCEL", help="構成一覧ExcelからBOMを更新（差分確認あり）")
        parser.add_argument("--force",         action="store_true", help="--update-master と組み合わせ: 確認なしで全件置き換え")
        parser.add_argument("--input",         metavar="EXCEL", help="注文Excelを処理してレポート生成")
        parser.add_argument("--send-email",    action="store_true", help="レポートをメール送信する")
        args = parser.parse_args()

        if args.setup_access:
            cmd_setup_access(args.setup_access)
        elif args.update_master:
            cmd_update_master(args.update_master, force=args.force)
        elif args.input:
            cmd_process(args.input, send_email=args.send_email)
        else:
            parser.print_help()
    else:
        # 引数なし → メニュー画面
        main_menu()


if __name__ == "__main__":
    main()
