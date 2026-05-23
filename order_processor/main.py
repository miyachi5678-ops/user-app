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
from src.db import initialize_db
from src.order_parser import parse_order_excel
from src.order_processor import process_orders
from src.excel_writer import write_report
from src.email_sender import send_report


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
    from src.reports.material_sorting import write_material_sorting_list, load_skip_list

    initialize_db(DB_PATH)

    # 注文を読み込む
    print(f"注文ファイルを読み込み中: {input_path}")
    orders = parse_order_excel(input_path)
    print(f"  → {len(orders)}件の注文を読み込みました")

    # BOM展開（メインレポート用）
    print("BOM展開中...")
    details, unknown = process_orders(orders, DB_PATH)
    print(f"  → {len(details)}件の明細を生成しました")
    if unknown:
        print(f"  ⚠ 構成未登録品番: {', '.join(unknown)}")

    # 出力ディレクトリ確認
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = os.path.basename(input_path).replace(".xlsx", "")

    # メインレポート出力
    output_path = os.path.join(OUTPUT_DIR, f"解析結果_{filename}_{timestamp}.xlsx")
    write_report(orders, details, unknown, output_path)
    print(f"レポートを保存しました: {output_path}")

    # ── 材料仕分けリスト出力 ──────────────────────────────────
    skip_list_path = os.path.join(os.path.dirname(__file__), "bom_skip_list.txt")
    skip_set = load_skip_list(skip_list_path)

    month_label = datetime.now().strftime("%Y年%m月")
    sorting_path = os.path.join(OUTPUT_DIR, f"材料仕分けリスト_{filename}_{timestamp}.xlsx")

    counts, sort_unmatched = write_material_sorting_list(
        orders, DB_PATH, sorting_path, month_label, skip_set
    )
    print(f"材料仕分けリストを保存しました: {sorting_path}")
    print(
        f"  ホンダ:{counts['ホンダ']} 相地:{counts['相地']}"
        f" 直:{counts['直']} 直鏡面:{counts['直 鏡面']}"
    )

    # 未ヒット品番の警告（コンソール）
    if sort_unmatched:
        print()
        print("━" * 62)
        print("⚠  【要確認】材料仕分けリスト ― BOM未登録品番")
        print("   以下の品番は構成表（BOM）にも除外リストにも見つかりませんでした。")
        print("   ジュケン等に確認し、問題なければ bom_skip_list.txt に追記してください。")
        print("   詳細は材料仕分けリストExcelの「⚠要確認」シートも参照してください。")
        print()
        for hinban, qty in sort_unmatched:
            print(f"   ・{hinban}  （発注数: {qty:g}）")
        print("━" * 62)
        print()

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
