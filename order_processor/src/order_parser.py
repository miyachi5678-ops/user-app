"""③ジュケンシートから注文データを読み込む"""
import openpyxl
from dataclasses import dataclass
from datetime import datetime, date


@dataclass
class Order:
    オーダーNo: str
    品番: str
    発注数: float
    納期: str | None
    最新納期: str | None
    工場: str | None
    発注日: str | None


def parse_order_excel(excel_path: str) -> list[Order]:
    """③ジュケンシートから注文リストを返す"""
    wb = openpyxl.load_workbook(excel_path, data_only=True)

    sheet_name = None
    for name in wb.sheetnames:
        if "ジュケン" in name or "juken" in name.lower():
            sheet_name = name
            break
    if sheet_name is None:
        raise ValueError(f"ジュケンシートが見つかりません。シート一覧: {wb.sheetnames}")

    ws = wb[sheet_name]
    orders = []

    # 行2がヘッダー、行3からデータ
    for row in ws.iter_rows(min_row=3, values_only=True):
        # B列=納期, C列=最新納期, D列=品番, E列=オーダーNo, F列=発注数, H列=工場, I列=発注日
        _, noki, saishinnoki, hinban, order_no, hatchusuu, _, kojyo, hatchubi, *_ = (
            list(row) + [None] * 15
        )[:15]

        if not hinban or not order_no:
            continue
        if str(order_no).strip() in ("計画", ""):
            continue

        orders.append(Order(
            オーダーNo=str(order_no).strip(),
            品番=str(hinban).strip(),
            発注数=float(hatchusuu) if hatchusuu else 1.0,
            納期=_format_date(noki),
            最新納期=_format_date(saishinnoki),
            工場=str(kojyo).strip() if kojyo else None,
            発注日=_format_date(hatchubi),
        ))

    return orders


def _format_date(value) -> str | None:
    if value is None:
        return None
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y/%m/%d")
    s = str(value).strip()
    # シリアル値（数値）が来た場合はスキップ
    try:
        float(s)
        return None
    except ValueError:
        return s if s else None
