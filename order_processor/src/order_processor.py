"""注文をBOMで展開し、必要部材を算出する"""
from dataclasses import dataclass
from .db import get_connection
from .order_parser import Order


@dataclass
class OrderDetail:
    オーダーNo: str
    親品番: str
    子品番: str | None
    数量: float
    最新納期: str | None
    工場: str | None
    形状: str | None
    R側: str | None
    直_長さ: float | None
    L側: str | None
    形状ラベル: str | None


def process_orders(orders: list[Order], db_path: str) -> tuple[list[OrderDetail], list[str]]:
    """
    注文リストをBOMで展開する。
    戻り値: (展開済み明細リスト, BOMに見つからなかった品番リスト)
    """
    conn = get_connection(db_path)
    details = []
    unknown_hinban = []

    for order in orders:
        bom_rows = conn.execute(
            "SELECT * FROM bom WHERE 親品番 = ?",
            (order.品番,)
        ).fetchall()

        if not bom_rows:
            if order.品番 not in unknown_hinban:
                unknown_hinban.append(order.品番)
            # BOMが見つからなくても明細には記録する
            details.append(OrderDetail(
                オーダーNo=order.オーダーNo,
                親品番=order.品番,
                子品番=None,
                数量=order.発注数,
                最新納期=order.最新納期,
                工場=order.工場,
                形状=None,
                R側=None,
                直_長さ=None,
                L側=None,
                形状ラベル=f"【構成未登録】{order.品番}",
            ))
            continue

        for bom in bom_rows:
            数量 = order.発注数 * (bom["員数"] or 1)
            details.append(OrderDetail(
                オーダーNo=order.オーダーNo,
                親品番=order.品番,
                子品番=bom["子品番"],
                数量=数量,
                最新納期=order.最新納期,
                工場=order.工場,
                形状=bom["形状"],
                R側=bom["R側"],
                直_長さ=bom["長さ"],
                L側=bom["L側"],
                形状ラベル=bom["形状ラベル"],
            ))

    conn.close()
    return details, unknown_hinban


def summarize_by_length(details: list[OrderDetail]) -> list[dict]:
    """子品番・長さ別に数量を集計する"""
    summary: dict[tuple, dict] = {}

    for d in details:
        key = (d.子品番, d.直_長さ, d.形状)
        if key not in summary:
            summary[key] = {
                "子品番": d.子品番,
                "直_長さ": d.直_長さ,
                "形状": d.形状,
                "形状ラベル": d.形状ラベル,
                "合計数量": 0,
                "納期一覧": set(),
            }
        summary[key]["合計数量"] += d.数量
        if d.最新納期:
            summary[key]["納期一覧"].add(d.最新納期)

    result = []
    for item in summary.values():
        item["納期一覧"] = "、".join(sorted(item["納期一覧"]))
        result.append(item)

    return sorted(result, key=lambda x: (x["子品番"] or "", x["直_長さ"] or 0))
