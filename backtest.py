"""
日本株 移動平均線バックテスト（複数銘柄比較）

ルール：
- 買い：株価が25日移動平均線を上抜けたとき
- 売り（利確）：株価が25日移動平均線を下抜けたとき
- 売り（損切）：買値から5%下落したとき
"""

import pandas as pd
import numpy as np

# --- 設定 ---
PERIOD_YEARS = 5    # 検証期間（年）
MA_DAYS = 25        # 移動平均の日数
STOP_LOSS_PCT = 0.05  # 損切りライン（5%）

# 銘柄リスト（コード、名前、参考株価）
STOCKS = [
    ("7011.T", "三菱重工", 3500),
    ("8058.T", "三菱商事", 3000),
    ("6758.T", "ソニー",   10000),
    ("6954.T", "ファナック", 25000),
    ("6501.T", "日立",     6000),
    ("6701.T", "NEC",      5000),
]


def make_sample_data(seed, start_price, period_years):
    np.random.seed(seed)
    dates = pd.date_range(end="2025-12-31", periods=period_years * 252, freq="B")
    returns = np.random.normal(0.0003, 0.015, len(dates))
    prices = start_price * (1 + returns).cumprod()
    df = pd.DataFrame({"close": prices}, index=dates)
    df["ma"] = df["close"].rolling(MA_DAYS).mean()
    return df.dropna()


def run_backtest(df, stop_loss_pct):
    trades = []
    position = None

    for i in range(1, len(df)):
        prev = df.iloc[i - 1]
        curr = df.iloc[i]

        if position is None:
            if prev["close"] < prev["ma"] and curr["close"] >= curr["ma"]:
                position = {
                    "entry_price": curr["close"],
                    "stop_loss": curr["close"] * (1 - stop_loss_pct),
                }
        else:
            sell_reason = None
            sell_price = curr["close"]

            if curr["close"] <= position["stop_loss"]:
                sell_reason = "損切り"
            elif prev["close"] >= prev["ma"] and curr["close"] < curr["ma"]:
                sell_reason = "利確"

            if sell_reason:
                profit_pct = (sell_price - position["entry_price"]) / position["entry_price"] * 100
                trades.append({"損益(%)": round(profit_pct, 2), "理由": sell_reason})
                position = None

    return pd.DataFrame(trades)


# --- 銘柄別比較 ---
print(f"検証条件：25日移動平均線ルール、損切り{STOP_LOSS_PCT*100:.0f}%、期間{PERIOD_YEARS}年\n")
print("=" * 72)
print(f"  {'銘柄':<10}  {'取引数':>5}  {'勝率':>6}  {'平均利益':>8}  {'平均損失':>8}  {'累計損益':>8}  {'損切り回数':>8}")
print("-" * 72)

results = []
for seed, (ticker, name, start_price) in enumerate(STOCKS):
    df = make_sample_data(seed, start_price, PERIOD_YEARS)
    result = run_backtest(df, STOP_LOSS_PCT)

    if result.empty:
        continue

    total = len(result)
    wins = len(result[result["損益(%)"] > 0])
    losses = total - wins
    win_rate = wins / total * 100
    avg_profit = result[result["損益(%)"] > 0]["損益(%)"].mean() if wins > 0 else 0
    avg_loss = result[result["損益(%)"] <= 0]["損益(%)"].mean() if losses > 0 else 0
    total_profit = result["損益(%)"].sum()
    stop_count = len(result[result["理由"] == "損切り"])

    results.append({
        "銘柄": name,
        "取引数": total,
        "勝率": win_rate,
        "平均利益": avg_profit,
        "平均損失": avg_loss,
        "累計損益": total_profit,
        "損切り回数": stop_count,
    })

    print(f"  {name:<10}  {total:>5}回  {win_rate:>5.1f}%  {avg_profit:>+7.2f}%  {avg_loss:>+7.2f}%  {total_profit:>+7.2f}%  {stop_count:>6}回")

print("=" * 72)

# ランキング（累計損益順）
sorted_results = sorted(results, key=lambda x: x["累計損益"], reverse=True)
print("\n【累計損益ランキング】")
for i, r in enumerate(sorted_results, 1):
    print(f"  {i}位  {r['銘柄']:<10}  {r['累計損益']:>+7.2f}%")

print()
print("【読み方】")
print("  勝率    : 利益が出たトレードの割合")
print("  平均利益: 勝ったときの平均プラス幅")
print("  平均損失: 負けたときの平均マイナス幅")
print("  累計損益: 5年間の全トレードを足した合計")
print("  損切り回数: 損切りルールが発動した回数")
