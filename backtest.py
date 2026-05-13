"""
トヨタ（7203）移動平均線バックテスト

ルール：
- 買い：株価が25日移動平均線を上抜けたとき
- 売り（利確）：株価が25日移動平均線を下抜けたとき
- 売り（損切）：買値からSTOP_LOSS_PCT%下落したとき
"""

import pandas as pd
import numpy as np

# --- 設定 ---
TICKER = "7203.T"       # トヨタ（東証）
PERIOD_YEARS = 5        # 検証期間（年）
MA_DAYS = 25            # 移動平均の日数

# 比較する損切りラインのリスト
STOP_LOSS_LIST = [0.03, 0.05, 0.08]  # 3%、5%、8%

# --- サンプルデータ生成 ---
print(f"サンプルデータで検証: {TICKER}（過去{PERIOD_YEARS}年相当）\n")

np.random.seed(42)
dates = pd.date_range(end="2025-12-31", periods=PERIOD_YEARS * 252, freq="B")
start_price = 2200
returns = np.random.normal(0.0003, 0.015, len(dates))
prices = start_price * (1 + returns).cumprod()

df = pd.DataFrame({"close": prices}, index=dates)
df["ma"] = df["close"].rolling(MA_DAYS).mean()
df.dropna(inplace=True)


def run_backtest(df, stop_loss_pct):
    trades = []
    position = None

    for i in range(1, len(df)):
        prev = df.iloc[i - 1]
        curr = df.iloc[i]

        if position is None:
            if prev["close"] < prev["ma"] and curr["close"] >= curr["ma"]:
                position = {
                    "entry_date": df.index[i],
                    "entry_price": curr["close"],
                    "stop_loss": curr["close"] * (1 - stop_loss_pct),
                }
        else:
            sell_reason = None
            sell_price = curr["close"]

            if curr["close"] <= position["stop_loss"]:
                sell_reason = "損切り"
            elif prev["close"] >= prev["ma"] and curr["close"] < curr["ma"]:
                sell_reason = "利確（MA下抜け）"

            if sell_reason:
                profit_pct = (sell_price - position["entry_price"]) / position["entry_price"] * 100
                trades.append({
                    "損益(%)": round(profit_pct, 2),
                    "理由": sell_reason,
                })
                position = None

    return pd.DataFrame(trades)


# --- 損切りラインごとに比較 ---
print("=" * 60)
print("  損切りライン比較")
print("=" * 60)
print(f"  {'損切り':>6}  {'取引数':>6}  {'勝率':>6}  {'平均利益':>8}  {'平均損失':>8}  {'累計損益':>8}")
print("-" * 60)

for sl in STOP_LOSS_LIST:
    result = run_backtest(df, sl)
    if result.empty:
        continue

    total = len(result)
    wins = len(result[result["損益(%)"] > 0])
    losses = total - wins
    win_rate = wins / total * 100
    avg_profit = result[result["損益(%)"] > 0]["損益(%)"].mean() if wins > 0 else 0
    avg_loss = result[result["損益(%)"] <= 0]["損益(%)"].mean() if losses > 0 else 0
    total_profit = result["損益(%)"].sum()

    print(f"  {sl*100:>5.0f}%  {total:>6}回  {win_rate:>5.1f}%  {avg_profit:>+7.2f}%  {avg_loss:>+7.2f}%  {total_profit:>+7.2f}%")

print("=" * 60)
print()
print("【読み方】")
print("  損切り  : 買値からこの%下がったら強制売却")
print("  取引数  : 5年間で何回売買したか")
print("  勝率    : 利益が出たトレードの割合")
print("  平均利益: 勝ったトレードの平均プラス幅")
print("  平均損失: 負けたトレードの平均マイナス幅")
print("  累計損益: 全トレードの損益を足し合わせた合計")
