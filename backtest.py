"""
トヨタ（7203）移動平均線バックテスト

ルール：
- 買い：株価が25日移動平均線を上抜けたとき
- 売り（利確）：株価が25日移動平均線を下抜けたとき
- 売り（損切）：買値から5%下落したとき
"""

import yfinance as yf
import pandas as pd

# --- 設定 ---
TICKER = "7203.T"       # トヨタ（東証）
PERIOD_YEARS = 5        # 検証期間（年）
MA_DAYS = 25            # 移動平均の日数
STOP_LOSS_PCT = 0.05    # 損切りライン（5%）

# --- データ取得（またはサンプルデータ生成）---
import numpy as np

print(f"サンプルデータで検証: {TICKER}（過去{PERIOD_YEARS}年相当）")

# トヨタ実際の株価に近いランダムウォークデータを生成（再現性のため seed 固定）
np.random.seed(42)
dates = pd.date_range(end="2025-12-31", periods=PERIOD_YEARS * 252, freq="B")
start_price = 2200  # 2020年頃のトヨタ株価（円）
returns = np.random.normal(0.0003, 0.015, len(dates))  # 日次リターン（平均+0.03%、標準偏差1.5%）
prices = start_price * (1 + returns).cumprod()

df = pd.DataFrame({"close": prices}, index=dates)

# --- 移動平均線の計算 ---
df["ma"] = df["close"].rolling(MA_DAYS).mean()
df.dropna(inplace=True)

# --- バックテスト ---
trades = []
position = None  # 保有中のポジション情報

for i in range(1, len(df)):
    prev = df.iloc[i - 1]
    curr = df.iloc[i]

    # 買いシグナル：前日は平均線以下、当日は平均線以上（上抜け）
    if position is None:
        if prev["close"] < prev["ma"] and curr["close"] >= curr["ma"]:
            position = {
                "entry_date": df.index[i],
                "entry_price": curr["close"],
                "stop_loss": curr["close"] * (1 - STOP_LOSS_PCT),
            }

    # 売りシグナル
    elif position is not None:
        sell_reason = None
        sell_price = curr["close"]

        # 損切り
        if curr["close"] <= position["stop_loss"]:
            sell_reason = "損切り"

        # 利確（移動平均線を下抜け）
        elif prev["close"] >= prev["ma"] and curr["close"] < curr["ma"]:
            sell_reason = "利確（MA下抜け）"

        if sell_reason:
            profit_pct = (sell_price - position["entry_price"]) / position["entry_price"] * 100
            trades.append({
                "買い日": position["entry_date"].strftime("%Y-%m-%d"),
                "売り日": df.index[i].strftime("%Y-%m-%d"),
                "買い値": round(float(position["entry_price"]), 0),
                "売り値": round(float(sell_price), 0),
                "損益(%)": round(profit_pct, 2),
                "理由": sell_reason,
            })
            position = None

# --- 結果集計 ---
print("\n" + "=" * 55)
print(f"  バックテスト結果：トヨタ（{TICKER}）過去{PERIOD_YEARS}年")
print("=" * 55)

if not trades:
    print("該当するトレードがありませんでした。")
else:
    result_df = pd.DataFrame(trades)

    total = len(result_df)
    wins = len(result_df[result_df["損益(%)"] > 0])
    losses = total - wins
    win_rate = wins / total * 100
    avg_profit = result_df[result_df["損益(%)"] > 0]["損益(%)"].mean() if wins > 0 else 0
    avg_loss = result_df[result_df["損益(%)"] <= 0]["損益(%)"].mean() if losses > 0 else 0
    total_profit = result_df["損益(%)"].sum()

    print(f"  総トレード数　: {total} 回")
    print(f"  勝ちトレード　: {wins} 回")
    print(f"  負けトレード　: {losses} 回")
    print(f"  勝率　　　　　: {win_rate:.1f}%")
    print(f"  平均利益　　　: +{avg_profit:.2f}%")
    print(f"  平均損失　　　: {avg_loss:.2f}%")
    print(f"  累計損益　　　: {total_profit:.2f}%")
    print("=" * 55)

    # 理由別集計
    print("\n【売り理由の内訳】")
    reason_counts = result_df["理由"].value_counts()
    for reason, count in reason_counts.items():
        avg = result_df[result_df["理由"] == reason]["損益(%)"].mean()
        print(f"  {reason}: {count}回 （平均損益 {avg:+.2f}%）")

    # 直近10件のトレード
    print("\n【直近10件のトレード】")
    print(result_df.tail(10).to_string(index=False))
