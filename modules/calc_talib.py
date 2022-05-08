# import time
import os
from glob import glob
import talib as ta
import pandas as pd
import mplfinance as mpf
import matplotlib.pyplot as plt
from pathlib import Path
from backtesting.lib import crossover
import numpy as np


def check_macd(df):
    """
    dfから最後の行を抜き出して利用
    """
    if crossover(df['macd'], df['macd_signal']):
        return 1
    elif crossover(df['macd_signal'], df['macd']):
        return -1
    else:
        return 0


def check_rci(df):
    """
    dfから最後の行を抜き出して利用
    """
    rci = df['rci']
    buy_entry = (rci < -80)
    sell_entry = (rci > 80)
    if buy_entry.values[-1]:
        return 1
    elif sell_entry.values[-1]:
        return -1
    else:
        return 0


def check_rsi(df):
    """
    dfから最後の行を抜き出して利用
    """
    rsi = df['rsi']
    close = df['Close']
    df['ema'] = close.ewm(span=20).mean()

    h, l, c_prev = df["High"], df["Low"], pd.Series(df['Close']).shift(1)
    tr = np.max([h - l, (c_prev - h).abs(), (c_prev - l).abs()], axis=0)
    atr = pd.Series(tr).rolling(20).mean().bfill().values
    df['atr'] = atr
    lower = df["ema"] - df["atr"]
    upper = df["ema"] + df["atr"]

    buy_entry = (rsi < 20) & (close < lower)
    sell_entry = (rsi > 80) & (close > upper)
    if buy_entry.values[-1]:
        return 1
    elif sell_entry.values[-1]:
        return -1
    else:
        return 0


def check_stoch_slow(df):
    """
    dfから最後の行を抜き出して利用
    """
    lower_20 = (df['slowK'] <= 20) & (df['slowD'] <= 20)
    higher_80 = (df['slowK'] >= 80) & (df['slowD'] >= 80)
    if crossover(df['slowK'], df['slowD']) and lower_20.values[-1]:
        return 1
    elif crossover(df['slowD'], df['slowK']) and higher_80.values[-1]:
        return -1
    else:
        return 0


def check_dmi(df):
    """
    dfから最後の行を抜き出して利用
    """
    if crossover(df['plus_di'], df['minus_di']) and df['adx'].pct_change(9).values[-1] > 1:
        return 1
    elif crossover(df['minus_di'], df['plus_di']) and df['adx'].pct_change(9).values[-1] > 1:
        return -1
    else:
        return 0


def check_three_rise_trend(df):
    cols = df.columns
    cols = [col for col in cols if col != "Close"]
    check = True
    for col in cols:
        if not all(df[col] > 0):
            check = False
            break
    if check:
        return 1
    else:
        return 0


def calc_macd(df, long_span=26, short_span=12, signal_span=9):
    # macd, macdsignal, _ = ta.MACD(close, fastperiod=12, slowperiod=26, signalperiod=9)
    df['span_long'] = df['Close'].ewm(span=long_span).mean()
    df['span_short'] = df['Close'].ewm(span=short_span).mean()
    df['macd'] = df['span_short'] - df['span_long']
    df['macd_signal'] = df['macd'].ewm(span=signal_span).mean()
    df = df.drop(['span_short', 'span_long'], axis=1)
    return df


def calc_rci(df, timeperiod=9):
    close = df['Close']
    rci = []
    for j in range(len(close)):
        if j < timeperiod:
            rci.append(np.nan)
        else:
            data = pd.DataFrame()
            data['close'] = list(close[j - timeperiod:j])
            data = data.reset_index()
            data = data.rename(columns={'index': 'original_index'})  # 最初のindex情報を残しておきます
            data = data.sort_values('close', ascending=False).reset_index(drop=True)  # closeの値が高い順に並べ替えます
            data = data.reset_index()  # indexをresetします
            data['index'] = [i + 1 for i in data['index']]  # resetしたindexは0〜8なので、1〜9の順位に変換するために1を足します
            data = data.rename(columns={'index': 'price_rank'})  # closeの高い順に並べたindexをprice rankという列名にします
            data = data.set_index('original_index')  # 元のindexに戻します
            data = data.sort_index()  # 元のindexに戻します
            data['date_rank'] = np.arange(timeperiod, 0, -1)  # 元のindexに日付順位をつけます
            data['delta'] = [(data.loc[ii, 'price_rank'] - data.loc[ii, 'date_rank'])**2 for ii in range(len(data))]  # dの値を計算します
            d = data['delta'].sum()  # dの値を計算します
            value = (1 - (6 * d) / (timeperiod**3 - timeperiod)) * 100  # rciを計算します
            rci.append(value)
    df['rci'] = rci
    return df


def calc_rsi(df,
             timeperiod=14):
    close = df['Close']
    rsi = ta.RSI(close, timeperiod=timeperiod)
    df['rsi'] = rsi
    return df


def calc_stoch_slow(df,
                    slowk_period=3,
                    slowd_period=3):
    high = df['High']
    low = df['Low']
    close = df['Close']
    tastoch = ta.STOCH(high, low, close,
                       slowk_period, slowd_period)
    df['slowK'] = tastoch[0]
    df['slowD'] = tastoch[1]
    return df


def calc_dmi(df, plus_timeperiod=14, minus_timeperiod=14, adx_timeperiod=9):
    # DMI
    high = df['High']
    low = df['Low']
    close = df['Close']
    plus_di = ta.PLUS_DI(high, low, close, timeperiod=plus_timeperiod)
    minus_di = ta.MINUS_DI(high, low, close, timeperiod=minus_timeperiod)
    adx = ta.ADX(high, low, close, timeperiod=adx_timeperiod)
    df['plus_di'] = plus_di
    df['minus_di'] = minus_di
    df['adx'] = adx
    return df


def calc_three_rise_trend(df):
    # DMI
    close = df['Close']
    df['diff'] = df['Close'] - df['Open']
    diff = df['diff'].tail(4)
    diff = diff.reset_index(drop=True)
    close_tail = close.tail(4)
    close_tail = close_tail.reset_index(drop=True)
    close_tail_last = close_tail.tail(1)
    close_tail_slice = close_tail[:3]
    close_tail_con = pd.concat([close_tail_last, close_tail_slice], axis=0)
    close_tail_con = close_tail_con.reset_index(drop=True)
    df_cl_df = pd.concat([diff, close_tail_con], axis=1)
    df_cl_df['ratio'] = df_cl_df['diff'] / df_cl_df['Close']
    df_cl_df = df_cl_df.tail(3)
    return df_cl_df


if __name__ == '__main__':
    data_dir_csvs = os.path.join(os.environ['CSV_DIR'], 'all_20220428')
    if not Path(data_dir_csvs).exists():
        raise NotImplementedError(data_dir_csvs)
        exit()

    datas = glob(os.path.join(data_dir_csvs, '*csv'))
    for i, f in enumerate(datas):
        df = pd.read_csv(f)
        df = df.dropna(axis=0)
        if df.empty:
            continue

        df['Date'] = pd.to_datetime(df['Date'])
        df = df.set_index('Date', drop=True)

        # 終値からMACDを計算(MACD)
        df = calc_macd(df)
        mdf = df
        apd = [
            mpf.make_addplot(mdf['macd'], panel=2, color='red'),  # パネルの2番地に赤で描画
            mpf.make_addplot(mdf['macd_signal'], panel=2, color='blue'),
        ]

        mpf.plot(mdf, type='candle', volume=True, addplot=apd)
        plt.close()

        # 終値からRSIを計算
        df = calc_rsi(df)
        mdf = df
        apd = [
            mpf.make_addplot(mdf['rsi'], panel=2, color='red'),
        ]

        mpf.plot(mdf, type='candle', volume=True, addplot=apd)
        plt.close()

        # stochastic(買われすぎ、売られすぎを判断)
        df = calc_stoch_slow(df)
        mdf = df
        apd = [
            mpf.make_addplot(mdf['slowK'], panel=2, color='red'),
            mpf.make_addplot(mdf['slowD'], panel=2, color='blue'),
        ]

        mpf.plot(mdf, type='candle', volume=True, addplot=apd)
        plt.close()

        # DMI
        df = calc_dmi(df)
        mdf = df
        apd = [
            mpf.make_addplot(mdf['plus_di'], panel=2, color='red'),
            mpf.make_addplot(mdf['minus_di'], panel=2, color='blue'),
            mpf.make_addplot(mdf['adx'], panel=2, color='green'),
        ]

        mpf.plot(mdf, type='candle', volume=True, addplot=apd)
        plt.close()

        df = calc_rci(df)
        mdf = df
        apd = [
            mpf.make_addplot(mdf["rci"],panel=2,color="red"),
        ]
        mpf.plot(mdf, type="candle",volume=True,addplot=apd)
