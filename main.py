import subprocess
import time
import os
import pandas as pd
from pathlib import Path
import pyautogui as py
import sys
from glob import glob
import datetime
from tqdm import tqdm

from modules import (
    start_gather_all_code,
    check_positions,
    excel_move,
    make_dataset,
    calc_macd,
    calc_rsi,
    calc_rci,
    calc_dmi,
    calc_stoch_slow,
    calc_three_rise_trend,
    load_csv,
    check_macd,
    check_rsi,
    check_rci,
    check_stoch_slow,
    check_dmi,
    check_three_rise_trend,
    check_buy_stay_sell,
    slack_notice
)

market_speed = None


def main():
    global market_speed
    excel_move()
    time.sleep(1)
    start_rss()

    today_datetime = datetime.date.today()
    today = today_datetime.strftime('%Y%m%d')

    td = datetime.timedelta(weeks=10)
    start_date = today_datetime - td
    start_date_str = start_date.strftime("%Y%m%d")

    t_delta = datetime.timedelta(hours=9)
    JST = datetime.timezone(t_delta, 'JST')
    now = datetime.datetime.now(JST)
    d = now.strftime('%Y/%m/%d %H:%M')
    slack_notice(text=f"Auto screening is starting at {d} !")

    # TODO: ここからコマンドラインの引数にしたがって処理をわける
    # 保有銘柄の状況を作成
    debug = False
    if not debug:
        check_positions(output_file_name="positions.xlsx",
                        output_dir=os.path.join(os.environ['EXCEL_DIR'], f'positions_{today}'))

        # # 分析用データ作成
        make_dataset("positions.xlsx",
                     chart_type="D",
                     start_date=start_date_str,
                     row_num=50,
                     today=today,
                     output_dir=os.path.join(os.environ['CSV_DIR'], f"positions_{today}"))

        # 分析
        summary_df = None
        pos_csv_files = glob(os.path.join(os.path.join(os.environ['CSV_DIR'], f'positions_{today}'), '*'))
        result_pos_path = os.path.join(os.path.join(os.environ['OUTPUT_DIR'], f'positions_{today}'))
        if not Path(result_pos_path).exists():
            os.makedirs(result_pos_path)
        for f in pos_csv_files:
            df = load_csv(f)
            if df.empty:
                continue
            df = df.dropna(how='any', axis=0)
            if df.shape[0] < 30:
                continue

            df_b1 = df.head(df.shape[0] - 1)
            df_b2 = df.head(df.shape[0] - 2)
            df_b3 = df.head(df.shape[0] - 3)
            df_b4 = df.head(df.shape[0] - 4)

            # macd
            df = calc_macd(df)
            df_b1 = calc_macd(df_b1)
            df_b2 = calc_macd(df_b2)
            df_b3 = calc_macd(df_b3)
            df_b4 = calc_macd(df_b4)
            # rsi
            df = calc_rsi(df)
            # rci
            df = calc_rci(df)
            # dmi
            df = calc_dmi(df)
            df_b1 = calc_dmi(df_b1)
            df_b2 = calc_dmi(df_b2)
            df_b3 = calc_dmi(df_b3)
            df_b4 = calc_dmi(df_b4)
            # stoch
            df = calc_stoch_slow(df)
            df_b1 = calc_stoch_slow(df_b1)
            df_b2 = calc_stoch_slow(df_b2)
            df_b3 = calc_stoch_slow(df_b3)
            df_b4 = calc_stoch_slow(df_b4)
            new_df = calc_three_rise_trend(df)
            # TODO: checkする
            ans_macd = check_macd(df)
            ans_macd_b1 = check_macd(df_b1)
            ans_macd_b2 = check_macd(df_b2)
            ans_macd_b3 = check_macd(df_b3)
            ans_macd_b4 = check_macd(df_b4)
            ans_rsi = check_rsi(df)
            ans_rci = check_rci(df)
            ans_stoch = check_stoch_slow(df)
            ans_stoch_b1 = check_stoch_slow(df_b1)
            ans_stoch_b2 = check_stoch_slow(df_b2)
            ans_stoch_b3 = check_stoch_slow(df_b3)
            ans_stoch_b4 = check_stoch_slow(df_b4)
            ans_dmi = check_dmi(df)
            ans_dmi_b1 = check_dmi(df_b1)
            ans_dmi_b2 = check_dmi(df_b2)
            ans_dmi_b3 = check_dmi(df_b3)
            ans_dmi_b4 = check_dmi(df_b4)
            ans_trt = check_three_rise_trend(new_df)
            ans_df = pd.DataFrame()
            ans_df['code'] = [os.path.basename(f)[:-4]]
            ans_df['macd'] = [ans_macd]
            ans_df['macd_b1'] = [ans_macd_b1]
            ans_df['macd_b2'] = [ans_macd_b2]
            ans_df['macd_b3'] = [ans_macd_b3]
            ans_df['macd_b4'] = [ans_macd_b4]
            ans_df['rsi'] = [ans_rsi]
            ans_df['rci'] = [ans_rci]
            ans_df['stoch'] = [ans_stoch]
            ans_df['stoch_b1'] = [ans_stoch_b1]
            ans_df['stoch_b2'] = [ans_stoch_b2]
            ans_df['stoch_b3'] = [ans_stoch_b3]
            ans_df['stoch_b4'] = [ans_stoch_b4]
            ans_df['dmi'] = [ans_dmi]
            ans_df['dmi_b1'] = [ans_dmi_b1]
            ans_df['dmi_b2'] = [ans_dmi_b2]
            ans_df['dmi_b3'] = [ans_dmi_b3]
            ans_df['dmi_b4'] = [ans_dmi_b4]
            ans_df['trt'] = [ans_trt]
            ans_df = check_buy_stay_sell(ans_df)
            ans_df.to_csv(os.path.join(result_pos_path, os.path.basename(f)), index=0)

            # slack
            if ans_df['check'][0] != 0:
                t_delta = datetime.timedelta(hours=9)
                JST = datetime.timezone(t_delta, 'JST')
                now = datetime.datetime.now(JST)
                d = now.strftime('%Y/%m/%d/%H:%M')
                slack_notice(text=f"{ans_df['code'][0]} is {ans_df['check'][0]} at {d} !")

            if summary_df is None:
                summary_df = ans_df
                continue
            summary_df = pd.concat([summary_df, ans_df], axis=0)
        summary_df.to_csv(os.path.join(os.environ['OUTPUT_DIR'], f'summary_positions_{today}.csv'), index=0)

        # すべての銘柄コードを取得
        start_gather_all_code(output_dir=os.path.join(os.environ['EXCEL_DIR'], 'all'))
    files = sorted(glob(os.path.join(os.path.join(os.environ['EXCEL_DIR'], 'all'), '*')))
    for f in files:
        make_dataset(os.path.basename(f),
                     chart_type="D",
                     start_date=start_date_str,
                     today=today,
                     row_num=50,
                     output_dir=os.path.join(os.environ['CSV_DIR'], f"all_{today}"))

    summary_df = None
    all_csv_files = glob(os.path.join(os.path.join(os.environ['CSV_DIR'], f'all_{today}'), '*'))
    result_all_path = os.path.join(os.path.join(os.environ['OUTPUT_DIR'], f'all_{today}'))
    if not Path(result_all_path).exists():
        os.makedirs(result_all_path)
    cnts = 0
    startname = None
    ans_df = None
    for f in tqdm(all_csv_files):
        df = load_csv(f)
        if df.empty:
            continue
        df = df.dropna(how='any', axis=0)
        if df.shape[0] < 30:
            continue

        df_b1 = df.head(df.shape[0] - 1)
        df_b2 = df.head(df.shape[0] - 2)
        df_b3 = df.head(df.shape[0] - 3)
        df_b4 = df.head(df.shape[0] - 4)

        # macd
        df = calc_macd(df)
        df_b1 = calc_macd(df_b1)
        df_b2 = calc_macd(df_b2)
        df_b3 = calc_macd(df_b3)
        df_b4 = calc_macd(df_b4)
        # rsi
        df = calc_rsi(df)
        # rci
        df = calc_rci(df)
        # dmi
        df = calc_dmi(df)
        df_b1 = calc_dmi(df_b1)
        df_b2 = calc_dmi(df_b2)
        df_b3 = calc_dmi(df_b3)
        df_b4 = calc_dmi(df_b4)
        # stoch
        df = calc_stoch_slow(df)
        df_b1 = calc_stoch_slow(df_b1)
        df_b2 = calc_stoch_slow(df_b2)
        df_b3 = calc_stoch_slow(df_b3)
        df_b4 = calc_stoch_slow(df_b4)
        new_df = calc_three_rise_trend(df)
        # TODO: checkする
        ans_macd = check_macd(df)
        ans_macd_b1 = check_macd(df_b1)
        ans_macd_b2 = check_macd(df_b2)
        ans_macd_b3 = check_macd(df_b3)
        ans_macd_b4 = check_macd(df_b4)
        ans_rsi = check_rsi(df)
        ans_rci = check_rci(df)
        ans_stoch = check_stoch_slow(df)
        ans_stoch_b1 = check_stoch_slow(df_b1)
        ans_stoch_b2 = check_stoch_slow(df_b2)
        ans_stoch_b3 = check_stoch_slow(df_b3)
        ans_stoch_b4 = check_stoch_slow(df_b4)
        ans_dmi = check_dmi(df)
        ans_dmi_b1 = check_dmi(df_b1)
        ans_dmi_b2 = check_dmi(df_b2)
        ans_dmi_b3 = check_dmi(df_b3)
        ans_dmi_b4 = check_dmi(df_b4)
        ans_trt = check_three_rise_trend(new_df)
        tmp_df = pd.DataFrame()
        tmp_df['code'] = [os.path.basename(f)[:-4]]
        tmp_df['macd'] = [ans_macd]
        tmp_df['macd_b1'] = [ans_macd_b1]
        tmp_df['macd_b2'] = [ans_macd_b2]
        tmp_df['macd_b3'] = [ans_macd_b3]
        tmp_df['macd_b4'] = [ans_macd_b4]
        tmp_df['rsi'] = [ans_rsi]
        tmp_df['rci'] = [ans_rci]
        tmp_df['stoch'] = [ans_stoch]
        tmp_df['stoch_b1'] = [ans_stoch_b1]
        tmp_df['stoch_b2'] = [ans_stoch_b2]
        tmp_df['stoch_b3'] = [ans_stoch_b3]
        tmp_df['stoch_b4'] = [ans_stoch_b4]
        tmp_df['dmi'] = [ans_dmi]
        tmp_df['dmi_b1'] = [ans_dmi_b1]
        tmp_df['dmi_b2'] = [ans_dmi_b2]
        tmp_df['dmi_b3'] = [ans_dmi_b3]
        tmp_df['dmi_b4'] = [ans_dmi_b4]
        tmp_df['trt'] = [ans_trt]
        tmp_df = check_buy_stay_sell(tmp_df)
        # slack
        if tmp_df['check'][0] != 0:
            t_delta = datetime.timedelta(hours=9)
            JST = datetime.timezone(t_delta, 'JST')
            now = datetime.datetime.now(JST)
            d = now.strftime('%Y/%m/%d/%H:%M')
            slack_notice(text=f"{tmp_df['code'][0]} is {tmp_df['check'][0]} at {d} !")

        cnts += 1
        if ans_df is None:
            ans_df = tmp_df.copy()
        else:
            ans_df = pd.concat([ans_df, tmp_df], axis = 0)
        if summary_df is None:
            summary_df = pd.DataFrame()

        if startname is None:
            startname = os.path.basename(f)[:-4]
        if cnts % 500 == 0:
            ans_df.to_csv(os.path.join(result_all_path, startname + '-' + os.path.basename(f)), index=0)
            summary_df = pd.concat([summary_df, ans_df], axis=0)
            ans_df = None
            startname = None
            cnts = 0

    if cnts != 0:
        ans_df.to_csv(os.path.join(result_all_path, startname + '-' + os.path.basename(f)), index=0)
        summary_df = pd.concat([summary_df, ans_df], axis=0)

    if not summary_df.empty or summary_df is not None:
        summary_df.to_csv(os.path.join(os.environ['OUTPUT_DIR'], f'summary_all_{today}.csv'), index=0)

    # slack
    t_delta = datetime.timedelta(hours=9)
    JST = datetime.timezone(t_delta, 'JST')
    now = datetime.datetime.now(JST)
    d = now.strftime('%Y/%m/%d %H:%M')
    slack_notice(text=f"Auto screening is done at {d} !")
    
    time.sleep(1)
    stop_rss()
    time.sleep(5)
    excel_move()
    sys.exit()
    return


# RSSを起動します
def start_rss():
    global market_speed
    # この下の行は重要です。binがあるところをディレクトリに設定しないと動きませんでした。
    os.chdir(os.environ['MARKET_SPEED_APP_PATH'])
    # EXEファイルを探して、下で指定してください。
    market_speed_path = os.environ['MARKET_SPEED_PATH']
    market_speed = subprocess.Popen(market_speed_path)
    # market_speed.wait()
    time.sleep(20)
    # 10秒程度だと、朝一では少ない。バージョンアップの確認が入ります。時にはもっと時間がかかるかもしれません。
    # 画面中央のクリックを入れないと不安定
    py.click(1174, 695)  # 画面中央をクリック 画面中央なのか確認をお願いします。
    time.sleep(1)
    py.typewrite(os.environ['MARKET_SPEED_PASSWORD'])
    time.sleep(2)
    py.doubleClick(1199, 705)  # 画面中央をクリック 画面中央なのか確認をお願いします。
    # py.press("Enter")
    time.sleep(10)
    # こちらも10秒では少ない


def startrss2():
    global market_speed
    os.chdir(os.environ['MARKET_SPEED_APP_PATH'])
    market_speed_path = os.environ['MARKET_SPEED_PATH']
    market_speed = subprocess.Popen(market_speed_path)
    time.sleep(10)



# RSSを停止します
def stop_rss():
    global market_speed
    market_speed.kill()

def sample():
    start_rss()
    time.sleep(5)
    time.sleep(5)
    #startrss2()
    time.sleep(5)
    stop_rss()
    stop_rss()
    return

if __name__ == '__main__':
   main()
   #sample()
