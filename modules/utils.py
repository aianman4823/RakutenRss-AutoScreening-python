import time
import pyautogui as py
import pandas as pd
import slackweb
import os


def load_csv(f: str):
    return pd.read_csv(f)


def excel_start_marketspeed():
    excel_maxsize()
    py.click(824, 70)
    time.sleep(1)
    py.click(40, 120)
    time.sleep(10)


def excel_move():
    py.hotkey("win", "m")


def excel_maxsize():
    py.hotkey("win", "up")


def check_buy_stay_sell(df):
    cols = df.columns
    cols = [col for col in cols if col != 'code' and col != 'trt']
    df['sum'] = df[cols].sum(axis=1)
    if all(df['sum'] >= 5) and all(df['trt'] == 1):
        df['check'] = 5
    elif all(df['sum'] >= 5) and all(df['trt'] == 0):
        df['check'] = 4
    elif all(df['sum'] == 4) and all(df['trt'] == 1):
        df['check'] = 3
    elif all(df['sum'] == 4) and all(df['trt'] == 0):
        df['check'] = 2
    elif all(df['sum'] == 3) and all(df['trt'] == 1):
        df['check'] = 1
    elif all(df['sum'] == -2):
        df['check'] = -1
    elif all(df['sum'] < -2):
        df['check'] = -2
    else:
        df['check'] = 0
    return df


def slack_notice(text):
    slack = slackweb.Slack(url=os.environ['SLACK'])
    slack.notify(text=text)
    return


if __name__ == '__main__':
    slack_notice(text="sample")
