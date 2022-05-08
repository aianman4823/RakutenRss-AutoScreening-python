import os
import time
import xlwings as xw
import pandas as pd
from pathlib import Path
from dotenv import load_dotenv
import string
from glob import glob

from modules import (
    excel_start_marketspeed
)

load_dotenv()


def make_dataset(open_file_name: str,
                 chart_type: str,
                 start_date: str,
                 row_num: int,
                 output_dir: str,
                 today: str):

    if "positions" in open_file_name:
        all_or_pos_dir = os.path.join(os.environ['EXCEL_DIR'], f"positions_{today}")
    else:
        all_or_pos_dir = os.path.join(os.environ['EXCEL_DIR'], "all")
    excel_path = os.path.join(all_or_pos_dir, open_file_name)
    if not Path(excel_path).exists():
        raise NotImplementedError(f"{open_file_name} was not found")
    wb = xw.Book(excel_path)
    wb.app.activate(steal_focus=True)
    addin_path = os.environ['ADDIN_PATH']
    wb.app.api.RegisterXLL(addin_path)

    excel_start_marketspeed()

    sheets_name = wb.sheets
    check = False
    for v in sheets_name:
        if v.name == "dataset":
            check = True
            break
    if check:
        sh = wb.sheets["dataset"]
    else:
        raise NotImplementedError("dataset sheets not found")
        return
    sh.activate()

    #   セルの値　取得・設定
    cols_name = sh.range("2:2").value
    cols_name = [v for v in cols_name if v != "" or v is not None]
    # バク発見: 銘柄コードのあとに半角スペース1つ必要
    if "銘柄コード " not in cols_name:
        if "銘柄コード" in cols_name:
            ind = cols_name.index("銘柄コード")
            cols_name[ind] = "銘柄コード "
        else:
            raise NotImplementedError("銘柄コード  not found")
            return

    code = cols_name.index("銘柄コード ")
    alf = list(string.ascii_uppercase)
    positions = sh.range(f'{alf[code]}:{alf[code]}').value
    positions = [int(v) for v in positions if v is not None and type(v) != str]

    while len(positions) > 0:
        v = positions.pop()
        wb.sheets.add(name=str(v))
        sh = wb.sheets[str(v)]
        sh.activate()
        time.sleep(1)
        sh.range("B1").value = v
        sh.range("C1").value = chart_type
        sh.range("D1").value = start_date
        sh.range("E1").value = row_num
        sh.range('A2:G2').value = ["銘柄名称", "日付", "始値", "高値", "安値", "終値", "出来高"]
        sh.range('A1').value = "=@RssChartPast(A2:G2, B1, C1, D1, E1)"
        time.sleep(1.5)
        alfcols = ['B', 'C', 'D', 'E', 'F', 'G']
        cols = ['Date', 'Open', 'High', 'Low', 'Close', 'Volume']
        df = pd.DataFrame()
        for i, p in enumerate(alfcols):
            data = sh.range(f'{p}3:{p}3100').value
            data = [v for v in data if v is not None and v != '--------']
            df[cols[i]] = data

        if not Path(output_dir).exists():
            os.makedirs(output_dir)

        df.to_csv(os.path.join(output_dir, f'{str(v)}.csv'), index=0)
        wb.sheets[str(v)].delete()

    # このファイルは市場が閉じたあとに実行
    app = xw.apps.active
    wb.close()
    app.kill()


if __name__ == '__main__':
    # make_dataset(open_file_name="positions.xlsm",
    #              chart_type='D',
    #              start_date='20140101',
    #              row_num=3000,
    #              output_dir="../csvs")

    excel_dir_path = os.environ['EXCEL_DIR']
    files = glob(os.path.join(excel_dir_path, 'all_brand_*'))
    for f in files:
        f = os.path.basename(f)
        make_dataset(open_file_name=f,
                     chart_type='D',
                     start_date='20140101',
                     row_num=3000,
                     output_dir=os.path.join("..", "csvs"))
