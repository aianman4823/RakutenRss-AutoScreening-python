import os
import time
import xlwings as xw
from dotenv import load_dotenv
import string
from pathlib import Path

from modules import (
    excel_start_marketspeed
)

load_dotenv()


def check_positions(output_file_name: str,
                    output_dir: str):
    wb = xw.Book()
    wb.app.activate()
    addin_path = os.environ['ADDIN_PATH']
    wb.app.api.RegisterXLL(addin_path)

    excel_start_marketspeed()

    sh = wb.sheets.add("dataset")
    sh.activate()

    other_names = [v for v in wb.sheets if v.name != "dataset"]
    for v in other_names:
        wb.sheets[v.name].delete()

    cols_name = ["銘柄コード ",
                 "銘柄名称",
                 "口座区分",
                 "保有数量",
                 "発注数量",
                 "平均取得価額",
                 "時価",
                 "前日比",
                 "前日比率",
                 "時価評価額",
                 "評価損益額",
                 "評価損益率",
                 "銘柄情報等"
                 "PER",
                 "PBR",
                 "配当利回り"]
    alf = list(string.ascii_uppercase)

    sh.range(f"{alf[0]}2:{alf[len(cols_name)- 1]}2").value = cols_name
    sh.range("A1").value = f"=@RssPositionList({alf[0]}2:{alf[len(cols_name)- 1]}2)"

    # このファイルは市場が閉じたあとに実行
    time.sleep(3)
    app = xw.apps.active
    if not Path(output_dir).exists():
        os.makedirs(output_dir)
    wb.save(os.path.join(output_dir, output_file_name))
    wb.close()
    app.kill()


if __name__ == '__main__':
    check_positions(output_file_name="positions.xlsm",
                    output_dir="../excels")
