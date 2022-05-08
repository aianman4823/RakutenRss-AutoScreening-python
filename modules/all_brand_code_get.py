import os
import time
import numpy as np
import xlwings as xw
from pathlib import Path
from dotenv import load_dotenv
from glob import glob

from modules import (
    excel_start_marketspeed
)

load_dotenv()


def start_gather_all_code(output_dir: str):
    if not Path(output_dir).exists():
        os.makedirs(output_dir)

    file_path = glob(os.path.join(output_dir, 'all_brand_*.xlsx'))
    if len(file_path) != 0:
        return

    num = 0
    # positions = [str(v).zfill(4) for v in range(1000, 10000)]
    positions = []
    start = 1000
    gap = 500
    while start < 10000:
        if start + gap > 10000:
            pos = np.arange(start, 10000, 1)
        else:
            pos = np.arange(start, start + gap, 1)
        start += gap
        positions.append(pos)

    while len(positions) > 0:
        position = positions.pop()
        wb = xw.Book()
        wb.app.activate()
        time.sleep(1)
        addin_path = os.environ['ADDIN_PATH']
        wb.app.api.RegisterXLL(addin_path)

        excel_start_marketspeed()

        sh = wb.sheets.add("dataset")
        sh.activate()

        other_names = [v for v in wb.sheets if v.name != "dataset"]
        for v in other_names:
            wb.sheets[v.name].delete()

        sh.range("A2:B2").value = ["銘柄コード", "銘柄名称", "市場部名称"]
        row_values = []
        for i, p in enumerate(position):
            row_values.append([str(p), f"=@RssMarket(A{i + 3}, $B$2)", f"=@RssMarket(A{i + 3}, $C$2)"])
        sh.range("A3").value = row_values
        time.sleep(5)

        rownum = sh.range('A1').current_region.last_cell.row
        Sheet_Max = rownum + 1
        for i in reversed(range(3, Sheet_Max)):
            if sh.range(f"B{i}").value == "" or sh.range(f"C{i}").value == "" or sh.range(f"B{i}").value is None or sh.range(f"C{i}").value is None:
                sh.range(f'{i}:{i}').delete()
                # time.sleep(0.1)
        num += 1
        time.sleep(2)

        app = xw.apps.active
        wb.save(os.path.join(output_dir, f'all_brand_{num}.xlsx'))
        wb.close()
        app.kill()
        time.sleep(1)


if __name__ == '__main__':
    start_gather_all_code(output_dir='../excels')
