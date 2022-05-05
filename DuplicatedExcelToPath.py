import os
import time
import shutil
import zipfile
import pandas as pd
from goods import goods

EXCELFILENAME = '租赁成交记录_' + time.strftime("%Y%m%d") + '.xls'
PATH = '去重_' + time.strftime("%Y%m%d")


def duplicateExcelAndSaveToPath():
    if not os.path.exists(PATH):
        os.mkdir(PATH)
    z = zipfile.ZipFile(f'{time.strftime("%Y%m%d")}.zip', 'w')
    for GoodsId in goods:
        pd.read_excel(EXCELFILENAME, sheet_name=goods[GoodsId]).drop_duplicates(
            subset=['租赁价格', '租赁时长', '租赁类型', '押金', '成交时间'], keep='last', inplace=False).to_excel(
            f'{PATH}/{goods[GoodsId]}_租赁成交记录_' + time.strftime("%Y%m%d") + '_去重.xls',
            sheet_name=goods[GoodsId])
        z.write(f'{PATH}/{goods[GoodsId]}_租赁成交记录_' + time.strftime("%Y%m%d") + '_去重.xls')
    z.write(EXCELFILENAME)
    z.close()
    shutil.rmtree(PATH)
    os.remove(EXCELFILENAME)


if __name__ == '__main__':
    duplicateExcelAndSaveToPath()
