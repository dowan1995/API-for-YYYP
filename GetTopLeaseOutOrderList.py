import os
import json
import xlrd
import xlwt
import time
import requests
import pandas as pd
from goods import goods
from xlutils.copy import copy

goods_id = 740

url = f"https://api.youpin898.com/api/trade/Order/GetTopLeaseOutOrderList?TemplateId={goods_id}"


def request_data():
    """
    GET 获取接口返回的 JSON数据
    """
    req = requests.get(url, timeout=30)
    req_json = req.json()
    return req_json


def json_file():
    """
    判断文件是否存在，不存在则创建；
    在该文件夹下创建 HTTP GET 返回的 JSON 数据文件
    """
    with open(f"{goods_id}_data.json", "w") as fp:
        fp.write(json.dumps(request_data()['Data'], sort_keys=True, indent=4, separators=(',', ': ')))


def main():
    request_data()
    json_file()

    filename_date_string = time.strftime("%Y%m%d")

    # 创建 Excel 工作表
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('sheet1')

    # 设置表头
    worksheet.write(0, 0, label="租赁价格")
    worksheet.write(0, 1, label="租赁时长")
    worksheet.write(0, 2, label="租赁类型")
    worksheet.write(0, 3, label="押金")
    worksheet.write(0, 4, label="成交时间")
    worksheet.write(0, 5, label="数据获取时间")

    # 如果不存在该文件就创建
    filename = f'{goods[goods_id]}_租赁成交记录_' + filename_date_string + '.xls'
    if not os.path.exists(filename):
        workbook.save(filename)

    # 读取 JSON 文件后删除该文件
    with open(f'{goods_id}_data.json', 'r') as f:
        data = json.load(f)
#         os.remove(f'{goods_id}_data.json')

    # pandas 读取 Excel，获取当前行数
    df = pd.read_excel(filename, engine='xlrd', sheet_name='sheet1')
    row_num = len(df.index.values) + 1

    # 多次写入，将旧 workbook 拷贝进新的 workbook，并在新的 workbook 基础上 write 数据
    oldworkbook = xlrd.open_workbook(filename)
    newworkbook = copy(oldworkbook)
    newworksheet = newworkbook.get_sheet('sheet1')

    # JSON数据写入表格
    val1 = row_num
    val2 = row_num
    val3 = row_num
    val4 = row_num
    val5 = row_num
    val6 = row_num
    for list_item in data:
        for key, value in list_item.items():
            data_date_string = time.strftime("%Y%m%d-%H%M%S")
            if key == "LeaseUnitPrice":
                newworksheet.write(val1, 0, value)
                val1 += 1
            elif key == "LeaseDays":
                newworksheet.write(val2, 1, value)
                val2 += 1
            elif key == "Type":
                if value == 1:
                    value = "长租"
                elif value == 2:
                    value = "短租"
                newworksheet.write(val3, 2, value)
                val3 += 1
            elif key == "LeaseDeposit":
                newworksheet.write(val4, 3, value)
                val4 += 1
            elif key == "DateTime":
                newworksheet.write(val5, 4, value)
                val5 += 1
            else:
                pass
        newworksheet.write(val6, 5, data_date_string)
        val6 += 1
    newworkbook.save(filename)


if __name__ == '__main__':
    main()
