import os
import json
import xlrd
import xlwt
import time
import requests
from goods import goods
from xlutils.copy import copy


def requestApiData(url):
    """
    GET 获取接口返回的 JSON 数据
    """
    req = requests.get(url, timeout=30)
    req_json = req.json()
    return req_json


def createJsonFile(goodsid):
    """
    判断文件是否存在，不存在则创建；
    在该文件夹下创建 HTTP GET 返回的 JSON 数据文件
    """
    with open(f"{goodsid}_data.json", "w") as fp:
        fp.write(json.dumps(requestApiData(API_URL)['Data'], sort_keys=True, indent=4, separators=(',', ': ')))


def readDataFromJson(goodsid):
    """

    :return:
    """
    with open(f'{goodsid}_data.json', 'r') as f:
        data = json.load(f)
        os.remove(f'{goodsid}_data.json')
    return data


def createExcel(excelfilename):
    """
    :return:
    """
    # 创建 Excel
    workbook = xlwt.Workbook(encoding='utf-8')
    workbook.add_sheet("initialization")
    workbook.save(excelfilename)


def createSheet(oldworkbook, goodsid, excelfilename):
    if f"{goods[goodsid]}" in oldworkbook.sheet_names():
        pass
    else:
        NewWorkBook = copy(oldworkbook)
        NewWorkSheet = NewWorkBook.add_sheet(goods[goodsid])
        NewWorkSheet.write(0, 0, label="租赁价格")
        NewWorkSheet.write(0, 1, label="租赁时长")
        NewWorkSheet.write(0, 2, label="租赁类型")
        NewWorkSheet.write(0, 3, label="押金")
        NewWorkSheet.write(0, 4, label="成交时间")
        NewWorkBook.save(excelfilename)


def writeDataToSheet(goodsid, excelfilename, data):
    OldWorkBook = xlrd.open_workbook(excelfilename)
    NewWorkBook = copy(OldWorkBook)
    NewWorkSheet = NewWorkBook.get_sheet(goods[goodsid])

    RowCount = xlrd.open_workbook(ExcelFileName).sheet_by_name(goods[GoodsId]).nrows

    # JSON数据写入表格
    val1 = RowCount
    val2 = RowCount
    val3 = RowCount
    val4 = RowCount
    val5 = RowCount
    for list_item in data:
        for key, value in list_item.items():
            if key == "LeaseUnitPrice":
                NewWorkSheet.write(val1, 0, value)
                val1 += 1
            elif key == "LeaseDays":
                NewWorkSheet.write(val2, 1, value)
                val2 += 1
            elif key == "Type":
                if value == 1:
                    value = "长租"
                elif value == 2:
                    value = "短租"
                NewWorkSheet.write(val3, 2, value)
                val3 += 1
            elif key == "LeaseDeposit":
                NewWorkSheet.write(val4, 3, value)
                val4 += 1
            elif key == "DateTime":
                NewWorkSheet.write(val5, 4, value)
                val5 += 1
            else:
                pass
    NewWorkBook.save(excelfilename)


if __name__ == '__main__':
    # Excel 文件名
    ExcelFileName = f'租赁成交记录_' + time.strftime("%Y%m%d") + '.xls'

    for GoodsId in goods:
        # 接口地址
        API_URL = f"https://api.youpin898.com/api/trade/Order/GetTopLeaseOutOrderList?TemplateId={GoodsId}"

        # 初始 Excel 生成
        if not os.path.exists(ExcelFileName):
            createExcel(ExcelFileName)

        # 判断有无对应商品名称的 Sheet
        createSheet(xlrd.open_workbook(ExcelFileName), GoodsId, ExcelFileName)

        # 创建 JSON 文件
        createJsonFile(GoodsId)

        # 根据 JSON 文件向表格内的 Sheet 写入数据
        writeDataToSheet(GoodsId, ExcelFileName, readDataFromJson(GoodsId))
