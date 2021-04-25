import openpyxl
import glob
import csv
from pprint import pprint
import re

# aaaaaa

new_cols_dict = {
    "日付": 1, #
    "決済方法": 3,
    "注文時間": 4, #
    "リピーターフラグ":5,
    "注文者": 6, #
    "注文者会社名":7,
    "商品名":8,
    "個数":9,
    "商品価格":10,
    "送料":11,
    "手数料":12,
    "消費税":13,
    "注文金額":14,
    "処理状況":15,
    "方法":16,
    "出荷日":17,
    "担当":18,
    "件数": 2
}

dates_dict = {}

payments_dict = {
    "代金引き換え":"D",
    "NP掛け払い(FREX B2B 後払)":"NK",
    "NP後払い(請求書後払い)":"NP"
}

wb_daicho = openpyxl.load_workbook('daicho01.xlsx')
ws_daicho = wb_daicho["受注管理表"]
ws_daicho_max_row = ws_daicho.max_row

with open("excel/honten_dummy20210417.csv", "r", encoding="shift-jis") as f:
    reader = csv.DictReader(f)

    for row in reader:
        print(row["日付"] +":"+ row["注文時間"]+":"+row["注文者"]+":"+row["商品名"])
        for field in reader.fieldnames:


