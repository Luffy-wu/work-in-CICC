# -*- coding: utf-8 -*-

import numpy as np
import pandas as pd
import xlwings as xw
import os

wb = xw.Book('result.xlsm')
sht = wb.sheets('基金数量')

dir_path = os.path.dirname(wb.fullname)
data = pd.read_excel(dir_path + '\\weekly_return.xlsx', header=0)
sub_data = data[data['date'] == sht.range('C1').value]

def scale_stats():
    count = 0
    type_list = ['股票多头', '股票多空', '市场中性', '债券策略', '宏观策略', '事件驱动',
             '相对价值', '管理期货', '多策略', '组合策略', '其他一级策略']
    for cate in type_list:
        sht.range('A3').offset(count, 0).value = cate
        sht.range('B3').offset(count, 0).value = sub_data[sub_data['type'] == cate].index.size
        count = count + 1