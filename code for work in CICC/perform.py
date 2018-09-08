# -*- coding: utf-8 -*-

import numpy as np
import pandas as pd
import xlwings as xw

wb = xw.Book('result.xlsm')
sht1 = wb.sheets('业绩数据')
sht2 = wb.sheets('业绩统计')

rows = sht1.range('A4').expand().value
df = pd.DataFrame(rows, columns = ['fund_name', 'annual_return', 'annual_std', 'max_retra', 'month_return', 'annual_sharpe', 'type'])

type_list = ['股票多头', '股票多空', '市场中性', '债券策略', '宏观策略', '事件驱动',
             '相对价值', '管理期货', '多策略', '组合策略', '其他一级策略']
             
def performance():
    count = 0
    for cate in type_list:
        sht2.range('A2').offset(4*count,0).value = cate
        
        sub_df = df[df['type'] == cate]
        
   #     sht2.range('C2').offset(4*count,0).value = sub_df['annual_return'].quantile(0.75)
   #     sht2.range('C3').offset(4*count,0).value = sub_df['annual_return'].quantile(0.5)
   #     sht2.range('C4').offset(4*count,0).value = sub_df['annual_return'].mean()
   #     sht2.range('C5').offset(4*count,0).value = sub_df['annual_return'].quantile(0.25)
        
        sht2.range('E2').offset(4*count,0).value = sub_df['annual_std'].quantile(0.75)
        sht2.range('E3').offset(4*count,0).value = sub_df['annual_std'].quantile(0.5)
        sht2.range('E4').offset(4*count,0).value = sub_df['annual_std'].mean()
        sht2.range('E5').offset(4*count,0).value = sub_df['annual_std'].quantile(0.25)
        
   #     sht2.range('E2').offset(4*count,0).value = sub_df['annual_sharpe'].quantile(0.75)
   #     sht2.range('E3').offset(4*count,0).value = sub_df['annual_sharpe'].quantile(0.5)
   #     sht2.range('E4').offset(4*count,0).value = sub_df['annual_sharpe'].mean()
   #     sht2.range('E5').offset(4*count,0).value = sub_df['annual_sharpe'].quantile(0.25)
        
        sht2.range('G2').offset(4*count,0).value = sub_df['max_retra'].quantile(0.75)
        sht2.range('G3').offset(4*count,0).value = sub_df['max_retra'].quantile(0.5)
        sht2.range('G4').offset(4*count,0).value = sub_df['max_retra'].mean()
        sht2.range('G5').offset(4*count,0).value = sub_df['max_retra'].quantile(0.25)
        
        sht2.range('C2').offset(4*count,0).value = sub_df['month_return'].quantile(0.75)
        sht2.range('C3').offset(4*count,0).value = sub_df['month_return'].quantile(0.5)
        sht2.range('C4').offset(4*count,0).value = sub_df['month_return'].mean()
        sht2.range('C5').offset(4*count,0).value = sub_df['month_return'].quantile(0.25)
        
        count = count+1