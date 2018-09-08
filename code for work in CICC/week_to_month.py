# -*- coding: utf-8 -*-

import numpy as np
import pandas as pd
import xlwings as xw
import os

wb = xw.Book('result.xlsm')
sht = wb.sheets('业绩数据')
start_date = sht.range('C1').value
end_date = sht.range('C2').value

dir_path = os.path.dirname(wb.fullname)
data = pd.read_excel(dir_path + '\\weekly_return.xlsx', header=0)

def get_month_rtn(df):
    sub_data = data[data['fund_name'] == df.fund_name.min()]
    nav1 = df[df['date'] == df['date'].min()]['last_nav'].max()
    if df['date'].min() > sub_data['date'].min():
        date1 = df['date'].min() - pd.Timedelta(days=7)
        while sub_data[sub_data['date'] == date1].size == 0:
            date1 = date1 - pd.Timedelta(days=7)
        nav1 = sub_data[sub_data['date'] == date1]['nav'].max()
    
    nav2 = df[df['date'] == df['date'].max()]['nav'].max()
    return nav2/nav1 - 1

def get_std_a(df):
    return np.std(df['weekly_return'])*np.sqrt(52)
    
def get_max_retra(df, start_date):
    nav_list = list(df['nav'])
    if (df['date'].min() - pd.Timedelta(days=7)) >= start_date:
        nav_list.insert(0, df['last_nav'][df.index[0]])                          
    
    peak_list = [max(nav_list[:i+1]) for i in range(len(nav_list))]
    return ((pd.Series(nav_list) / pd.Series(peak_list)) - 1).min()
    
def main_fun():
    sub_data = data[(data['date'] >= start_date) & (data['date'] < end_date)]
    
    # filter funds with less than 3 records in a month
    grouped = sub_data.groupby('fund_name')
    name_list = (grouped.apply(len) >= 3)[(grouped.apply(len) >= 3)].index
    sub_data = sub_data[sub_data['fund_name'].isin(name_list)]
    
    # calculate annual return, annual std, sharpe ratio, max retracement
    grouped = sub_data.groupby('fund_name')
    
    cate = grouped['type'].apply(min)
    rtn_m = grouped.apply(lambda x: get_month_rtn(x))
    rtn_a = rtn_m.apply(lambda x: (1+x)**(12) - 1)
    std_a = grouped.apply(get_std_a)
    sharpe = (rtn_a - 0.015)/std_a
    max_retra = grouped.apply(lambda x: get_max_retra(x, start_date))


    result = pd.DataFrame({'type':cate, 'annual_return':rtn_a, 'annual_std':std_a, 'sharpe':sharpe, 'max_retra':max_retra, 'month_return':rtn_m},
                      index = rtn_a.index)
    sht.range('A3').expand().clear_contents()
    sht.range('A3').value = result