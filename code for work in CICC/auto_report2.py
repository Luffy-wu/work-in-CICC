# -*- coding: utf-8 -*-

import numpy as np
import pandas as pd
import xlwings as xw
from datetime import datetime
from string import digits
from re import match
import os
from glob import glob

class AssetClass(object):
    bond = '债券'
    bond_subtypes = ['债券','ABS', '商业性债', '私募债', '政策性债', '央行票据', 
                     '分离债', '次级债金券', '企债', '国债', '短期融资', 
                     '可转债', '公司债', '标准券','可交换债']
    
    equity = '股票'
    equity_subtypes = ['非公优先股', '股票', '优先股', '创业板', 'B转H']
    
    fund = '基金'
    fund_subtypes = ['ETF', '开放基金', '基金']
    
    derivative = '衍生品'
    derivative_subtypes = ['SWAP', '权证', '期货', 
                           '指数', '期权', '股指期货', '商品期货', '国债期货']
    cash = '现金'
    cash_subtypes = ['银行存款', '清算备付金', '券商保证金']
    
    margin = '保证金'
    margin_subtypes = ['期货交易存出保证金', '个股期权存出保证金']    
        
    others = '其他'
    others_subtypes = ['债券借贷', '指定', '网络服务','正回购','逆回购']

    not_cash_types = ['债券', '股票', '基金', '期货交易存出保证金', '个股期权存出保证金','资产管理计划','SWAP','正回购','逆回购','ABS','应收利息','可转债','可交换债','中期票据','短期融资券']
    
    associative_list = []
    associative_list.extend([(x, '债券') for x in bond_subtypes])
    associative_list.extend([(x, '股票') for x in equity_subtypes])
    associative_list.extend([(x, '基金') for x in fund_subtypes])
    associative_list.extend([(x, '衍生品') for x in derivative_subtypes])
    associative_list.extend([(x, '现金') for x in cash_subtypes])
    associative_list.extend([(x, '保证金') for x in margin_subtypes])
    associative_list.extend([(x, '其他') for x in others_subtypes])
    
    mapping = dict(associative_list)

class AccountingSubjects(object):
    # 银行存款
    cash_deposits = '1002'
    
    # 清算备付金
    cash_provisions = '1021'
    
    # 证券清算款 
    cash_Liquidation = '3003'
    
    # 存出保证金
    margin = '1031'
    margin_broker = '103106'
    margin_future = '103113'
    margin_option = '103131'
    
    money_all = [cash_deposits,cash_Liquidation]
    money_all_digits = 8
    
    other_money = [cash_provisions,margin]
    other_money_digits = 6
    
    # 股票投资
    equities = '1102'
    equities_digits = 14
    
    # 债券
    bonds = '1103'
    bonds_sh = ['11030','11031','11032']
    bonds_sz = ['11033','11034','11035']
    bonds_ib = ['11036','11037']
    
    #应收债券和银行利息
    bonds_intrest = '1204'
    interest_bank = '12040101'
    interest_sh = ['1204101','1204130']
    interest_sz = ['1204103','1204330']
    interest_ib = ['1204105','1204530']
    
    #应收股利或基金红利
    dividend = '1203'
    
    #回购 
    buyback = ['2202','1202']
    buyback_digits = 14
    
    #ABS
    ABS = '1104'
    ABS_digits = 14
    
    # 基金投资
    funds = '1105'
    
    #信托投资
    trust = '1201'
    trust_fund = '12010401'
    
    # 互换和场外其他
    swaps = '1107'
    
    # 场内期权
    options = '1041'
    options_margins = '104102'
    options_mktValue = '104103'
    
    # 衍生工具和套期工具
    derivatives = '3102'
    hedging_positions = '3201'
    index_futures = ['IC', 'IF', 'IH']
    bond_futures = ['TF','T']
    futures_digits = 14
    all_derivatives = [derivatives, hedging_positions, swaps]
    
    examine = [cash_deposits, 
               cash_provisions, 
               equities,
               margin,
               margin_broker,
               margin_future,
               margin_option,
               derivatives, 
               hedging_positions, 
               swaps,
               bonds, 
               funds, 
               options,
               options_margins,
               options_mktValue,ABS,
               '资产类合计:','资产类合计：']
               
    set_examime = set(examine)
               
# helper function
remove_digits = str.maketrans('', '', digits)
def map_name_to_bond_subtype(name):
    if name[-2::] == 'EB':
        return '可交换债'
    elif name[-2::] == '转债':
        return '可转债'
    elif name[-6:-3] == 'MTN':
        return '中期票据'
    elif name[-5:-3] == 'CP':
        return '短期融资券'
    else: return '债券'
    
def map_code_to_future_ticker(code):
    first_letter = code.translate(remove_digits)[0]       
    ind = code.index(first_letter)
    return code[ind::]

def map_code_to_bond_ticker(code):
    prefix = code[8::]
    if code[:5] in ['11030','11031']:
        suffix = '.SH'
    elif code[:5] in ['11033']:
        suffix = '.SZ'
    else:
        suffix = '.IB'
    return prefix + suffix

def map_code_to_bond_ticker1(code):
    prefix = code[8::]
    if code[:5] in AccountingSubjects.bonds_sh:
        suffix = '.SH'
    elif code[:5] in AccountingSubjects.bonds_sz:
        suffix = '.SZ'
    else:
        suffix = '.IB'
    return prefix + suffix

def map_code_to_interest_ticker(code):
    prefix = code[8::]
    if code[:7] in AccountingSubjects.interest_sh:
        suffix = '.SH'
    elif code[:7] in AccountingSubjects.interest_sz:
        suffix = '.SZ'
    elif code[:7] in AccountingSubjects.interest_ib:
        suffix = '.IB'
    else:suffix = ''
    return prefix + suffix

def test_bond_future(ticker):
    for stamp in AccountingSubjects.bond_futures:
        if stamp in ticker:
            return True
    return False

def test_index_future(ticker):
    for stamp in AccountingSubjects.index_futures:
        if stamp in ticker:
            return True
    return False
def export_positions(file, counter):#从估值表中提取数据的函数
    wb = xw.Book(file)
    sht = wb.sheets[0]
    time_temp = list(sht.range('A1:A5').value)
    #找到列标题所在行，
    if  '科目代码' in time_temp:
        header_index = time_temp.index('科目代码')
    elif '科目编码' in time_temp:
        header_index = time_temp.index('科目编码')
    else:
        header_index = list(sht.range('B1:B5').value).index('科目代码')
    time_cell = [l for l in file if l in digits]
    if len(time_cell)>=8: #如果估值表的名称里有时间，则从估值表名称里提取
        time_cell = time_cell[-8::]
        time_cell_str = ''.join(time_cell)
        this_date = datetime.strptime(time_cell_str, '%Y%m%d')
    else:#如果估值表里没有时间，就去找A3单元格
        this_date = time_temp[2]
    df = pd.read_excel(file,header=header_index)
        #去掉列名中的空格,给列重命名
    column_name = [col.replace(' ','') for col in df.columns] #逐一检查，去掉空格，生成新列名
    column_name[0] = '科目代码' #第一列统一命名为科目代码，因为有的傻逼的表第一列叫科目编码              
    if '证券市值' in column_name:
        column_name[column_name.index('证券市值')] = '市值'
    df.columns = column_name
    if not AccountingSubjects.set_examime.intersection(set(df['科目代码'].values)):
        # empty table
        wb.close()
        return [],[]
    #找到资产净值和总资产
    col = list(df['科目代码'])
    if '基金资产净值:' in col:
        net_row = col.index('基金资产净值:')#种类为基金估值表
    elif '集合计划资产净值：' in col:
        net_row = col.index('集合计划资产净值：')  #种类为集合计划估值表
    elif '资产资产净值：' in col:
        net_row = col.index('资产资产净值：')
    else: 
        #print("表中没找到净资产，用比例来计算净资产")
        net_row = []
    if '资产类合计：' in col:
        asset_row = col.index('资产类合计：')
    elif '资产类合计:' in col:
        asset_row = col.index('资产类合计:')
    else: 
        #print("表中没找到总资产")
        asset_row = [] 
    assets = dict()
    if  not net_row or not asset_row: #如果没找到净资产或者总资产，就用证券投资合计/占总资产比例来计算 
        col_temp = ''
        if '证券投资合计:' in col:
            col_temp = '证券投资合计:'
        elif '证券投资合计：' in col:
            col_temp = '证券投资合计：'
        else:
            col_temp = df['科目代码'][1] #都找不到就去用第二行的市值和占净值比例算
        if not '市值占净值%' in column_name:#只有洛书出现这种情况,数量其实是比例 
            asset_row = col.index(col_temp)
            assets['val'] = df['市值'][asset_row]/df['数量'][asset_row]
        else:
            asset_row = col.index(col_temp)
            if isinstance(df['市值'][asset_row],str):
                assets['val'] = float(df['市值'][asset_row].replace(',',''))/float(df['市值占净值%'][asset_row].replace('%',''))*100
            else:assets['val'] = df['市值'][asset_row]/df['市值占净值%'][asset_row]
        assets['net'] = assets['val']
        assets['ratio'] = 1
    else:
        if isinstance(df['市值'][asset_row],str):
            assets['val'] = float(df['市值'][asset_row].replace(',',''))
            assets['net'] = float(df['市值'][net_row].replace(',',''))
        else:
            assets['val'] = df['市值'][asset_row]
            assets['net'] = df['市值'][net_row]
        assets['ratio'] = assets['val']/assets['net']
    output = []
    flg=0 #判断是否是新版估值表，如果是则债券代码需要修改
    ###########################################################################
    for row in range(0,df.index.size):
        code = df['科目代码'][row]
        if isinstance(code,str):
            code=code.replace(".","")
            code=code.replace(" ","")
            if code[-2:] in ["SH","SZ","CY"]:
                code=code[:-2]
                flg=1
        #过滤南方基金的表的第一列中的数字,并提取累计净值,实收资本
        if isinstance(code,str) :
            temp = code.translate(remove_digits)
            if temp == '实收资本':
                capital = df['市值'][row]
            if temp != '':
                code = temp
                if '累计单位净值' in code:
                    a_net_value = df['科目名称'][row]
                if '昨日单位净值' in code:
                    last_net_value = df['科目名称'][row]
                if '今日单位净值' in code:
                    today_net_value = df['科目名称'][row]
                if '期初单位净值' in code:
                    initial_net_value = df['科目名称'][row]
            if match('^[a-zA-Z]+$',temp):  #易方达的表里面MTN的代码有四个字母,判断有字母
                code = df['科目代码'][row]
       ###############################
        mktVal = df['市值'][row]
        if not mktVal or row == 0:#过滤
            continue
        mktVal = float(mktVal.replace(',','')) if isinstance(mktVal,str) else mktVal
        mktVal_r = mktVal/assets['val']#此处用的总资产
        qty = df['数量'][row] if '数量' in column_name else df['证券数量'][row]
        if isinstance(qty,str):
            if '%' in qty:
                qty = float(qty.replace('%',''))/100 
            else:qty = float(qty.replace(',',''))   
        name = df['科目名称'][row]
        valAdd = df['估值增值'][row] if '估值增值' in column_name else 0
        if '市价' in column_name:
            p = df['市价'][row]
        elif '行情收市价' in column_name: 
            p = df['行情收市价'][row]
        else: p = df['行情价格'][row] if '行情价格' in column_name else mktVal/qty 
        # deal with margin
        if code == AccountingSubjects.margin_future:
            a_subtype = '期货交易存出保证金'
            a_type = AssetClass.mapping[a_subtype]
            pos = [this_date, 'cash_CNY', 'cash_CNY', mktVal, 1, 
                   mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
            output.append(pos)        
        elif code == AccountingSubjects.margin_option:
            a_subtype  = '个股期权存出保证金'
            a_type = AssetClass.mapping[a_subtype]
            pos = [this_date, 'cash_CNY', 'cash_CNY', mktVal, 1, 
                   mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
            output.append(pos)
        # filter
        elif not isinstance(name,str):
            continue
        # deal with equities
        elif code[:4] == AccountingSubjects.equities:
            if np.isnan(qty) or np.isnan(mktVal):
                # no quantity or ratio or mktvalue
                continue
            if len(code) == AccountingSubjects.equities_digits:
                ticker = code[8::]
                if ticker[0] == 'H':
                    # deal with HK share
                    ticker = str(int(ticker[1::])) + '.HK'
                a_subtype = '股票'
                a_type = AssetClass.mapping[a_subtype]
                if np.isnan(mktVal_r): mktVal_r = 0
                pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
                output.append(pos)
        
        # deal with funds
        elif code[:4] == AccountingSubjects.funds:
            if np.isnan(qty) or np.isnan(mktVal):
                # no quantity or ratio or mktvalue
                continue
            if len(code) == AccountingSubjects.equities_digits:
                ticker = code[8::]
                a_subtype = '基金'
                a_type = AssetClass.mapping[a_subtype]
                if np.isnan(mktVal_r): mktVal_r = 0
                pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
                output.append(pos)
        #deal with trust
        elif code[:8] ==  AccountingSubjects.trust_fund:
            if len(code)>8:
                ticker = code[8::]
                a_subtype = '基金'
                a_type = AssetClass.mapping[a_subtype]
                if np.isnan(mktVal_r): mktVal_r = 0
                pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
                output.append(pos)
        
        # deal with bonds
        elif code[:4] == AccountingSubjects.bonds:
            if  np.isnan(mktVal):#########np.isnan(qty) or
                # no quantity or ratio or mktvalue
                continue
            if len(code) == AccountingSubjects.equities_digits or len(code) == 17 and code[0] in digits:
                
                a_subtype = map_name_to_bond_subtype(name)
                if flg==1:
                    ticker = map_code_to_bond_ticker1(code)
                else:
                    ticker = map_code_to_bond_ticker(code)
                a_type = '债券'
                if np.isnan(mktVal_r): mktVal_r = 0
                pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
                #print(ticker, name, qty, p, mktVal, mktVal_r, a_type, a_subtype)
                output.append(pos)
        elif code[:4] == AccountingSubjects.ABS:
            if len(code) == AccountingSubjects.ABS_digits:
                ticker = map_code_to_bond_ticker(code)
                a_subtype = 'ABS'
                a_type = '债券'
                pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
                output.append(pos)
        elif code[:4] == AccountingSubjects.bonds_intrest:
            if len(code) == AccountingSubjects.equities_digits or (len(code) == 17 and code[0] in digits):
                ticker = map_code_to_interest_ticker(code)
                a_type = '现金存款类' if code[:8] == AccountingSubjects.interest_bank else '利息'
                a_subtype = '现金存款类' if code[:8] == AccountingSubjects.interest_bank else '应收利息'
                pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
                output.append(pos)
        elif code[:4] == AccountingSubjects.dividend:
            if len(code) == AccountingSubjects.equities_digits:
                ticker = code[8::]
                a_type = '红利' 
                a_subtype = '现金存款类'
                pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
                output.append(pos)
        elif code[:4] in AccountingSubjects.buyback:
            if len(code) == AccountingSubjects.buyback_digits:
                ticker = name
                a_subtype = '正回购' if code[:4]=='2202' else '逆回购'
                a_type = '其他'
                qty = 1
                p = mktVal
                mktVal = mktVal 
                mktVal_r = mktVal_r
                pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
                output.append(pos)
        elif code[:4] in AccountingSubjects.money_all and len(code) == AccountingSubjects.money_all_digits:
            ticker = code
            a_subtype = '现金存款类'
            a_type = '现金存款类'
            qty = mktVal
            p = 1
            pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
            output.append(pos)
        elif code[:4] in AccountingSubjects.other_money and len(code) == AccountingSubjects.other_money_digits:
            ticker = code
            a_subtype = '现金存款类'
            a_type = '现金存款类'
            qty = mktVal
            p = 1
            pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
            output.append(pos)
        # deal with derivatives
        elif code[:4] in AccountingSubjects.all_derivatives:
            if np.isnan(qty) or len(code)<=10:
                # no quantity
                continue
            if not code.translate(remove_digits):#如果没有字母，则是委外资产管理计划
                ticker = code
                a_type = '其他'
                a_subtype = '资产管理计划'
            elif len(code) == AccountingSubjects.futures_digits:
                ticker = map_code_to_future_ticker(code)
                if '购' in name or '沽' in name:
                    a_subtype = '期权'
                elif 'TRS' in name:
                    a_subtype = 'SWAP'
                else:
                    if test_bond_future(ticker):
                        a_subtype = '国债期货'
                    elif test_index_future(ticker):
                        a_subtype = '股指期货'
                    elif '国债' in name:
                        a_subtype = '国债期货'
                    else:
                        a_subtype = '商品期货'
                a_type = AssetClass.mapping[a_subtype]
            pos = [this_date, ticker, name, qty, p, mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
            output.append(pos)
            
        # deal with options
        elif code[:6] == AccountingSubjects.options_mktValue:
            if np.isnan(mktVal):
                # no value
                continue
            a_subtype = '场内期权'
            a_type = '场内期权'
            pos = [this_date, '场内期权总市值', '场内期权总市值', mktVal, 1, 
                   mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
            output.append(pos)
            
        # deal with options margin
        elif code[:6] == AccountingSubjects.options_margins:
            if np.isnan(mktVal):
                # no market value
                continue
            a_subtype = '场内期权保证金'
            a_type = '场内期权保证金'
            pos = [this_date, 'cash_CNY', 'cash_CNY', mktVal, 1, 
                   mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
            output.append(pos) 
    
    otherdata = [os.path.split(file)[1],this_date,a_net_value,last_net_value,today_net_value,assets['net']/10000,capital,initial_net_value]
    wb.close()
    return output,otherdata

def export_positions_New2(file, counter):#从估值表中提取数据的函数
    wb = xw.Book(file)
    sht = wb.sheets[0]
    time_temp = list(sht.range('A1:A5').value)
    #找到列标题所在行，
    if  '科目代码' in time_temp:
        header_index = time_temp.index('科目代码')
    elif '科目编码' in time_temp:
        header_index = time_temp.index('科目编码')
    else:
        header_index = list(sht.range('B1:B5').value).index('科目代码')
    time_cell = [l for l in file if l in digits]
    if len(time_cell)>=8: #如果估值表的名称里有时间，则从估值表名称里提取
        time_cell = time_cell[-8::]
        time_cell_str = ''.join(time_cell)
        this_date = datetime.strptime(time_cell_str, '%Y%m%d')
    else:#如果估值表里没有时间，就去找A3单元格
        this_date = time_temp[2]
    df = pd.read_excel(file,header=header_index)
        #去掉列名中的空格,给列重命名
    column_name = [col.replace(' ','') for col in df.columns] #逐一检查，去掉空格，生成新列名
    column_name[0] = '科目代码' #第一列统一命名为科目代码，因为有的傻逼的表第一列叫科目编码              
    if '证券市值' in column_name:
        column_name[column_name.index('证券市值')] = '市值'
    df.columns = column_name
    if not AccountingSubjects.set_examime.intersection(set(df['科目代码'].values)):
        # empty table
        wb.close()
        return []
    #找到资产净值和总资产
    col = list(df['科目代码'])
    if '基金资产净值:' in col:
        net_row = col.index('基金资产净值:')#种类为基金估值表
    elif '集合计划资产净值：' in col:
        net_row = col.index('集合计划资产净值：')  #种类为集合计划估值表
    elif '资产资产净值：' in col:
        net_row = col.index('资产资产净值：')
    else: 
        #print("表中没找到净资产，用比例来计算净资产")
        net_row = []
    if '资产类合计：' in col:
        asset_row = col.index('资产类合计：')
    elif '资产类合计:' in col:
        asset_row = col.index('资产类合计:')
    else: 
        #print("表中没找到总资产")
        asset_row = [] 
    assets = dict()
    if  not net_row or not asset_row: #如果没找到净资产或者总资产，就用证券投资合计/占总资产比例来计算 
        col_temp = ''
        if '证券投资合计:' in col:
            col_temp = '证券投资合计:'
        elif '证券投资合计：' in col:
            col_temp = '证券投资合计：'
        else:
            col_temp = df['科目代码'][1] #都找不到就去用第二行的市值和占净值比例算
        if not '市值占净值%' in column_name:#只有洛书出现这种情况,数量其实是比例 
            asset_row = col.index(col_temp)
            assets['val'] = df['市值'][asset_row]/df['数量'][asset_row]
        else:
            asset_row = col.index(col_temp)
            if isinstance(df['市值'][asset_row],str):
                assets['val'] = float(df['市值'][asset_row].replace(',',''))/float(df['市值占净值%'][asset_row].replace('%',''))*100
            else:assets['val'] = df['市值'][asset_row]/df['市值占净值%'][asset_row]
        assets['net'] = assets['val']
        assets['ratio'] = 1
    else:
        if isinstance(df['市值'][asset_row],str):
            assets['val'] = float(df['市值'][asset_row].replace(',',''))
            assets['net'] = float(df['市值'][net_row].replace(',',''))
        else:
            assets['val'] = df['市值'][asset_row]
            assets['net'] = df['市值'][net_row]
        assets['ratio'] = assets['val']/assets['net']
    output = []
    flg=0 #判断是否是新版估值表，如果是则债券代码需要修改
    ###########################################################################
    for row in range(0,df.index.size):
        code = df['科目代码'][row]
        if isinstance(code,str):
            code=code.replace(".","")
            code=code.replace(" ","")
            if code[-2:] in ["SH","SZ","CY"]:
                code=code[:-2]
                flg=1
        #过滤南方基金的表的第一列中的数字,并提取累计净值,实收资本
        if isinstance(code,str) :
            temp = code.translate(remove_digits)
            if temp == '实收资本':
                capital = df['市值'][row]
            if temp != '':
                code = temp
                if '累计单位净值' in code:
                    a_net_value = df['科目名称'][row]
                if '昨日单位净值' in code:
                    last_net_value = df['科目名称'][row]
                if '今日单位净值' in code:
                    today_net_value = df['科目名称'][row]
                if '期初单位净值' in code:
                    initial_net_value = df['科目名称'][row]
            if match('^[a-zA-Z]+$',temp):  #易方达的表里面MTN的代码有四个字母,判断有字母
                code = df['科目代码'][row]
       ###############################
        mktVal = df['市值'][row]
        if not mktVal or row == 0:#过滤
            continue
        mktVal = float(mktVal.replace(',','')) if isinstance(mktVal,str) else mktVal
        mktVal_r = mktVal/assets['val']#此处用的总资产
        qty = df['数量'][row] if '数量' in column_name else df['证券数量'][row]
        if isinstance(qty,str):
            if '%' in qty:
                qty = float(qty.replace('%',''))/100 
            else:qty = float(qty.replace(',',''))   
        name = df['科目名称'][row]
        valAdd = df['估值增值'][row] if '估值增值' in column_name else 0
        if '市价' in column_name:
            p = df['市价'][row]
        elif '行情收市价' in column_name: 
            p = df['行情收市价'][row]
        else: p = df['行情价格'][row] if '行情价格' in column_name else mktVal/qty 
        # deal with margin
        if code == AccountingSubjects.margin_future:
            a_subtype = '期货交易存出保证金'
            a_type = AssetClass.mapping[a_subtype]
            pos = [os.path.split(file)[1][:15],this_date, 'cash_CNY', 'cash_CNY', mktVal, 1, 
                   mktVal, mktVal_r, a_type, a_subtype, valAdd,counter,assets['net']]
            output.append(pos)        
        elif code == AccountingSubjects.margin_option:
            a_subtype  = '个股期权存出保证金'
            a_type = AssetClass.mapping[a_subtype]
            pos = [os.path.split(file)[1][:15],this_date, 'cash_CNY', 'cash_CNY', mktVal, 1, 
                   mktVal, mktVal_r, a_type, a_subtype, valAdd,counter,assets['net']]
            output.append(pos)
        # filter
        elif not isinstance(name,str):
            continue
        # deal with equities
        elif code[:4] == AccountingSubjects.equities:
            if np.isnan(qty) or np.isnan(mktVal):
                # no quantity or ratio or mktvalue
                continue
            if len(code) == AccountingSubjects.equities_digits:
                ticker = code[8::]
                if ticker[0] == 'H':
                    # deal with HK share
                    ticker = str(int(ticker[1::])) + '.HK'
                a_subtype = '股票'
                a_type = AssetClass.mapping[a_subtype]
                if np.isnan(mktVal_r): mktVal_r = 0
                pos = [os.path.split(file)[1][:15],this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter,assets['net']]
                output.append(pos)
        
        # deal with funds
        elif code[:4] == AccountingSubjects.funds:
            if np.isnan(qty) or np.isnan(mktVal):
                # no quantity or ratio or mktvalue
                continue
            if len(code) == AccountingSubjects.equities_digits:
                ticker = code[8::]
                a_subtype = '基金'
                a_type = AssetClass.mapping[a_subtype]
                if np.isnan(mktVal_r): mktVal_r = 0
                pos = [os.path.split(file)[1][:15],this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter,assets['net']]
                output.append(pos)
        #deal with trust
        elif code[:8] ==  AccountingSubjects.trust_fund:
            if len(code)>8:
                ticker = code[8::]
                a_subtype = '基金'
                a_type = AssetClass.mapping[a_subtype]
                if np.isnan(mktVal_r): mktVal_r = 0
                pos = [os.path.split(file)[1][:15],this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter,assets['net']]
                output.append(pos)
        
        # deal with bonds
        elif code[:4] == AccountingSubjects.bonds:
            if  np.isnan(mktVal):#########np.isnan(qty) or
                # no quantity or ratio or mktvalue
                continue
            if len(code) == AccountingSubjects.equities_digits or len(code) == 17 and code[0] in digits:
                
                a_subtype = map_name_to_bond_subtype(name)
                if flg==1:
                    ticker = map_code_to_bond_ticker1(code)
                else:
                    ticker = map_code_to_bond_ticker(code)
                a_type = '债券'
                if np.isnan(mktVal_r): mktVal_r = 0
                pos = [os.path.split(file)[1][:15],this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter,assets['net']]
                #print(ticker, name, qty, p, mktVal, mktVal_r, a_type, a_subtype)
                output.append(pos)
        elif code[:4] == AccountingSubjects.ABS:
            if len(code) == AccountingSubjects.ABS_digits:
                ticker = map_code_to_bond_ticker(code)
                a_subtype = 'ABS'
                a_type = '债券'
                pos = [os.path.split(file)[1][:15],this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter,assets['net']]
                output.append(pos)
        elif code[:4] == AccountingSubjects.bonds_intrest:
            if len(code) == AccountingSubjects.equities_digits or (len(code) == 17 and code[0] in digits):
                ticker = map_code_to_interest_ticker(code)
                a_type = '现金存款类' if code[:8] == AccountingSubjects.interest_bank else '利息'
                a_subtype = '现金存款类' if code[:8] == AccountingSubjects.interest_bank else '应收利息'
                pos = [os.path.split(file)[1][:15],this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter,assets['net']]
                output.append(pos)
        elif code[:4] == AccountingSubjects.dividend:
            if len(code) == AccountingSubjects.equities_digits:
                ticker = code[8::]
                a_type = '红利' 
                a_subtype = '现金存款类'
                pos = [os.path.split(file)[1][:15],this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter,assets['net']]
                output.append(pos)
        elif code[:4] in AccountingSubjects.buyback:
            if len(code) == AccountingSubjects.buyback_digits:
                ticker = name
                a_subtype = '正回购' if code[:4]=='2202' else '逆回购'
                a_type = '其他'
                qty = 1
                p = mktVal
                mktVal = mktVal 
                mktVal_r = mktVal_r
                pos = [os.path.split(file)[1][:15],this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter,assets['net']]
                output.append(pos)
        elif code[:4] in AccountingSubjects.money_all and len(code) == AccountingSubjects.money_all_digits:
            ticker = code
            a_subtype = '现金存款类'
            a_type = '现金存款类'
            qty = mktVal
            p = 1
            pos = [os.path.split(file)[1][:15],this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter,assets['net']]
            output.append(pos)
        elif code[:4] in AccountingSubjects.other_money and len(code) == AccountingSubjects.other_money_digits:
            ticker = code
            a_subtype = '现金存款类'
            a_type = '现金存款类'
            qty = mktVal
            p = 1
            pos = [os.path.split(file)[1][:15],this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter,assets['net']]
            output.append(pos)
        # deal with derivatives
        elif code[:4] in AccountingSubjects.all_derivatives:
            if np.isnan(qty) or len(code)<=10:
                # no quantity
                continue
            if not code.translate(remove_digits):#如果没有字母，则是委外资产管理计划
                ticker = code
                a_type = '其他'
                a_subtype = '资产管理计划'
            elif len(code) == AccountingSubjects.futures_digits:
                ticker = map_code_to_future_ticker(code)
                if '购' in name or '沽' in name:
                    a_subtype = '期权'
                elif 'TRS' in name:
                    a_subtype = 'SWAP'
                else:
                    if test_bond_future(ticker):
                        a_subtype = '国债期货'
                    elif test_index_future(ticker):
                        a_subtype = '股指期货'
                    elif '国债' in name:
                        a_subtype = '国债期货'
                    else:
                        a_subtype = '商品期货'
                a_type = AssetClass.mapping[a_subtype]
            pos = [os.path.split(file)[1][:15],this_date, ticker, name, qty, p, mktVal, mktVal_r, a_type, a_subtype, valAdd,counter,assets['net']]
            output.append(pos)
            
        # deal with options
        elif code[:6] == AccountingSubjects.options_mktValue:
            if np.isnan(mktVal):
                # no value
                continue
            a_subtype = '场内期权'
            a_type = '场内期权'
            pos = [os.path.split(file)[1][:15],this_date, '场内期权总市值', '场内期权总市值', mktVal, 1, 
                   mktVal, mktVal_r, a_type, a_subtype, valAdd,counter,assets['net']]
            output.append(pos)
            
        # deal with options margin
        elif code[:6] == AccountingSubjects.options_margins:
            if np.isnan(mktVal):
                # no market value
                continue
            a_subtype = '场内期权保证金'
            a_type = '场内期权保证金'
            pos = [os.path.split(file)[1][:15],this_date, 'cash_CNY', 'cash_CNY', mktVal, 1, 
                   mktVal, mktVal_r, a_type, a_subtype, valAdd,counter,assets['net']]
            output.append(pos) 
    
    #otherdata = [os.path.split(file)[1],this_date,a_net_value,last_net_value,today_net_value,assets['net']/10000,capital,initial_net_value]
    wb.close()
    return output
    
#aborted
def export_asset4cmpr(file, counter):#从估值表中提取数据的函数
    wb = xw.Book(file)
    sht = wb.sheets[0]
    time_temp = list(sht.range('A1:A5').value)
    #找到列标题所在行，
    if  '科目代码' in time_temp:
        header_index = time_temp.index('科目代码')
    elif '科目编码' in time_temp:
        header_index = time_temp.index('科目编码')
    else:
        header_index = list(sht.range('B1:B5').value).index('科目代码')
    time_cell = [l for l in file if l in digits]
    if len(time_cell)>=8: #如果估值表的名称里有时间，则从估值表名称里提取
        time_cell = time_cell[-8::]
        time_cell_str = ''.join(time_cell)
        this_date = datetime.strptime(time_cell_str, '%Y%m%d')
    else:#如果估值表里没有时间，就去找A3单元格
        this_date = time_temp[2]
    df = pd.read_excel(file,header=header_index)
    #去掉列名中的空格,给列重命名
    column_name = [col.replace(' ','') for col in df.columns] #逐一检查，去掉空格，生成新列名
    column_name[0] = '科目代码' #第一列统一命名为科目代码，因为有的傻逼的表第一列叫科目编码              
    if '证券市值' in column_name:
        column_name[column_name.index('证券市值')] = '市值'
    df.columns = column_name
    if not AccountingSubjects.set_examime.intersection(set(df['科目代码'].values)):
        # empty table
        wb.close()
        return [],[]
    col = list(df['科目代码'])
    #找到股票、基金、期货市值合计
    stk_not_found = 0
    fnd_not_found = 0
    drv_not_found = 0
    if '其中股票投资' in col:
        stocks_total_row = col.index('其中股票投资')
    else:
        stk_not_found = 1
    if '其中基金投资' in col:
        funds_total_row = col.index('其中基金投资')
    else:
        fnd_not_found = 1
    if '其中其他衍生工具投资' in col:
        derivatives_total_row = col.index('其中其他衍生工具投资')
    elif '股指期货投资合计：' in col:
        derivatives_total_row = col.index('股指期货投资合计：')
    elif '股指期货投资合计:' in col:
        derivatives_total_row = col.index('股指期货投资合计:')
    else:
        drv_not_found = 1
    if '资产资产净值：' in col:
        value_total_row = col.index('资产资产净值：')
    elif '基金资产净值：' in col:
        value_total_row = col.index('基金资产净值：')
    elif '资产资产净值:' in col:
        value_total_row = col.index('资产资产净值:')
    elif '基金资产净值:' in col:
        value_total_row = col.index('基金资产净值:')
    elif '集合计划资产净值:' in col:
        value_total_row = col.index('集合计划资产净值:')
    elif '集合计划资产净值:' in col:
        value_total_row = col.index('集合计划资产净值:')
    ######
    #找到资产净值和总资产
    if '基金资产净值:' in col:
        net_row = col.index('基金资产净值:')#种类为基金估值表
    elif '集合计划资产净值：' in col:
        net_row = col.index('集合计划资产净值：')  #种类为集合计划估值表
    elif '资产资产净值：' in col:
        net_row = col.index('资产资产净值：')
    else: 
        #print("表中没找到净资产，用比例来计算净资产")
        net_row = []
    if '资产类合计：' in col:
        asset_row = col.index('资产类合计：')
    elif '资产类合计:' in col:
        asset_row = col.index('资产类合计:')
    else: 
        #print("表中没找到总资产")
        asset_row = [] 
    assets = dict()
    
    assets['vltt'] = df['市值'][value_total_row] 
    assets['stk'] = df['市值'][stocks_total_row] if stk_not_found == 0 else 0
    assets['fnd'] = df['市值'][funds_total_row] if fnd_not_found == 0 else 0
    assets['drv'] = df['市值'][derivatives_total_row] if drv_not_found == 0 else 0
    assets['vltt'] = float(str(assets['vltt']).replace(',',''))
    assets['stk'] = float(str(assets['stk']).replace(',',''))
    assets['fnd'] = float(str(assets['fnd']).replace(',',''))
    assets['drv'] = float(str(assets['drv']).replace(',',''))
    assets['bnd'] = assets['vltt']-assets['stk']-assets['fnd']-assets['drv']
    print('vltt','stk','fnd','drv','bnd')
    print(assets['vltt'],assets['stk'],assets['fnd'],assets['drv'],assets['bnd'])
    if  not net_row or not asset_row: #如果没找到净资产或者总资产，就用证券投资合计/占总资产比例来计算 
        col_temp = ''
        if '证券投资合计:' in col:
            col_temp = '证券投资合计:'
        elif '证券投资合计：' in col:
            col_temp = '证券投资合计：'
        else:
            col_temp = df['科目代码'][1] #都找不到就去用第二行的市值和占净值比例算
        if not '市值占净值%' in column_name:#只有洛书出现这种情况,数量其实是比例 
            asset_row = col.index(col_temp)
            assets['val'] = df['市值'][asset_row]/df['数量'][asset_row]
        else:
            asset_row = col.index(col_temp)
            if isinstance(df['市值'][asset_row],str):
                assets['val'] = float(df['市值'][asset_row].replace(',',''))/float(df['市值占净值%'][asset_row].replace('%',''))*100
            else:assets['val'] = df['市值'][asset_row]/df['市值占净值%'][asset_row]
        assets['net'] = assets['val']
        assets['ratio'] = 1
    else:
        if isinstance(df['市值'][asset_row],str):
            assets['val'] = float(df['市值'][asset_row].replace(',',''))
            assets['net'] = float(df['市值'][net_row].replace(',',''))
        else:
            assets['val'] = df['市值'][asset_row]
            assets['net'] = df['市值'][net_row]
        assets['ratio'] = assets['val']/assets['net']
    output = []
    flg=0 #判断是否是新版估值表，如果是则债券代码需要修改
    ###########################################################################
    for row in range(0,df.index.size):
        code = df['科目代码'][row]
        if isinstance(code,str):
            code=code.replace(".","")
            code=code.replace(" ","")
            if code[-2:] in ["SH","SZ","CY"]:
                code=code[:-2]
                flg=1
        #过滤南方基金的表的第一列中的数字,并提取累计净值,实收资本
        if isinstance(code,str) :
            temp = code.translate(remove_digits)
            if temp == '实收资本':
                capital = df['市值'][row]
            if temp != '':
                code = temp
                if '累计单位净值' in code:
                    a_net_value = df['科目名称'][row]
                if '昨日单位净值' in code:
                    last_net_value = df['科目名称'][row]
                if '今日单位净值' in code:
                    today_net_value = df['科目名称'][row]
                if '期初单位净值' in code:
                    initial_net_value = df['科目名称'][row]
            if match('^[a-zA-Z]+$',temp):  #易方达的表里面MTN的代码有四个字母,判断有字母
                code = df['科目代码'][row]
       ###############################
        mktVal = df['市值'][row]
        if not mktVal or row == 0:#过滤
            continue
        mktVal = float(mktVal.replace(',','')) if isinstance(mktVal,str) else mktVal
        mktVal_r = mktVal/assets['val']#此处用的总资产
        qty = df['数量'][row] if '数量' in column_name else df['证券数量'][row]
        if isinstance(qty,str):
            if '%' in qty:
                qty = float(qty.replace('%',''))/100 
            else:qty = float(qty.replace(',',''))   
        name = df['科目名称'][row]
        valAdd = df['估值增值'][row] if '估值增值' in column_name else 0
        if '市价' in column_name:
            p = df['市价'][row]
        elif '行情收市价' in column_name: 
            p = df['行情收市价'][row]
        else: p = df['行情价格'][row] if '行情价格' in column_name else mktVal/qty 
        # deal with margin
        if code == AccountingSubjects.margin_future:
            a_subtype = '期货交易存出保证金'
            a_type = AssetClass.mapping[a_subtype]
            pos = [this_date, 'cash_CNY', 'cash_CNY', mktVal, 1, 
                   mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
            output.append(pos)        
        elif code == AccountingSubjects.margin_option:
            a_subtype  = '个股期权存出保证金'
            a_type = AssetClass.mapping[a_subtype]
            pos = [this_date, 'cash_CNY', 'cash_CNY', mktVal, 1, 
                   mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
            output.append(pos)
        # filter
        elif not isinstance(name,str):
            continue
        # deal with equities
        elif code[:4] == AccountingSubjects.equities:
            if np.isnan(qty) or np.isnan(mktVal):
                # no quantity or ratio or mktvalue
                continue
            if len(code) == AccountingSubjects.equities_digits:
                ticker = code[8::]
                if ticker[0] == 'H':
                    # deal with HK share
                    ticker = str(int(ticker[1::])) + '.HK'
                a_subtype = '股票'
                a_type = AssetClass.mapping[a_subtype]
                if np.isnan(mktVal_r): mktVal_r = 0
                pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
                output.append(pos)
        
        # deal with funds
        elif code[:4] == AccountingSubjects.funds:
            if np.isnan(qty) or np.isnan(mktVal):
                # no quantity or ratio or mktvalue
                continue
            if len(code) == AccountingSubjects.equities_digits:
                ticker = code[8::]
                a_subtype = '基金'
                a_type = AssetClass.mapping[a_subtype]
                if np.isnan(mktVal_r): mktVal_r = 0
                pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
                output.append(pos)
        #deal with trust
        elif code[:8] ==  AccountingSubjects.trust_fund:
            if len(code)>8:
                ticker = code[8::]
                a_subtype = '基金'
                a_type = AssetClass.mapping[a_subtype]
                if np.isnan(mktVal_r): mktVal_r = 0
                pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
                output.append(pos)
        
        # deal with bonds
        elif code[:4] == AccountingSubjects.bonds:
            if  np.isnan(mktVal):#########np.isnan(qty) or
                # no quantity or ratio or mktvalue
                continue
            if len(code) == AccountingSubjects.equities_digits or len(code) == 17 and code[0] in digits:
                
                a_subtype = map_name_to_bond_subtype(name)
                if flg==1:
                    ticker = map_code_to_bond_ticker1(code)
                else:
                    ticker = map_code_to_bond_ticker(code)
                a_type = '债券'
                if np.isnan(mktVal_r): mktVal_r = 0
                pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
                #print(ticker, name, qty, p, mktVal, mktVal_r, a_type, a_subtype)
                output.append(pos)
        elif code[:4] == AccountingSubjects.ABS:
            if len(code) == AccountingSubjects.ABS_digits:
                ticker = map_code_to_bond_ticker(code)
                a_subtype = 'ABS'
                a_type = '债券'
                pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
                output.append(pos)
        elif code[:4] == AccountingSubjects.bonds_intrest:
            if len(code) == AccountingSubjects.equities_digits or (len(code) == 17 and code[0] in digits):
                ticker = map_code_to_interest_ticker(code)
                a_type = '现金存款类' if code[:8] == AccountingSubjects.interest_bank else '利息'
                a_subtype = '现金存款类' if code[:8] == AccountingSubjects.interest_bank else '应收利息'
                pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
                output.append(pos)
        elif code[:4] == AccountingSubjects.dividend:
            if len(code) == AccountingSubjects.equities_digits:
                ticker = code[8::]
                a_type = '红利' 
                a_subtype = '现金存款类'
                pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
                output.append(pos)
        elif code[:4] in AccountingSubjects.buyback:
            if len(code) == AccountingSubjects.buyback_digits:
                ticker = name
                a_subtype = '正回购' if code[:4]=='2202' else '逆回购'
                a_type = '其他'
                qty = 1
                p = mktVal
                mktVal = mktVal 
                mktVal_r = mktVal_r
                pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
                output.append(pos)
        elif code[:4] in AccountingSubjects.money_all and len(code) == AccountingSubjects.money_all_digits:
            ticker = code
            a_subtype = '现金存款类'
            a_type = '现金存款类'
            qty = mktVal
            p = 1
            pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
            output.append(pos)
        elif code[:4] in AccountingSubjects.other_money and len(code) == AccountingSubjects.other_money_digits:
            ticker = code
            a_subtype = '现金存款类'
            a_type = '现金存款类'
            qty = mktVal
            p = 1
            pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
            output.append(pos)
        # deal with derivatives
        elif code[:4] in AccountingSubjects.all_derivatives:
            if np.isnan(qty) or len(code)<=10:
                # no quantity
                continue
            if not code.translate(remove_digits):#如果没有字母，则是委外资产管理计划
                ticker = code
                a_type = '其他'
                a_subtype = '资产管理计划'
            elif len(code) == AccountingSubjects.futures_digits:
                ticker = map_code_to_future_ticker(code)
                if '购' in name or '沽' in name:
                    a_subtype = '期权'
                elif 'TRS' in name:
                    a_subtype = 'SWAP'
                else:
                    if test_bond_future(ticker):
                        a_subtype = '国债期货'
                    elif test_index_future(ticker):
                        a_subtype = '股指期货'
                    elif '国债' in name:
                        a_subtype = '国债期货'
                    else:
                        a_subtype = '商品期货'
                a_type = AssetClass.mapping[a_subtype]
            pos = [this_date, ticker, name, qty, p, mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
            output.append(pos)
            
        # deal with options
        elif code[:6] == AccountingSubjects.options_mktValue:
            if np.isnan(mktVal):
                # no value
                continue
            a_subtype = '场内期权'
            a_type = '场内期权'
            pos = [this_date, '场内期权总市值', '场内期权总市值', mktVal, 1, 
                   mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
            output.append(pos)
            
        # deal with options margin
        elif code[:6] == AccountingSubjects.options_margins:
            if np.isnan(mktVal):
                # no market value
                continue
            a_subtype = '场内期权保证金'
            a_type = '场内期权保证金'
            pos = [this_date, 'cash_CNY', 'cash_CNY', mktVal, 1, 
                   mktVal, mktVal_r, a_type, a_subtype, valAdd,counter]
            output.append(pos) 
    otherdata = [os.path.split(file)[1],this_date,a_net_value,last_net_value,today_net_value,assets['net']/10000,capital,initial_net_value]
    asset4cmpr_data = [os.path.split(file)[1][:15],this_date,assets['stk'],assets['fnd'],assets['drv'],assets['bnd'],assets['vltt']]
    wb.close()
    return otherdata,asset4cmpr_data
    
def isNum(value):
	try:
		value + 1
	except TypeError:
		return False
	else:
		return True  
  
#以下是主程序
print("说明：请输入生成日报的类型并按回车：")
print("  直接处理底仓产品报告则输入0")
print("  若改为分步处理底仓报告，则先输入01，再输入02")
print("    执行其中第一步，即抓取并导出数据（到Mid文件夹里），则输入01")
print("    执行其中第二步，即读取Mid文件夹中的数据并生成报告，则输入02")
print("  处理母基金报告则输入1")
choice = str(input("请输入0或01或02或1      "))
if choice == '0':
    report_type = '底仓产品'
if choice == '1':
    report_type = '母基金'
if choice == '01':
    report_type = '底仓产品第一步（抓取并导出数据）'
if choice == '02':
    report_type = '底仓产品第二步（读取并处理数据）'
# report_type = '底仓产品' if choice == '0' else '母基金'
# report_type = '底仓产品分步' if choice == '2' else '底仓产品'
path_now = os.path.abspath('.')
path_reports = []
if report_type == '母基金':
    path_mom=path_now+"\Input\母基金估值表存放" 
    moms = glob(path_mom + "\*.xls")#取得所有母基金xls估值表的路径
    moms.extend(glob(path_mom + "\*.xlsx"))#取得所有母基金xlsx估值表的路径
    count = 0
    mom_holding_path=path_now+'\母基金持有份额.xlsx'
    
    for mom_num in moms:
        mom_holding=pd.DataFrame()
        MOM_output,MOM_otherdata = export_positions(mom_num, 0)
        del(MOM_otherdata[5])#因为估值表中没有净资产，这里出现nan，删掉它
        mom_num = os.path.split(mom_num)[1]
        mom_num = mom_num.split(".")[0]
        file_paths = glob(path_mom + '\\' +mom_num+"\*.xlsx")#获取每个母基金文件夹下的子估值表
        file_paths.extend(glob(path_mom + '\\' +mom_num+"\*.xls"))
        new_excel_name = "日报模板_母基金.xlsx" #模板的名字
        all_positions_history = []
        all_other_data_history = []
        net_values_file = path_now+"\Input\历史净值数据存放\MOM历史净值.xlsx"
        i = 0
        for file_path in file_paths:
            output,otherdata = export_positions(file_path, i)
            all_positions_history.extend(output)
            all_other_data_history.append(otherdata)
            df=pd.read_excel(mom_holding_path,sheetname=mom_num[0:4]) #edit
            mom_holding=mom_holding.append(df[[1]][df.子基金==otherdata[0][0:6]])
            i += 1
            print('正在处理'+file_path)
            print('请确保所有估值表日期一致，当前估值表的日期是{}'.format(otherdata[1]))
        if all_positions_history !=[]:#如果不为空，再进行下面操作
            output_excel = xw.Book(new_excel_name)
            all_positions_df = pd.DataFrame(all_positions_history,columns = ('Date', 'Code', 'Name', 'Quantity', 'Price', 
                       'Turnover', 'Ratio', 'Asset_Type', 'Asset_Subtype', 'ValueAdd','from which fund'))
            all_positions_df = all_positions_df.drop('Date',1)
            output_excel.sheets['position'].range('A1:K3000').clear_contents()
            output_excel.sheets['position'].range('A1:K3000').value = all_positions_df
            all_other_data_history_df = pd.DataFrame(all_other_data_history)
            all_other_data_history_df.columns = ['估值表','估值表日期','累计单位净值','昨日单位净值','今日单位净值','净资产（万元）','实收资本','初始净值']
            output_excel.sheets['position'].range('O9:V18').clear_contents()
            output_excel.sheets['position'].range('O9:V18').value = all_other_data_history_df
            output_excel.sheets['position'].range('Q4:W5').clear_contents()
            output_excel.sheets['position'].range('Q4:W5').value = MOM_otherdata
            output_excel.sheets['position'].range('Y10').value=mom_holding.values
            nv_excel = xw.Book(net_values_file)
            nv_sht = nv_excel.sheets[mom_num]
            nv_data = nv_sht.range('A2:C2').expand('down').value
            if nv_data[-1][0] < MOM_otherdata[1]:
                nv_data.append([MOM_otherdata[1],'',MOM_otherdata[4]])
            nv_sht.range('A2:C2').value = nv_data    #把净值数据更新到原excel中 
            nv_excel.save()
            nv_excel.close()
            output_excel.sheets('净值数据').range('A2:C2').expand('down').clear_contents()
            output_excel.sheets('净值数据').range('A2:C2').value = nv_data
            output_excel.sheets('表头').range('A7').value = mom_num
            
            output_excel.save(path_now+"\Output\明细\{}.xlsx".format(mom_num))
            new_wb = xw.Book(path_now+"\Output\明细\{}.xlsx".format(mom_num))
            new_wb.close()
    wb_merge = xw.Book(path_now + "\日报模板_母基金汇总.xlsm")
    for mom_num in moms:
        mom_num = os.path.split(mom_num)[1]
        mom_num = mom_num.split(".")[0]
        new_wb = xw.Book(path_now+"\Output\明细\{}.xlsx".format(mom_num))
        VBA_merge = wb_merge.macro('report_merge')
        VBA_merge(mom_num,count)
#            VBA_photo = wb_merge.macro('Chart_to_photo')
#            VBA_photo()
#            VBA_to_value = wb_merge.macro('to_value')
#            VBA_to_value()
        wb_merge.sheets('表头').range('H4:H4').value = new_wb.sheets('表头').range('H4:H4').value
        count += 1      
    wb_merge.save(path_now+"\Output\MOM日报.xlsm")
        
elif report_type == '底仓产品':
    file_paths = glob(path_now +"\Input\子基金估值表存放\*.xlsx")
    file_paths.extend(glob(path_now +"\Input\子基金估值表存放\*.xls"))
    new_excel_name = "日报模板_底仓.xlsx"#模板的名字
    wb_merge = xw.Book("日报模板_底仓汇总.xlsm")#汇总的模板的名字
    count = 0
    delete_chart_flag = 0;
    for file_path in file_paths:
        all_positions_history = []
        all_other_data_history = []      
        output,otherdata = export_positions(file_path, 1)
        net_values_file = path_now+"\Input\历史净值数据存放\子基金历史净值.xlsx"
        all_positions_history.extend(output)
        all_other_data_history.extend(otherdata)
        print('正在处理'+file_path)
        print('请确保所有估值表日期一致，当前估值表的日期是{}'.format(otherdata[1]))
        if all_positions_history !=[]:#如果不为空，再进行下面操作
            output_excel = xw.Book(new_excel_name)
            all_positions_df = pd.DataFrame(all_positions_history,columns = ('Date', 'Code', 'Name', 'Quantity', 'Price', 
                       'Turnover', 'Ratio', 'Asset_Type', 'Asset_Subtype', 'ValueAdd','from which fund'))
            all_positions_df = all_positions_df.drop('Date',1)
            all_positions_df = all_positions_df.drop('from which fund',1)
            output_excel.sheets['position'].range('A1:J3000').clear_contents()
            output_excel.sheets['position'].range('A1:J3000').value = all_positions_df
            all_other_data_history_df = pd.DataFrame(all_other_data_history)
            all_other_data_history_df.index = ['估值表','估值表日期','累计单位净值','昨日单位净值','今日单位净值','净资产（万元）','实收资本','期初净值']
            output_excel.sheets['position'].range('N10:O17').clear_contents()
            output_excel.sheets['position'].range('N10:N17').value = all_other_data_history_df
            #提出债券信息和利息 
            bond_df = all_positions_df[all_positions_df['Asset_Type'] == '债券'] 
            bond_df = bond_df[['Code','Name','Turnover']]
            interest_df = all_positions_df[all_positions_df['Asset_Type'] == '利息'] #提出利息信息
            interest_df = interest_df[['Code','Turnover']]
            bond_all_df = pd.merge(bond_df,interest_df,on = 'Code',how = 'left')
            bond_all_df = bond_all_df.drop('Code',1)
            bond_all_df.columns = ['简称','持有市值','应收利息']
            output_excel.sheets['债券投资'].range('A1:D1000').clear_contents()
            output_excel.sheets['债券投资'].range('A1:D1000').value = bond_all_df
            #导出股票
            equity_df = all_positions_df[all_positions_df['Asset_Type'] == '股票']
            equity_df_1 = equity_df[['Turnover','Code']]
            equity_df_1.index = equity_df['Name']
            equity_df_1.columns = ['持有市值','代码']
            output_excel.sheets['MOM行业偏离度'].range('A1:C3000').clear_contents()
            output_excel.sheets['MOM行业偏离度'].range('A1:C3000').value = equity_df_1
            output_excel.sheets['MOM市值集中度'].range('A1:C3000').clear_contents()
            output_excel.sheets['MOM市值集中度'].range('A1:C3000').value = equity_df_1
            delete_chart_flag = 1 if equity_df_1.size<2 else 0
            nv_excel = xw.Book(net_values_file)
            nv_sht = nv_excel.sheets[otherdata[0][:6]]
            nv_data = nv_sht.range('A2:C2').expand('down').value
            if nv_data[-1][0] < otherdata[1]:
                nv_data.append([otherdata[1],'',otherdata[4]])
            nv_sht.range('A2:C2').value = nv_data   #把净值数据更新到原excel中 
            output_excel.sheets('净值数据').range('A2:C2').expand('down').clear_contents()
            output_excel.sheets('净值数据').range('A2:C2').value = nv_data
            new_report_name = path_now+"\Output\明细\{}report.xlsx".format(otherdata[0][:15])
            output_excel.save(new_report_name)
            new_wb = xw.Book(new_report_name)
            VBA_merge = wb_merge.macro('report_merge')
            VBA_merge("{}report".format(otherdata[0][:15]),count,delete_chart_flag)
            VBA_photo = wb_merge.macro('Chart_to_photo')
            VBA_photo()
            VBA_to_value = wb_merge.macro('to_value')
            VBA_to_value()
            wb_merge.sheets('表头').range('A7:J7').offset(count,0).value = new_wb.sheets('表头').range('A7:J7').value
            wb_merge.sheets('表头').range('N7:S7').offset(count,0).value = new_wb.sheets('表头').range('D20:I20').value
            wb_merge.sheets('表头').range('H4:H4').value = new_wb.sheets('表头').range('H4:H4').value
            wb_merge.sheets('表头').range('AF7:AJ7').offset(count,0).value = new_wb.sheets('表头').range('D42:H42').value
            new_wb.close()  
            count += 1
            nv_excel.save()
            nv_excel.close()
    wb_merge.save(path_now+"\Output\底仓日报汇总.xlsm")
    #elif report_type == '底仓产品New2':
    file_paths = glob(path_now +"\Input\子基金估值表存放\*.xlsx")
    file_paths.extend(glob(path_now +"\Input\子基金估值表存放\*.xls"))
    new_excel_name = "日报模板_底仓.xlsx"#模板的名字
    wb_merge = xw.Book("日报模板_底仓汇总.xlsm")#汇总的模板的名字
    wb_New2 = xw.Book(path_now+"\Output\底仓日报汇总.xlsm")
    wb_prf_rate_detail = xw.Book(path_now+"\Output\大类资产盈亏导出.xlsx")
    count = 0
    delete_chart_flag = 0;
    output_data_history = []
    for file_path in file_paths:
        print('正在处理'+file_path)
        output = export_positions_New2(file_path, 1)
        net_values_file = path_now+"\Input\历史净值数据存放\子基金历史净值.xlsx"
        output_data_history.extend(output)
    file_paths = glob(path_now +"\Input\子基金估值表存放（昨日）\*.xlsx")
    file_paths.extend(glob(path_now +"\Input\子基金估值表存放（昨日）\*.xls"))    
    for file_path in file_paths:
        print('正在处理'+file_path)
        output = export_positions_New2(file_path, 1)
        net_values_file = path_now+"\Input\历史净值数据存放\子基金历史净值.xlsx"
        output_data_history.extend(output)
    position2d_df = pd.DataFrame(output_data_history,columns = ('Fundname', 'Date', 'Code', 'Name', 'Quantity', 'Price', 
                       'Turnover', 'Ratio', 'Asset_Type', 'Asset_Subtype', 'ValueAdd','from which fund','NV'))
    bb = []
    cc = []
    aa = list(set(list(position2d_df['Fundname'])))
    for e in aa:
        if e[0] in '0123456789QWERTYUIOPASDFGHJKLZXCVBNMqwertyuiopasdfghjklzxcvbnm': 
            bb.append(e)
        else:
            cc.append(e)
    aa = []
    aa.extend(sorted(cc))
    aa.extend(sorted(bb))
    for fundname in aa:
        fundk_df = position2d_df[position2d_df['Fundname'] == fundname]
        fundk_stk_df = fundk_df[fundk_df['Asset_Type'] == '股票']
        fundk_fnd_df = fundk_df[fundk_df['Asset_Type'] == '基金']
        fundk_bnd_df = fundk_df[fundk_df['Asset_Type'] == '债券']
        fundk_drv_df = fundk_df[fundk_df['Asset_Type'] == '衍生品']
        # 处理股票
        day2_stk = list(set(list(fundk_stk_df['Date'])))
        if len(day2_stk) == 2:
            print('处理股票')
            stk_label = 1
            thisday = day2_stk[0]
            lastday = day2_stk[1]
            stk_group=fundk_stk_df.groupby('Date')
            stk_df_thisday = stk_group.get_group(thisday)
            stk_df_lastday = stk_group.get_group(lastday)
            print('取交集前',len(list(stk_df_lastday['Name'])),len(list(stk_df_thisday['Name'])))
            #stk_df_thisday与stk_df_lastday取共有项
            stk_df_thisday = stk_df_thisday[stk_df_thisday.Name.isin(list(stk_df_lastday['Name']))]
            stk_df_lastday = stk_df_lastday[stk_df_lastday.Name.isin(list(stk_df_thisday['Name']))]
            print('取交集后',len(list(stk_df_lastday['Name'])),len(list(stk_df_thisday['Name'])))
            stk_df_thisday = stk_df_thisday.reset_index(drop=True)
            stk_df_lastday = stk_df_lastday.reset_index(drop=True)
            #for aaa in stk_df_thisday['Code']:
                #print(aaa,type(aaa))
            for i in range(0,len(stk_df_thisday['Code'])):
                stk_df_thisday.loc[i,'Code'] = str(stk_df_thisday.loc[i,'Code'])
            stk_df_thisday['Code'] = sorted(stk_df_thisday['Code'])
            for i in range(0,len(stk_df_lastday['Code'])):
                stk_df_lastday.loc[i,'Code'] = str(stk_df_lastday.loc[i,'Code'])
            stk_df_lastday['Code'] = sorted(stk_df_lastday['Code'])
            stk_df_thisday = stk_df_thisday.sort_values(by='Code')
            stk_df_thisday = stk_df_thisday.reset_index(drop=True)
            stk_df_lastday = stk_df_lastday.sort_values(by='Code')
            stk_df_lastday = stk_df_lastday.reset_index(drop=True)
            stk_df_lastday.Price = stk_df_thisday.Price - stk_df_lastday.Price
            fnl_stk_df = stk_df_lastday
            fnl_stk_df.rename(columns={'Price':'DeltaPrice'}, inplace = True)
            fnl_stk_df['Profit'] = fnl_stk_df['DeltaPrice'] * fnl_stk_df['Quantity']
            fund_name = fnl_stk_df.loc[0,:]['Fundname']
            fund_NV = fnl_stk_df.loc[0,:]['NV']
            stk_sum_profit = sum(list(fnl_stk_df['Profit']))
            stk_profit_rate = stk_sum_profit / fnl_stk_df.loc[0,:]['NV']
            print('fund_name','fund_NV','stk_sum_profit','stk_profit_rate')
            print(fund_name, fund_NV, stk_sum_profit, stk_profit_rate)
        else:
            stk_label = 0
        # 处理基金
        day2_fnd = list(set(list(fundk_fnd_df['Date'])))
        if len(day2_fnd) == 2:
            print('处理基金')
            fnd_label = 1
            thisday = day2_fnd[0]
            lastday = day2_fnd[1]
            fnd_group=fundk_fnd_df.groupby('Date')
            fnd_df_thisday = fnd_group.get_group(thisday)
            fnd_df_lastday = fnd_group.get_group(lastday)
            print('取交集前',len(list(fnd_df_lastday['Name'])),len(list(fnd_df_thisday['Name'])))
            #fnd_df_thisday与fnd_df_lastday取共有项
            fnd_df_thisday = fnd_df_thisday[fnd_df_thisday.Name.isin(list(fnd_df_lastday['Name']))]
            fnd_df_lastday = fnd_df_lastday[fnd_df_lastday.Name.isin(list(fnd_df_thisday['Name']))]
            print('取交集后',len(list(fnd_df_lastday['Name'])),len(list(fnd_df_thisday['Name'])))
            fnd_df_thisday = fnd_df_thisday.sort_values(by='Code')
            fnd_df_thisday = fnd_df_thisday.reset_index(drop=True)
            fnd_df_lastday = fnd_df_lastday.sort_values(by='Code')
            fnd_df_lastday = fnd_df_lastday.reset_index(drop=True)
            fnd_df_lastday.Price = fnd_df_thisday.Price - fnd_df_lastday.Price
            fnl_fnd_df = fnd_df_lastday
            fnl_fnd_df.rename(columns={'Price':'DeltaPrice'}, inplace = True)
            fnl_fnd_df['Profit'] = fnl_fnd_df['DeltaPrice'] * fnl_fnd_df['Quantity']
            fund_name = fnl_fnd_df.loc[0,:]['Fundname']
            fund_NV = fnl_fnd_df.loc[0,:]['NV']
            fnd_sum_profit = sum(list(fnl_fnd_df['Profit']))
            fnd_profit_rate = fnd_sum_profit / fnl_fnd_df.loc[0,:]['NV']
            print('fund_name','fund_NV','fnd_sum_profit','fnd_profit_rate')
            print(fund_name, fund_NV, fnd_sum_profit, fnd_profit_rate)
        else:
            fnd_label = 0
        # 处理债券
        day2_bnd = list(set(list(fundk_bnd_df['Date'])))
        if len(day2_bnd) == 2:
            print('处理债券')
            bnd_label = 1
            thisday = day2_bnd[0]
            lastday = day2_bnd[1]
            bnd_group=fundk_bnd_df.groupby('Date')
            bnd_df_thisday = bnd_group.get_group(thisday)
            bnd_df_lastday = bnd_group.get_group(lastday)
            print('取交集前',len(list(bnd_df_lastday['Name'])),len(list(bnd_df_thisday['Name'])))
            #bnd_df_thisday与bnd_df_lastday取共有项
            bnd_df_thisday = bnd_df_thisday[bnd_df_thisday.Name.isin(list(bnd_df_lastday['Name']))]
            bnd_df_lastday = bnd_df_lastday[bnd_df_lastday.Name.isin(list(bnd_df_thisday['Name']))]
            print('取交集后',len(list(bnd_df_lastday['Name'])),len(list(bnd_df_thisday['Name'])))
            bnd_df_thisday = bnd_df_thisday.sort_values(by='Code')
            bnd_df_thisday = bnd_df_thisday.reset_index(drop=True)
            bnd_df_lastday = bnd_df_lastday.sort_values(by='Code')
            bnd_df_lastday = bnd_df_lastday.reset_index(drop=True)
            bnd_df_lastday.Price = bnd_df_thisday.Price - bnd_df_lastday.Price
            fnl_bnd_df = bnd_df_lastday
            fnl_bnd_df = fnl_bnd_df.dropna()
            fnl_bnd_df = fnl_bnd_df.reset_index(drop=True)
            fnl_bnd_df.rename(columns={'Price':'DeltaPrice'}, inplace = True)
            fnl_bnd_df['Profit'] = fnl_bnd_df['DeltaPrice'] * fnl_bnd_df['Quantity']
            fund_name = fnl_bnd_df.loc[0,:]['Fundname']
            fund_NV = fnl_bnd_df.loc[0,:]['NV']
            bnd_sum_profit = sum(list(fnl_bnd_df['Profit']))
            bnd_profit_rate = bnd_sum_profit / fnl_bnd_df.loc[0,:]['NV']
            print('fund_name','fund_NV','bnd_sum_profit','bnd_profit_rate')
            print(fund_name, fund_NV, bnd_sum_profit, bnd_profit_rate)            
        else:
            bnd_label = 0
        # 处理衍生品   
        day2_drv = list(set(list(fundk_drv_df['Date'])))
        if len(day2_drv) == 2:
            print('处理衍生品')
            drv_label = 1
            thisday = day2_drv[0]
            lastday = day2_drv[1]
            drv_group=fundk_drv_df.groupby('Date')
            drv_df_thisday = drv_group.get_group(thisday)
            drv_df_lastday = drv_group.get_group(lastday)
            print('取交集前',len(list(drv_df_lastday['Name'])),len(list(drv_df_thisday['Name'])))
            #drv_df_thisday与drv_df_lastday取共有项
            drv_df_thisday = drv_df_thisday[drv_df_thisday.Name.isin(list(drv_df_lastday['Name']))]
            drv_df_lastday = drv_df_lastday[drv_df_lastday.Name.isin(list(drv_df_thisday['Name']))]
            print('取交集后',len(list(drv_df_lastday['Name'])),len(list(drv_df_thisday['Name'])))
            drv_df_thisday = drv_df_thisday.sort_values(by='Code')
            drv_df_thisday = drv_df_thisday.reset_index(drop=True)
            drv_df_lastday = drv_df_lastday.sort_values(by='Code')
            drv_df_lastday = drv_df_lastday.reset_index(drop=True)
            drv_df_lastday.Price = drv_df_thisday.Price - drv_df_lastday.Price
            fnl_drv_df = drv_df_lastday
            fnl_drv_df.rename(columns={'Price':'DeltaPrice'}, inplace = True)
            fnl_drv_df['Profit'] = fnl_drv_df['DeltaPrice'] * fnl_drv_df['Quantity']
            fund_name = fnl_drv_df.loc[0,:]['Fundname']
            fund_NV = fnl_drv_df.loc[0,:]['NV']
            drv_sum_profit = sum(list(fnl_drv_df['Profit']))
            drv_profit_rate = drv_sum_profit / fnl_drv_df.loc[0,:]['NV']
            print('fund_name','fund_NV','drv_sum_profit','drv_profit_rate')
            print(fund_name, fund_NV, drv_sum_profit, drv_profit_rate)            
        else:
            drv_label = 0
        wb_New2.sheets('表头').range('U7:U7').offset(count,0).value = fund_name[:6]
        wb_New2.sheets('表头').range('W7:W7').offset(count,0).value = stk_profit_rate if stk_label == 1 else '/'
        wb_New2.sheets('表头').range('X7:X7').offset(count,0).value = bnd_profit_rate if bnd_label == 1 else '/'
        wb_New2.sheets('表头').range('Y7:Y7').offset(count,0).value = fnd_profit_rate if fnd_label == 1 else '/'
        wb_New2.sheets('表头').range('Z7:Z7').offset(count,0).value = drv_profit_rate if drv_label == 1 else '/'
        wb_prf_rate_detail.sheets('Sheet1').range('A2:A2').offset(count,0).value = fund_name[:6]
        wb_prf_rate_detail.sheets('Sheet1').range('C2:C2').offset(count,0).value = fund_NV
        wb_prf_rate_detail.sheets('Sheet1').range('D2:D2').offset(count,0).value = stk_sum_profit if stk_label == 1 else '/'
        wb_prf_rate_detail.sheets('Sheet1').range('E2:E2').offset(count,0).value = bnd_sum_profit if bnd_label == 1 else '/'
        wb_prf_rate_detail.sheets('Sheet1').range('F2:F2').offset(count,0).value = fnd_sum_profit if fnd_label == 1 else '/'
        wb_prf_rate_detail.sheets('Sheet1').range('G2:G2').offset(count,0).value = drv_sum_profit if drv_label == 1 else '/'  
        wb_prf_rate_detail.sheets('Sheet1').range('H2:H2').offset(count,0).value = stk_profit_rate if stk_label == 1 else '/'
        wb_prf_rate_detail.sheets('Sheet1').range('I2:I2').offset(count,0).value = bnd_profit_rate if bnd_label == 1 else '/'
        wb_prf_rate_detail.sheets('Sheet1').range('J2:J2').offset(count,0).value = fnd_profit_rate if fnd_label == 1 else '/'
        wb_prf_rate_detail.sheets('Sheet1').range('K2:K2').offset(count,0).value = drv_profit_rate if drv_label == 1 else '/'        
        count += 1
    wb_New2.save(path_now+"\Output\底仓日报汇总.xlsm")
    wb_prf_rate_detail.save(path_now+"\Output\大类资产盈亏导出.xlsx")

            
elif report_type == '底仓产品第一步（抓取并导出数据）':
    file_paths = glob(path_now +"\Input\子基金估值表存放\*.xlsx")
    file_paths.extend(glob(path_now +"\Input\子基金估值表存放\*.xls"))
    dpl_file_paths = file_paths
    mid_file_paths = []
    for path in dpl_file_paths:
        path = path.replace('input','mid')
        path = path.replace('Input','Mid')
        path2 = path.replace('估值表','估值sheet')
        mid_file_paths.append(path2)
    mid_file_path0 = path_now +"\Mid\今日子基金估值表导出数据.xlsx"
    mid_excel0 = xw.Book(mid_file_path0)
    new_excel_name = "日报模板_底仓.xlsx"#模板的名字
    wb_merge = xw.Book("日报模板_底仓汇总.xlsm")#汇总的模板的名字
    count = 0
    delete_chart_flag = 0;
    # 开始导出固定格式持仓表及资产表
    for file_path in file_paths:
        all_positions_history = []
        all_other_data_history = []      
        output,otherdata = export_positions(file_path, 1)
        net_values_file = path_now+"\Input\历史净值数据存放\子基金历史净值.xlsx"
        all_positions_history.extend(output)
        all_other_data_history.extend(otherdata)
        sheet_name_holding = file_path
        aaaa = path_now+"\Input\子基金估值表存放\w"
        aaaa = aaaa[:-1]
        sheet_name_holding = sheet_name_holding.replace(aaaa,'')
        fundname = sheet_name_holding[:6]
        #file_path = file_path.replace('input','mid')
        #file_path = file_path.replace('Input','Mid')
        #file_path2 = file_path.replace('估值表','估值sheet')
        #wb_mid = xw.Book(file_path2)
        #wb_mid.sheets['Sheet1'].range('A2:Z10000').clear_contents()
        #wb_mid.sheets['Sheet1'].range('A2:Z10000').value = all_positions_history
        #wb_mid.sheets['Sheet2'].range('A2:Z10000').clear_contents()
        #countlt = 0
        #for data in all_other_data_history: 
            #wb_mid.sheets['Sheet2'].range('A2').offset(countlt,0).value = data
            #countlt = countlt + 1
        #wb_mid.save(file_path2)
        mid_excel0.sheets[fundname+"持仓"].range('A1:Z10000').clear_contents()
        all_positions_history_df = pd.DataFrame(all_positions_history,columns = ('Date', 'Code', 'Name', 'Quantity', 'Price', 
                       'Turnover', 'Ratio', 'Asset_Type', 'Asset_Subtype', 'ValueAdd','from which fund'))
        all_positions_history_df = all_positions_history_df.drop('from which fund',1)
        all_positions_history_df = all_positions_history_df.drop('ValueAdd',1)
        mid_excel0.sheets[fundname+"持仓"].range('A1:Z10000').value = all_positions_history_df
        countlt = 0
        for data in all_other_data_history: 
            mid_excel0.sheets[fundname+"资产"].range('A2').offset(countlt,0).value = data
            countlt = countlt + 1
        mid_excel0.save(mid_file_path0)
    #elif report_type == '底仓产品New2':
    file_paths = glob(path_now +"\Input\子基金估值表存放\*.xlsx")
    file_paths.extend(glob(path_now +"\Input\子基金估值表存放\*.xls"))
    mid_file_path = path_now +"\Mid\昨日+今日持仓汇总.xlsx"
    new_excel_name = "日报模板_底仓.xlsx"#模板的名字
    wb_merge = xw.Book("日报模板_底仓汇总.xlsm")#汇总的模板的名字
    wb_New2 = xw.Book(path_now+"\Output\底仓日报汇总.xlsm")
    wb_prf_rate_detail = xw.Book(path_now+"\Output\大类资产盈亏导出.xlsx")
    mid_excel = xw.Book(mid_file_path)
    count = 0
    delete_chart_flag = 0;
    # 开始导出昨日+今日持仓表
    output_data_history = []
    for file_path in file_paths:
        print('正在处理'+file_path)
        output = export_positions_New2(file_path, 1)
        net_values_file = path_now+"\Input\历史净值数据存放\子基金历史净值.xlsx"
        output_data_history.extend(output)
    file_paths = glob(path_now +"\Input\子基金估值表存放（昨日）\*.xlsx")
    file_paths.extend(glob(path_now +"\Input\子基金估值表存放（昨日）\*.xls"))    
    for file_path in file_paths:
        print('正在处理'+file_path)
        output = export_positions_New2(file_path, 1)
        net_values_file = path_now+"\Input\历史净值数据存放\子基金历史净值.xlsx"
        output_data_history.extend(output)
    position2d_df = pd.DataFrame(output_data_history,columns = ('Fundname', 'Date', 'Code', 'Name', 'Quantity', 'Price', 
                       'Turnover', 'Ratio', 'Asset_Type', 'Asset_Subtype', 'ValueAdd','from which fund','NV'))
    position2d_df = position2d_df.drop('from which fund',1)
    position2d_df = position2d_df.drop('ValueAdd',1)
    mid_excel.sheets['Sheet1'].range('A1:Z10000').clear_contents()
    mid_excel.sheets['Sheet1'].range('A1:Z10000').value = position2d_df   
    mid_excel.save(mid_file_path)
    wb_merge.close()
    wb_prf_rate_detail.close()
    wb_New2.close()
    

elif report_type == '底仓产品第二步（读取并处理数据）':
    file_paths = glob(path_now +"\Input\子基金估值表存放\*.xlsx")
    file_paths.extend(glob(path_now +"\Input\子基金估值表存放\*.xls"))
    dpl_file_paths = file_paths
    mid_file_paths = []
    for path in dpl_file_paths:
        path = path.replace('input','mid')
        path = path.replace('Input','Mid')
        path2 = path.replace('估值表','估值sheet')
        mid_file_paths.append(path2)
    mid_file_path0 = path_now +"\Mid\今日子基金估值表导出数据.xlsx"
    mid_excel0 = xw.Book(mid_file_path0)
    new_excel_name = "日报模板_底仓.xlsx"#模板的名字
    wb_merge = xw.Book("日报模板_底仓汇总.xlsm")#汇总的模板的名字
    count = 0
    delete_chart_flag = 0;
    # 开始处理导出表
    for file_path in file_paths:
        output,otherdata = export_positions(file_path, 1)
        sheet_name_holding = file_path
        aaaa = path_now+"\Input\子基金估值表存放\w"
        aaaa = aaaa[:-1]
        sheet_name_holding = sheet_name_holding.replace(aaaa,'')
        fundname = sheet_name_holding[:6]        
        #file_path = file_path.replace('input','mid')
        #file_path = file_path.replace('Input','Mid')
        #file_path2 = file_path.replace('估值表','估值sheet')
        #all_positions_df = pd.read_excel(file_path2,'Sheet1')
        #all_other_data_history_df = pd.read_excel(file_path2,'Sheet2')
        all_positions_df = pd.read_excel(mid_file_path0,fundname+"持仓")
        all_other_data_history_df = pd.read_excel(mid_file_path0,fundname+"资产")       
        all_other_data_history_df.index = ['估值表','估值表日期','累计单位净值','昨日单位净值','今日单位净值','净资产（万元）','实收资本','期初净值']
        net_values_file = path_now+"\Input\历史净值数据存放\子基金历史净值.xlsx"
        print('正在处理'+file_path)
        print('请确保所有估值表日期一致，当前估值表的日期是{}'.format(otherdata[1]))
        if 1 == 1:#如果不为空，再进行下面操作
            output_excel = xw.Book(new_excel_name)
            all_positions_df = all_positions_df.drop('Date',1)
            output_excel.sheets['position'].range('A1:J3000').clear_contents()
            output_excel.sheets['position'].range('A1:J3000').value = all_positions_df
            output_excel.sheets['position'].range('N10:O17').clear_contents()
            output_excel.sheets['position'].range('N10:O18').value = all_other_data_history_df
            #提出债券信息和利息 
            bond_df = all_positions_df[all_positions_df['Asset_Type'] == '债券'] 
            bond_df = bond_df[['Code','Name','Turnover']]
            interest_df = all_positions_df[all_positions_df['Asset_Type'] == '利息'] #提出利息信息
            interest_df = interest_df[['Code','Turnover']]
            bond_all_df = pd.merge(bond_df,interest_df,on = 'Code',how = 'left')
            bond_all_df = bond_all_df.drop('Code',1)
            bond_all_df.columns = ['简称','持有市值','应收利息']
            output_excel.sheets['债券投资'].range('A1:D1000').clear_contents()
            output_excel.sheets['债券投资'].range('A1:D1000').value = bond_all_df
            #导出股票
            equity_df = all_positions_df[all_positions_df['Asset_Type'] == '股票']
            equity_df_1 = equity_df[['Turnover','Code']]
            equity_df_1.index = equity_df['Name']
            equity_df_1.columns = ['持有市值','代码']
            output_excel.sheets['MOM行业偏离度'].range('A1:C3000').clear_contents()
            output_excel.sheets['MOM行业偏离度'].range('A1:C3000').value = equity_df_1
            output_excel.sheets['MOM市值集中度'].range('A1:C3000').clear_contents()
            output_excel.sheets['MOM市值集中度'].range('A1:C3000').value = equity_df_1
            delete_chart_flag = 1 if equity_df_1.size<2 else 0
            nv_excel = xw.Book(net_values_file)
            nv_sht = nv_excel.sheets[otherdata[0][:6]]
            nv_data = nv_sht.range('A2:C2').expand('down').value
            if nv_data[-1][0] < otherdata[1]:
                nv_data.append([otherdata[1],'',otherdata[4]])
            nv_sht.range('A2:C2').value = nv_data   #把净值数据更新到原excel中 
            output_excel.sheets('净值数据').range('A2:C2').expand('down').clear_contents()
            output_excel.sheets('净值数据').range('A2:C2').value = nv_data
            new_report_name = path_now+"\Output\明细\{}report.xlsx".format(otherdata[0][:15])
            output_excel.save(new_report_name)
            new_wb = xw.Book(new_report_name)
            VBA_merge = wb_merge.macro('report_merge')
            VBA_merge("{}report".format(otherdata[0][:15]),count,delete_chart_flag)
            VBA_photo = wb_merge.macro('Chart_to_photo')
            VBA_photo()
            VBA_to_value = wb_merge.macro('to_value')
            VBA_to_value()
            wb_merge.sheets('表头').range('A7:J7').offset(count,0).value = new_wb.sheets('表头').range('A7:J7').value
            wb_merge.sheets('表头').range('N7:S7').offset(count,0).value = new_wb.sheets('表头').range('D20:I20').value
            wb_merge.sheets('表头').range('H4:H4').value = new_wb.sheets('表头').range('H4:H4').value
            wb_merge.sheets('表头').range('AF7:AJ7').offset(count,0).value = new_wb.sheets('表头').range('D42:H42').value
            new_wb.close()  
            count += 1
            nv_excel.save()
            nv_excel.close()
    wb_merge.save(path_now+"\Output\底仓日报汇总.xlsm")
    #elif report_type == '底仓产品New2':
    file_paths = glob(path_now +"\Input\子基金估值表存放\*.xlsx")
    file_paths.extend(glob(path_now +"\Input\子基金估值表存放\*.xls"))
    mid_file_path = path_now +"\Mid\昨日+今日持仓汇总.xlsx"
    new_excel_name = "日报模板_底仓.xlsx"#模板的名字
    wb_merge = xw.Book("日报模板_底仓汇总.xlsm")#汇总的模板的名字
    wb_New2 = xw.Book(path_now+"\Output\底仓日报汇总.xlsm")
    wb_prf_rate_detail = xw.Book(path_now+"\Output\大类资产盈亏导出.xlsx")
    mid_excel = xw.Book(mid_file_path)
    count = 0
    delete_chart_flag = 0;
    # 开始处理导出表
    position2d_df = pd.read_excel(mid_file_path)
    bb = []
    cc = []
    aa = list(set(list(position2d_df['Fundname'])))
    for e in aa:
        if e[0] in '0123456789QWERTYUIOPASDFGHJKLZXCVBNMqwertyuiopasdfghjklzxcvbnm': 
            bb.append(e)
        else:
            cc.append(e)
    aa = []
    aa.extend(sorted(cc))
    aa.extend(sorted(bb))
    for fundname in aa:
        fundk_df = position2d_df[position2d_df['Fundname'] == fundname]
        fundk_stk_df = fundk_df[fundk_df['Asset_Type'] == '股票']
        fundk_fnd_df = fundk_df[fundk_df['Asset_Type'] == '基金']
        fundk_bnd_df = fundk_df[fundk_df['Asset_Type'] == '债券']
        fundk_drv_df = fundk_df[fundk_df['Asset_Type'] == '衍生品']
        # 处理股票
        day2_stk = list(set(list(fundk_stk_df['Date'])))
        if len(day2_stk) == 2:
            print('处理股票')
            stk_label = 1
            thisday = day2_stk[0]
            lastday = day2_stk[1]
            stk_group=fundk_stk_df.groupby('Date')
            stk_df_thisday = stk_group.get_group(thisday)
            stk_df_lastday = stk_group.get_group(lastday)
            print('取交集前',len(list(stk_df_lastday['Name'])),len(list(stk_df_thisday['Name'])))
            #stk_df_thisday与stk_df_lastday取共有项
            stk_df_thisday = stk_df_thisday[stk_df_thisday.Name.isin(list(stk_df_lastday['Name']))]
            stk_df_lastday = stk_df_lastday[stk_df_lastday.Name.isin(list(stk_df_thisday['Name']))]
            print('取交集后',len(list(stk_df_lastday['Name'])),len(list(stk_df_thisday['Name'])))
            stk_df_thisday = stk_df_thisday.reset_index(drop=True)
            stk_df_lastday = stk_df_lastday.reset_index(drop=True)
            #for aaa in stk_df_thisday['Code']:
                #print(aaa,type(aaa))
            for i in range(0,len(stk_df_thisday['Code'])):
                stk_df_thisday.loc[i,'Code'] = str(stk_df_thisday.loc[i,'Code'])
            stk_df_thisday['Code'] = sorted(stk_df_thisday['Code'])
            for i in range(0,len(stk_df_lastday['Code'])):
                stk_df_lastday.loc[i,'Code'] = str(stk_df_lastday.loc[i,'Code'])
            stk_df_lastday['Code'] = sorted(stk_df_lastday['Code'])
            stk_df_thisday = stk_df_thisday.sort_values(by='Code')
            stk_df_thisday = stk_df_thisday.reset_index(drop=True)
            stk_df_lastday = stk_df_lastday.sort_values(by='Code')
            stk_df_lastday = stk_df_lastday.reset_index(drop=True)
            stk_df_lastday.Price = stk_df_thisday.Price - stk_df_lastday.Price
            fnl_stk_df = stk_df_lastday
            fnl_stk_df.rename(columns={'Price':'DeltaPrice'}, inplace = True)
            fnl_stk_df['Profit'] = fnl_stk_df['DeltaPrice'] * fnl_stk_df['Quantity']
            fund_name = fnl_stk_df.loc[0,:]['Fundname']
            fund_NV = fnl_stk_df.loc[0,:]['NV']
            stk_sum_profit = sum(list(fnl_stk_df['Profit']))
            stk_profit_rate = stk_sum_profit / fnl_stk_df.loc[0,:]['NV']
            print('fund_name','fund_NV','stk_sum_profit','stk_profit_rate')
            print(fund_name, fund_NV, stk_sum_profit, stk_profit_rate)
        else:
            stk_label = 0
        # 处理基金
        day2_fnd = list(set(list(fundk_fnd_df['Date'])))
        if len(day2_fnd) == 2:
            print('处理基金')
            fnd_label = 1
            thisday = day2_fnd[0]
            lastday = day2_fnd[1]
            fnd_group=fundk_fnd_df.groupby('Date')
            fnd_df_thisday = fnd_group.get_group(thisday)
            fnd_df_lastday = fnd_group.get_group(lastday)
            print('取交集前',len(list(fnd_df_lastday['Name'])),len(list(fnd_df_thisday['Name'])))
            #fnd_df_thisday与fnd_df_lastday取共有项
            fnd_df_thisday = fnd_df_thisday[fnd_df_thisday.Name.isin(list(fnd_df_lastday['Name']))]
            fnd_df_lastday = fnd_df_lastday[fnd_df_lastday.Name.isin(list(fnd_df_thisday['Name']))]
            print('取交集后',len(list(fnd_df_lastday['Name'])),len(list(fnd_df_thisday['Name'])))
            fnd_df_thisday = fnd_df_thisday.sort_values(by='Code')
            fnd_df_thisday = fnd_df_thisday.reset_index(drop=True)
            fnd_df_lastday = fnd_df_lastday.sort_values(by='Code')
            fnd_df_lastday = fnd_df_lastday.reset_index(drop=True)
            fnd_df_lastday.Price = fnd_df_thisday.Price - fnd_df_lastday.Price
            fnl_fnd_df = fnd_df_lastday
            fnl_fnd_df.rename(columns={'Price':'DeltaPrice'}, inplace = True)
            fnl_fnd_df['Profit'] = fnl_fnd_df['DeltaPrice'] * fnl_fnd_df['Quantity']
            fund_name = fnl_fnd_df.loc[0,:]['Fundname']
            fund_NV = fnl_fnd_df.loc[0,:]['NV']
            fnd_sum_profit = sum(list(fnl_fnd_df['Profit']))
            fnd_profit_rate = fnd_sum_profit / fnl_fnd_df.loc[0,:]['NV']
            print('fund_name','fund_NV','fnd_sum_profit','fnd_profit_rate')
            print(fund_name, fund_NV, fnd_sum_profit, fnd_profit_rate)
        else:
            fnd_label = 0
        # 处理债券
        day2_bnd = list(set(list(fundk_bnd_df['Date'])))
        if len(day2_bnd) == 2:
            print('处理债券')
            bnd_label = 1
            thisday = day2_bnd[0]
            lastday = day2_bnd[1]
            bnd_group=fundk_bnd_df.groupby('Date')
            bnd_df_thisday = bnd_group.get_group(thisday)
            bnd_df_lastday = bnd_group.get_group(lastday)
            print('取交集前',len(list(bnd_df_lastday['Name'])),len(list(bnd_df_thisday['Name'])))
            #bnd_df_thisday与bnd_df_lastday取共有项
            bnd_df_thisday = bnd_df_thisday[bnd_df_thisday.Name.isin(list(bnd_df_lastday['Name']))]
            bnd_df_lastday = bnd_df_lastday[bnd_df_lastday.Name.isin(list(bnd_df_thisday['Name']))]
            print('取交集后',len(list(bnd_df_lastday['Name'])),len(list(bnd_df_thisday['Name'])))
            bnd_df_thisday = bnd_df_thisday.sort_values(by='Code')
            bnd_df_thisday = bnd_df_thisday.reset_index(drop=True)
            bnd_df_lastday = bnd_df_lastday.sort_values(by='Code')
            bnd_df_lastday = bnd_df_lastday.reset_index(drop=True)
            bnd_df_lastday.Price = bnd_df_thisday.Price - bnd_df_lastday.Price
            fnl_bnd_df = bnd_df_lastday
            fnl_bnd_df = fnl_bnd_df.dropna()
            fnl_bnd_df = fnl_bnd_df.reset_index(drop=True)
            fnl_bnd_df.rename(columns={'Price':'DeltaPrice'}, inplace = True)
            fnl_bnd_df['Profit'] = fnl_bnd_df['DeltaPrice'] * fnl_bnd_df['Quantity']
            fund_name = fnl_bnd_df.loc[0,:]['Fundname']
            fund_NV = fnl_bnd_df.loc[0,:]['NV']
            bnd_sum_profit = sum(list(fnl_bnd_df['Profit']))
            bnd_profit_rate = bnd_sum_profit / fnl_bnd_df.loc[0,:]['NV']
            print('fund_name','fund_NV','bnd_sum_profit','bnd_profit_rate')
            print(fund_name, fund_NV, bnd_sum_profit, bnd_profit_rate)            
        else:
            bnd_label = 0
        # 处理衍生品   
        day2_drv = list(set(list(fundk_drv_df['Date'])))
        if len(day2_drv) == 2:
            print('处理衍生品')
            drv_label = 1
            thisday = day2_drv[0]
            lastday = day2_drv[1]
            drv_group=fundk_drv_df.groupby('Date')
            drv_df_thisday = drv_group.get_group(thisday)
            drv_df_lastday = drv_group.get_group(lastday)
            print('取交集前',len(list(drv_df_lastday['Name'])),len(list(drv_df_thisday['Name'])))
            #drv_df_thisday与drv_df_lastday取共有项
            drv_df_thisday = drv_df_thisday[drv_df_thisday.Name.isin(list(drv_df_lastday['Name']))]
            drv_df_lastday = drv_df_lastday[drv_df_lastday.Name.isin(list(drv_df_thisday['Name']))]
            print('取交集后',len(list(drv_df_lastday['Name'])),len(list(drv_df_thisday['Name'])))
            drv_df_thisday = drv_df_thisday.sort_values(by='Code')
            drv_df_thisday = drv_df_thisday.reset_index(drop=True)
            drv_df_lastday = drv_df_lastday.sort_values(by='Code')
            drv_df_lastday = drv_df_lastday.reset_index(drop=True)
            drv_df_lastday.Price = drv_df_thisday.Price - drv_df_lastday.Price
            fnl_drv_df = drv_df_lastday
            fnl_drv_df.rename(columns={'Price':'DeltaPrice'}, inplace = True)
            fnl_drv_df['Profit'] = fnl_drv_df['DeltaPrice'] * fnl_drv_df['Quantity']
            fund_name = fnl_drv_df.loc[0,:]['Fundname']
            fund_NV = fnl_drv_df.loc[0,:]['NV']
            drv_sum_profit = sum(list(fnl_drv_df['Profit']))
            drv_profit_rate = drv_sum_profit / fnl_drv_df.loc[0,:]['NV']
            print('fund_name','fund_NV','drv_sum_profit','drv_profit_rate')
            print(fund_name, fund_NV, drv_sum_profit, drv_profit_rate)            
        else:
            drv_label = 0
        wb_New2.sheets('表头').range('U7:U7').offset(count,0).value = fund_name[:6]
        wb_New2.sheets('表头').range('W7:W7').offset(count,0).value = stk_profit_rate if stk_label == 1 else '/'
        wb_New2.sheets('表头').range('X7:X7').offset(count,0).value = bnd_profit_rate if bnd_label == 1 else '/'
        wb_New2.sheets('表头').range('Y7:Y7').offset(count,0).value = fnd_profit_rate if fnd_label == 1 else '/'
        wb_New2.sheets('表头').range('Z7:Z7').offset(count,0).value = drv_profit_rate if drv_label == 1 else '/'
        wb_prf_rate_detail.sheets('Sheet1').range('A2:A2').offset(count,0).value = fund_name[:6]
        wb_prf_rate_detail.sheets('Sheet1').range('C2:C2').offset(count,0).value = fund_NV
        wb_prf_rate_detail.sheets('Sheet1').range('D2:D2').offset(count,0).value = stk_sum_profit if stk_label == 1 else '/'
        wb_prf_rate_detail.sheets('Sheet1').range('E2:E2').offset(count,0).value = bnd_sum_profit if bnd_label == 1 else '/'
        wb_prf_rate_detail.sheets('Sheet1').range('F2:F2').offset(count,0).value = fnd_sum_profit if fnd_label == 1 else '/'
        wb_prf_rate_detail.sheets('Sheet1').range('G2:G2').offset(count,0).value = drv_sum_profit if drv_label == 1 else '/'  
        wb_prf_rate_detail.sheets('Sheet1').range('H2:H2').offset(count,0).value = stk_profit_rate if stk_label == 1 else '/'
        wb_prf_rate_detail.sheets('Sheet1').range('I2:I2').offset(count,0).value = bnd_profit_rate if bnd_label == 1 else '/'
        wb_prf_rate_detail.sheets('Sheet1').range('J2:J2').offset(count,0).value = fnd_profit_rate if fnd_label == 1 else '/'
        wb_prf_rate_detail.sheets('Sheet1').range('K2:K2').offset(count,0).value = drv_profit_rate if drv_label == 1 else '/'        
        count += 1
    wb_New2.save(path_now+"\Output\底仓日报汇总.xlsm")
    wb_prf_rate_detail.save(path_now+"\Output\大类资产盈亏导出.xlsx")
    wb_merge.close()
    mid_excel0.close()
    mid_excel.close()
        
        
        
        
        