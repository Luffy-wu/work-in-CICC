# -*- coding: utf-8 -*-

import numpy as np
import pandas as pd
import xlwings as xw
from datetime import datetime
from string import digits
import glob
import os
import tensorflow as tf
from FileDialog import *
import tkFileDialog


#import pdb 
#该程序目前适合的产品：包括中金资管，农行系，招商资管,中信和国君资管(需要检查一下估值表格式有没有问题)
#该程序的原理是通过识别估值表中的会计科目代码，提取持仓的信息，因为不同基金和资管的估值表格式不一样，所以用了很多if判断
print("输入1代表非有小数点类型估植表，输入2代表小数点类型：\n")
choice=input("请输入：")

filename = tkFileDialog.askopenfilename(initialdir ='/Users/luffy/')
#dlg = win64ui.CreateFileDialog(1) # 1表示打开文件对话框
#dlg.SetOFNInitialDir(r'C:\Users\Lixinglin\Desktop') # 设置打开文件对话框中的初始显示目录
#dlg.DoModal()
file_name_eg = filename.GetPathName()
file_path = os.path.dirname(file_name_eg) # 获取选择的文件的母文件夹地址
new_excel_name = os.path.split(file_name_eg)[1] #给新的excel起名
new_excel_name = new_excel_name.split(".")[0][:-6] + ".xlsx"


class AssetClass(object):
    bond = '债券'
    bond_subtypes = ['债券','ABS', '商业性债', '私募债', '政策性债', '央行票据', 
                     '分离债', '次级债金券', '企债', '国债', '短期融资', 
                     '可转债', '公司债', '标准券','可交换债']
    
    equity = '股票'
    equity_subtypes = ['非公优先股', '股票', '优先股', '创业板', 'B转H']
    
    fund = '基金'
    fund_subtypes = ['ETF', '开放基金', '基金','场外基金']
    
    derivative = '衍生品'
    derivative_subtypes = ['SWAP', '权证', '期货', 
                           '指数', '期权', '股指期货', '商品期货', '国债期货']
    cash = '现金'
    cash_subtypes = ['银行存款', '清算备付金', '券商保证金']
    
    margin = '保证金'
    margin_subtypes = ['期货交易存出保证金', '个股期权存出保证金']    
        
    others = '其他'
    others_subtypes = ['债券借贷', '指定', '网络服务','正回购','逆回购']

    not_cash_types = ['债券', '股票', '基金', '保证金','资产管理计划','SWAP','正回购','逆回购','ABS','可转债','可交换债','期货交易存出保证金', '个股期权存出保证金']
    
    associative_list = []
    associative_list.extend([(x, '债券') for x in bond_subtypes])
    associative_list.extend([(x, '股票') for x in equity_subtypes])
    associative_list.extend([(x, '基金') for x in fund_subtypes])
    associative_list.extend([(x, '衍生品') for x in derivative_subtypes])
    associative_list.extend([(x, '现金') for x in cash_subtypes])
    associative_list.extend([(x, '保证金') for x in margin_subtypes])
    associative_list.extend([(x, '其他') for x in others_subtypes])
    
    mapping = dict(associative_list)

class AccountingSubjects_1(object):
        
    pfunds = '110802'     ############## 新增项 注意和1108区分 #######################################
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
    
    #其他理财(场外期权)
    OTC = '1109'
    OTC_digits_min = 9
    
    # 债券
    bonds = '1103'
    bonds_sh = ['11031']
    bonds_sz = ['11033']
    bonds_ib = ['11035']
    
    #回购 
    buyback = ['2202','1202']
    buyback_digits = 14
    
    #ABS
    ABS = '1104'
    ABS_digits = 14
    
    # 基金投资
    funds = '1105'
    
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
    bond_futures = ['TF','T1']
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
               options_mktValue,
               '资产类合计:']
               
    set_examime = set(examine)

class AccountingSubjects_2(object):
    # 银行存款
    cash_deposits = '1002'
    
    # 清算备付金
    cash_provisions = '1021'
    
    # 证券清算款 
    cash_Liquidation = '3003'
    
    # 存出保证金
    margin='1031'
    margin_sh='103101'
    margin_sz='103102'
    margin_hgt='103106'
    margin_sgt='103107'
    margin_swap='103110'
    margin_jjs='103104'
    margin_etf='103105'
    margin_future = '103103'
    margin_option = '103111'
    
    money_all = [cash_deposits,cash_Liquidation]
    money_all_digits = 8
    
    other_money = [cash_provisions,margin]
    other_money_digits = 6
    
    # 股票投资
    equities = '1102'
    equities_digits = 17
    
    #其他理财(场外期权)####################################
    OTC = '1108'
    OTC_digits_min = 9
    
    # 债券
    bonds = '1103'
    bonds_sh = ['11030','11032','11031']
    bonds_sz = ['11033','11034','11035']
    bonds_ib = ['11036']
    
    #回购 
    buyback = ['2202','1202']
    buyback_digits = 17
    
    #ABS
    ABS = '1104'
    ABS_digits = 17
    
    # 基金投资
    funds = '1105'
    pfunds = '110802'     ############## 新增项 注意和1108区分 #######################################
    
    # 互换和场外其他
    swaps = '1021'
    
    # 场内期权 ###############################
    options = '1041'
    options_margins = '104102'
    options_mktValue = '104103'
    
    # 衍生工具和套期工具
    derivatives = '3102'
    hedging_positions = '3201'
    index_futures = ['IC', 'IF', 'IH']
    bond_futures = ['TF','T']
    futures_digits = 17
    all_derivatives = [derivatives, hedging_positions, swaps]
    
    examine = [cash_deposits, 
               cash_provisions, 
               equities,
               margin,
               margin_future,
               margin_option,
               derivatives, 
               hedging_positions, 
               swaps,
               bonds, 
               funds, 
               options,
               options_margins,
               options_mktValue,
               '资产类合计:']
               
    set_examime = set(examine)

# helper function
remove_digits = str.maketrans('', '', digits)

if choice=='2':
        AccountingSubjects=AccountingSubjects_2
elif choice=='1':
        AccountingSubjects=AccountingSubjects_1 

def isnumm(value):
    try:
        value + 1
    except TypeError:
        return -1
    else:
        return 1
        
def map_name_to_bond_subtype(name):
    if name[-2::] == 'EB':
        return '可交换债'
    elif name[-2::] == '转债':
        return '可转债'
    else: return '债券'
    
def map_code_to_future_ticker(code):
    first_letter = code.translate(remove_digits)[0]       
    ind = code.index(first_letter)
    return code[ind::]

def map_code_to_bond_ticker(code):
    prefix = code[8::]
    if code[:5] in AccountingSubjects.bonds_sh:
        suffix = '.SH'
    elif code[:5] in AccountingSubjects.bonds_sz:
        suffix = '.SZ'
    else:
        suffix = '.IB'
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
def glob_file_names():
    f_list1 = glob.glob(file_path+"\*.xls")
    f_list2 = glob.glob(file_path+"\*.xlsx")
    f_list1.extend(f_list2)
    return f_list1

def export_positions(file, counter):
    wb = xw.Book(file)
    sht = wb.sheets[0]
    #确定标题行在excel中所在的位置(行)
    time_temp = list(sht.range('A1:A5').value)
    time_temp=[u"" if x is None else x for x in time_temp]
    header_index = 0
#    pdb.set_trace() 
    for i in time_temp:
        if "科目" in i:
            break
        else:header_index += 1
    #从估值表的名称中提取日期 
    time_cell =[] 
    for l in file:
        if l=='(':
            break
        elif l in digits:
            time_cell.append(l)

    
    #[l for l in file if l in digits]
    if len(time_cell)<8:
        print("估值表名称中未找到日期，请在这句话的位置更改程序，从估值表中提取日期")
        os._exit(0)#退出程序 
    time_cell = time_cell[-8::]
    print(time_cell)
    time_cell_str = ''.join(time_cell)
    this_date = datetime.strptime(time_cell_str, '%Y%m%d')
    print(time_cell_str)
    #把excel读入Dataframe
    df = pd.read_excel(file, header=header_index)
    df = df.fillna('None')
    #去掉列名中的空格,给列重命名
    column_name = [col.replace(' ','') for col in df.columns] #逐一检查，去掉空格，生成新列名
    if '科目编码' in column_name:
        column_name[column_name.index('科目编码')] = '科目代码'
    if '证券市值' in column_name:#把表示市值的列命名为"市值"
        column_name[column_name.index('证券市值')] = '市值'
    elif '市值-本币' in column_name:
        column_name[column_name.index('市值-本币')] = '市值'
    elif '市值' not in column_name:
        print('没找到表示市值的列，请查看估值表除了什么幺蛾子并且在此修改程序')
        os._exit(0)
    if '证券数量'  in column_name: #把表示数量的列命名为"数量"
        column_name[column_name.index('证券数量')] = '数量'
    elif '数量' not in column_name:
        print('没找到表示数量的列，请查看估值表除了什么幺蛾子并且在此修改程序')
    if '行情收市价' in column_name:#把表示价格的列命名为"市价"
        column_name[column_name.index('行情收市价')] = '市价'
    elif '行情' in column_name: 
        column_name[column_name.index('行情')] = '市价'
    elif '行情价格' in column_name: 
        column_name[column_name.index('行情价格')] = '市价'
    elif '市价' not in column_name:
        print('没有找到市价，将用市值除以数量计算市价')
    if '估值增值-本币' in column_name:    #把表示估值增值的列命名为"估值增值"
        column_name[column_name.index('估值增值-本币')] = '估值增值'
    df.columns = column_name
#    pdb.set_trace() 
    for row in df.index:
        if not isinstance(df['数量'][row],float):
            if  df['数量'][row] != 'None':
                df['数量'][row]=df['数量'][row].replace(",","")
                df['数量'][row]=float(df['数量'][row])
    for row in df.index:
        if not isinstance(df['市值'][row],float):
            if  df['市值'][row] != 'None':
                df['市值'][row]=df['市值'][row].replace(",","")
                df['市值'][row]=float(df['市值'][row])
#    for row in df.index:
#        if  df['市值占净值%'][row] != 'None':
#            df['市值占净值%'][row].replace(",","")
#            df['市值占净值%'][row]=float(df['市值占净值%'][row])       
    for row in df.index:
        df['科目代码'][row] = str(df['科目代码'][row])
    if not AccountingSubjects.set_examime.intersection(set(df['科目代码'].values)):
        # empty table
        print('[Note]: Skip empty file {} at {}.'.format(counter, time_cell_str))
        wb.close()
        return [],[]
    #找到资产净值和总资产
#    pdb.set_trace() 
    col = list(df['科目代码'])
    if '基金资产净值:' in col:
        net_row = col.index('基金资产净值:') #种类为基金估值表
    elif '集合计划资产净值：' in col:
        net_row = col.index('集合计划资产净值：') #种类为集合计划估值表
    elif '资产净值' in col:
        net_row = col.index('资产净值')#种类为招商资管的估值表？
    else: 
        print("没找到净资产")
        net_row = [] #目前只见到洛书是这样的
    if '资产类合计：' in col:
        asset_row = col.index('资产类合计：')
    elif '资产类合计:' in col:
        asset_row = col.index('资产类合计:')
    elif '资产合计' in col:
        asset_row = col.index('资产合计')
    else: 
        print("没找到总资产")
        asset_row = [] 
    assets = dict()
    if  not net_row or not asset_row: #如果没找到净资产或者总资产，就用证券投资合计/占总资产比例来计算 
        col_temp = ''
        if '证券投资合计:' in col:
            col_temp = '证券投资合计:'
        elif '证券投资合计：' in col:
            col_temp = '证券投资合计：'
        elif '1002' in col:#都找不到就去找银行存款
            col_temp = '1002' 
        else: 
            print("连银行存款都没找到,无法计算总资产,请查看估值表出了什么幺蛾子并在此修改程序")
            os._exit(0)
        if '市值占净值%' in column_name:
            asset_row = col.index(col_temp)
            assets['val'] = df['市值'][asset_row]/df['市值占净值%'][asset_row]
        elif '市值占比' in column_name:
            asset_row = col.index(col_temp)
            assets['val'] = df['市值'][asset_row]/df['市值占比'][asset_row]
        else:#其它情况，目前只有洛书的估值表出现这种情况,数量那一列其实是比例 
            print("此估值表为极其特殊的洛书估值表，请自行提取银行存款，备付金，保证金")
            asset_row = col.index(col_temp)
            assets['val'] = df['市值'][asset_row]/df['数量'][asset_row]
        assets['net'] = assets['val']
        assets['ratio'] = 1
    else:
        assets['val'] = df['市值'][asset_row]
        assets['net'] = df['市值'][net_row]
        if isnumm(assets['val']) == -1:
            assets['val'] = float(assets['val'].replace(',',''))
        if isnumm(assets['net']) == -1:
            assets['net'] = float(assets['net'].replace(',',''))
        assets['ratio'] = assets['val']/assets['net']
    output = []
    net_val = 0
    ###########################################################################
    for row in df.index:
        code = df['科目代码'][row]
        if choice=='2':
            if isinstance(code,str):
                code=code.replace(".","")
                code=code.replace(" ",".")
        mktVal = df['市值'][row]
         
        #提取单位净值
        if code == '基金单位净值：' or code == '基金单位净值:'or code == '单位净值' or code =='今日单位净值：' or code =='今日单位净值':
            net_val = df['科目名称'][row]
           # pdb.set_trace()
        if isinstance(mktVal,str):#如果市值那列是字符串，过滤
            continue
        mktVal_r = mktVal/assets['val']#此处用的总资产计算比例，如果用净资产的话会因为基金赎回和负债等因素导致比例和大于1
        qty = df['数量'][row] if '数量' in column_name else df['证券数量'][row]
        name = df['科目名称'][row]
        valAdd = df['估值增值'][row] if '估值增值' in column_name else 0
        p = df['市价'][row] if '市价' in column_name else mktVal/qty
        # deal with margin
        if code == AccountingSubjects.margin_future:
            a_subtype = '期货交易存出保证金'
            a_type = AssetClass.mapping[a_subtype]
            pos = [this_date, 'cash_CNY', 'cash_CNY', mktVal, 1, 
                   mktVal, mktVal_r, a_type, a_subtype, valAdd]
            output.append(pos)        
        elif code == AccountingSubjects.margin_option:
            a_subtype  = '个股期权存出保证金'
            a_type = AssetClass.mapping[a_subtype]
            pos = [this_date, 'cash_CNY', 'cash_CNY', mktVal, 1, 
                   mktVal, mktVal_r, a_type, a_subtype, valAdd]
            output.append(pos)
        #deal with buyback
        elif str(code)[:4] in AccountingSubjects.buyback:
            if len(code) == AccountingSubjects.buyback_digits:
                ticker = name
                a_subtype = '正回购' if code[:4]=='2202' else '逆回购'
                a_type = '其他'
                qty = 1 
                p = mktVal
                mktVal_r = mktVal_r
                pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd]
                output.append(pos)  
        # filter
        elif isinstance(qty,str) and isinstance(p,str):
            continue
        elif qty == 0 or p == 0:
            continue
        # deal with equities
        elif code[:4] == AccountingSubjects.equities:
            if isinstance(qty,str) or isinstance(mktVal,str):
                # no quantity or ratio or mktvalue
                continue
            if len(code) >= AccountingSubjects.equities_digits-1: #港股的代码比A股少一位
                ticker = code[8::] #招商资管的表股票和基金是20位的，后9位是股票和基金的代码
                if ticker[0] == 'H':
                    # deal with HK share
                    ticker = str(int(ticker[1::])) + '.HK'
                elif ticker[-2:]=='SH' or ticker[-2:]=='SZ':
                    ticker=ticker[:-3]
                elif ticker[-2:]=='HG':
                    ticker=str(int(ticker[:-3]))+".HK"
                a_subtype = '股票'
                a_type = AssetClass.mapping[a_subtype]
                if np.isnan(mktVal_r): mktVal_r = 0
                pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd]
                output.append(pos)
        elif code[:4] == AccountingSubjects.OTC and code[:6] != AccountingSubjects.pfunds :   ################# 新增项：and语句 ################
            if len(code)>AccountingSubjects.OTC_digits_min:
                ticker = code[(AccountingSubjects.OTC_digits_min-1):]
                a_subtype = '场外期权'
                a_type = '其它理财'
                qty = -qty if mktVal<0 else qty
                pos = [this_date, ticker, name, qty, p, 
                           mktVal, mktVal_r, a_type, a_subtype, valAdd]
                output.append(pos)            
        # deal with funds
        elif code[:4] == AccountingSubjects.funds:
            if isinstance(qty,str) or isinstance(mktVal,str):
                # no quantity or ratio or mktvalue
                continue
            if len(code) >= AccountingSubjects.equities_digits:
                ticker = code[8:14] #招商资管的表股票和基金是20位的，后9位是股票和基金的代码
                a_subtype = '基金' 
                a_type = AssetClass.mapping[a_subtype]
                if np.isnan(mktVal_r): mktVal_r = 0
                pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd]
                output.append(pos)
        # deal with bonds
        elif code[:4] == AccountingSubjects.bonds:
            if isinstance(qty,str) or isinstance(mktVal,str):
                # no quantity or ratio or mktvalue
                continue
            if len(code) == AccountingSubjects.equities_digits or len(code)==17:#有一些表的债券是17位？
                ticker = map_code_to_bond_ticker(code)
                a_subtype = map_name_to_bond_subtype(name)
                a_type = '债券'
                if np.isnan(mktVal_r): mktVal_r = 0
                pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd]
                #print(ticker, name, qty, p, mktVal, mktVal_r, a_type, a_subtype)
                output.append(pos)
        elif code[:4] == AccountingSubjects.ABS:
            if len(code) == AccountingSubjects.ABS_digits:
                ticker = code[-6::]
                a_subtype = 'ABS'
                a_type = '债券'
                pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd]
                output.append(pos)  
        # deal with derivatives
        elif code[:4] in AccountingSubjects.all_derivatives:
            if isinstance(qty,str) or len(code)<=10:
                # no quantity
                continue
            if not code.translate(remove_digits):#如果没有字母，则是委外资产管理计划
                ticker = code
                a_type = '其他'
                a_subtype = '资产管理计划'
            ########### 新增项：下一行 or 后面的内容：有的表的futures_digits也即科目代码的长度会与其他不同（比如多1位），可能会导致bug，所以加上or ####################
            elif len(code) == AccountingSubjects.futures_digits or len(code) == AccountingSubjects.futures_digits +1 :      
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
            qty = -np.abs(qty) if mktVal<0 else qty
            pos = [this_date, ticker, name, qty, p, mktVal, mktVal_r, a_type, a_subtype, valAdd]
            output.append(pos)
            
        # deal with options
        elif code[:6] == AccountingSubjects.options_mktValue:
            if isinstance(mktVal,str):
                # no value
                continue
            a_subtype = '场内期权'
            a_type = '场内期权'
            qty = -np.abs(qty) if mktVal<0 else qty
            pos = [this_date, '场内期权总市值', '场内期权总市值', mktVal, 1, 
                   mktVal, mktVal_r, a_type, a_subtype, valAdd]
            output.append(pos)
            
        # deal with pfunds
        elif code[:6] == AccountingSubjects.pfunds:
            if isinstance(qty,str) or isinstance(mktVal,str):
                # no quantity or ratio or mktvalue
                continue
            if len(code) >= AccountingSubjects.equities_digits:
                ticker = code[8:14] #招商资管的表股票和基金是20位的，后9位是股票和基金的代码
                a_subtype = '场外基金' 
                a_type = AssetClass.mapping[a_subtype]
                if np.isnan(mktVal_r): mktVal_r = 0
                pos = [this_date, ticker, name, qty, p, 
                       mktVal, mktVal_r, a_type, a_subtype, valAdd]
                output.append(pos)                

        # deal with options margin
        elif code[:6] == AccountingSubjects.options_margins:
            if isinstance(mktVal,str):
                # no market value
                continue
            a_subtype = '场内期权保证金'
            a_type = '场内期权保证金'
            pos = [this_date, 'cash_CNY', 'cash_CNY', mktVal, 1, 
                   mktVal, mktVal_r, a_type, a_subtype, valAdd]
            output.append(pos)
    print(output)        
    # append one more cash row
    agg_ratio = sum([pos[6] for pos in output 
                     if pos[-2] in AssetClass.not_cash_types])
    agg_mkt_val = sum([pos[5] for pos in output 
                       if pos[-2] in AssetClass.not_cash_types])
    if not agg_ratio:
        # no assets holdings
        cash_ratio = assets['ratio']
        cash_val = assets['val']
    else:
        cash_ratio = 1 - agg_ratio #之前是assets['ratio'] - agg_ratio
        cash_val = assets['val'] - agg_mkt_val  
    cash_pos = [this_date, 'cash_CNY', 'cash_CNY', cash_val, 1, 
        cash_val, cash_ratio, '现金', '资金余额', 0]
    output.append(cash_pos)
    #从E3:L3提取基金净值
    if net_val == 0:#如果在估值表正文中没有找到单位净值，那么看看从E3:L3中有没有
        net_val_cell = sht.range('E3:L3').value
        temp = 0
        for temp in net_val_cell:
            if not isinstance(temp,str):
                continue
            elif temp[:4] == '单位净值':
                net_val = temp
                break
        if temp == 0:
            print('E3:L3中还是没有找到基金的单位净值，请检查估值表，并在此处更改程序')
            os._exit(0)
        else:
            net_val = net_val.split(':')[1] if ':' in net_val else net_val.split('：')[1]
            net_val = float(net_val[:5])
    net_value = [[this_date,net_val,assets['val'],assets['net']]]
    wb.close()
    #print(output)
    return output,net_value

if __name__ == '__main__':#该模块只能直接运行，不能在别的地方导入运行
    files_to_read = glob_file_names()
    all_positions_history = []
    all_net_value_history = []
    num_files = len(files_to_read)
    dtemp=file_path.split(os.sep)
    dtemp[-1]=new_excel_name
    new_excel_name=os.sep.join(dtemp)
    counter = 1
    for f in files_to_read:
        print(f)
        output,net_value = export_positions(f, counter)
        all_positions_history.extend(output)
        all_net_value_history.extend(net_value)
        print('Finished {} in {}.'.format(counter, num_files))
        counter += 1
    all_positions_df = pd.DataFrame(all_positions_history)
    all_net_value_df = pd.DataFrame(all_net_value_history)
    with pd.ExcelWriter(new_excel_name) as writer:    
        all_positions_df.to_excel(writer,sheet_name='position', index=False, 
                   header=('Date', 'Code', 'Name', 'Quantity', 'Price', 
                   'Turnover', 'Ratio', 'Asset_Type', 'Asset_Subtype', 'ValueAdd'))
        all_net_value_df.to_excel(writer,sheet_name='net_value', index=False, 
                   header=('Date', 'net_value','asset','net_asset'))
    xw.Book(new_excel_name)
    



