# -*- coding: utf-8 -*-

import numpy as np
import pandas as pd
import xlwings as xw
import os
import win32ui


#获取原数据
dlg = win32ui.CreateFileDialog(1) # 1表示打开文件对话框
dlg.SetOFNInitialDir(r'C:\Users\Lixiangwu\Desktop') # 设置打开文件对话框中的初始显示目录
dlg.DoModal()
file = dlg.GetPathName()
file_path = os.path.dirname(file) # 获取选择的文件的母文件夹地址


#如果含有基金则需判断是否有货币基金，如果有删去
def del_MF(posi,sht1):
    print("发现持仓包含基金，正在移除货币型基金\n")
    npo=len(posi)
    sht1.range("k1").value="monetary_fund"
    
    for i in range(npo):
        sht1.cells(i+2,11).value='=IF(\"{}\"=\"基金\",IF(ISNUMBER(FIND(\"货币\",F_Info_InvestType(\"{}\"))),1,0),0)'.format(posi.Asset_Type[i],posi.Code[i])
    
    pomf=sht1.range("a2").expand().value
    pomf=pd.DataFrame(pomf,columns=sht1.range("a1").expand("right").value)
    pomf=pomf[pomf.monetary_fund != 1]
    for i in range(len(pomf)):
        if "HK" in pomf.iloc[i,1] and pomf.iloc[i,7]=="股票":
            pomf.iat[i,8]="港股"
    print("移除完成\n")
    return pomf.iloc[:,0:10]

#已经拥有正确的现金表达式之后，生成一个新的execl文件以供直接粘贴到factset模板 ，将nocash和cash merge之后计算ratio   
def to_factset(nocash,cash,nv,file):
    fp=file.split(os.sep)[0:-1]
    filename=file.split(os.sep)[-1]
    filename=filename.split('.')[0]
    filename=filename+'factset辅助.xlsx'
    fp.append(filename)
    filename=os.sep.join(fp)
    
    cash=pd.DataFrame(cash,columns=nocash.columns)
    df=[cash,nocash]
    df=pd.concat(df)
    for t in range(len(df)):
        tempdate=df.iat[t,0]
        tvalue=nv[nv.Date==tempdate].net_asset
        tempratio=df.iat[t,5]/tvalue
        df.iat[t,6]=tempratio
    df.to_excel(filename,index=False)


#已经拥有正确的现金表达式之后，生成一个新的excel文件包含 1.正确的持仓明细（用nv算ratio） 2. 第二张sheet上直接作图
def to_asset_allocation(cash,posi,nv,file):
    fp=file.split(os.sep)[0:-1]
    filename=file.split(os.sep)[-1]
    filename=filename.split('.')[0]
    filename=filename+'大类资产.xlsx'
    fp.append(filename)
    filename=os.sep.join(fp)
    
    #将正确的持仓数据用df表示
    tempdf=posi[(posi.Asset_Subtype!='期货交易存出保证金') & (posi.Asset_Subtype!='资金余额') & \
                (posi.Asset_Subtype!='个股期权存出保证金') & (posi.Asset_Subtype!='逆回购') & (posi.Asset_Subtype!='逆回购') &\
                (posi.Asset_Type!='现金')]
    cash=pd.DataFrame(cash,columns=posi.columns)
    df=[tempdf,cash]
    df=pd.concat(df)
    
    #根据turnover计算ratio
    for t in range(len(df)):
        tempdate=df.iat[t,0]
        tvalue=nv[nv.Date==tempdate].net_asset
        tempratio=df.iat[t,5]/tvalue
        df.iat[t,6]=tempratio

    writer=pd.ExcelWriter(filename)
    
    df.to_excel(writer,index=False,sheet_name='position')
    #按照固定的格式生成大类资产配置表
    aa=df.groupby(['Date','Asset_Subtype'])['Ratio'].sum()
    ast=df.Asset_Subtype.unique()
    aadf=pd.DataFrame(np.zeros([len(nv),len(ast)]),columns=ast,index=nv.Date)
    Date=[]
    asst=[]
    for i in range(len(aa)):
        Date.append(aa.index[i][0])
        asst.append(aa.index[i][1])
    aa=pd.DataFrame(aa)
    aa['Date']=Date
    aa['ast']=asst
    for j in range(len(nv)):
        for k in range(len(ast)):
            if not aa[(aa.Date==nv.Date[j])&(aa.ast==ast[k])].empty:
                aadf.iat[j,k]=aa[(aa.Date==nv.Date[j])&(aa.ast==ast[k])].Ratio[0]
            '''for h in range(len(aa)):
                if aa.index[h][0]==nv.Date[j] and aa.index[h][1]==ast[k]:
                    aadf.iat[j,k]=aa[h]'''
    tempnv=nv.net_value.apply(float)
    aadf['产品净值']=list(tempnv)
    aadf.to_excel(writer,sheet_name='asset_allocation')
    writer.save()
    writer.close()

    
if __name__ == '__main__':
    wb=xw.Book(file)
    sht1=wb.sheets("position")
    sht2=wb.sheets("net_value")
    posi=sht1.range("a2").expand().value
    nv=sht2.range("a2").expand().value
    posi=pd.DataFrame(posi,columns=sht1.range("a1").expand("right").value)
    nv=pd.DataFrame(nv,columns=sht2.range("a1").expand("right").value)
    
    # asset_subtype HK shares
    for i in range(len(posi)):
        if "HK" in posi.iloc[i,1] and posi.iloc[i,7]=="股票":
            posi.iat[i,8]="港股"
        
    # is monetary fund included?
    if len(posi[posi.Asset_Subtype=='基金'])>0:
        posi=del_MF(posi,sht1)
    print(posi.Asset_Subtype.unique())
    
    #删去原来的现金，利用nv中的net_asset计算现金
    nocash=posi[(posi.Asset_Type == '基金')| (posi.Asset_Type=='股票') | (posi.Asset_Type=='债券')]
    nocashAmount=nocash.groupby('Date')["Turnover"].sum()
    cash=[]
    temp=[]

    #某些产品可能没有股票债券基金的投资（如只投资期货的产品），补全其nocash部分（为0）
    if len(nocashAmount)!=len(nv):
        newnocash=pd.Series(np.zeros(len(nv))) 
        newnocash.index=nv.Date
        for nt in range(len(newnocash)):
            for nn in range(len(nocashAmount)):
                if newnocash.index[nt]==nocashAmount.index[nn]:
                    newnocash[nt]=nocashAmount[nn]
        nocashAmount=newnocash
    for i in range(len(nocashAmount)):
        quan=nv.net_asset[i]-nocashAmount[i]
        temp=[nv.Date[i],'cash_CNY','cash_CNY',quan,1,quan,0,'现金','资金余额',0]
        cash.append(temp)

    #根据不同选择利用cash和posi以及nocash生成样本
    print("输入0代表生成factset格式,输入1代表生成资产配置格式\n")
    
    choice = input("请输入0、1或2：") 
    if choice=='0':
        to_factset(nocash,cash,nv,file)
    elif choice=='1':
        to_asset_allocation(cash,posi,nv,file)
    else:
        to_factset(nocash,cash,nv,file)
        to_asset_allocation(cash,posi,nv,file)
    
    wb.close()