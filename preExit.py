# -*- coding: utf-8 -*-
"""
Created on Sat Dec 25 21:14:25 2021

@author: MarqueLiu
"""

import pandas as pd
import numpy as np
from openpyxl import load_workbook

### 退出顺序文件
def ExitOrder():
    path1=r'煤炭退出/退出/总表.xlsx'
    path2=r'煤炭退出/退出/四个情景数据.xlsx'
    path3=r'煤炭退出/所有煤矿产量.xlsx'
    path4=r'煤炭退出/退出/退出顺序.xlsx'
    
    writer=pd.ExcelWriter(path4,engine='openpyxl')
    
    for retire in ['规模','经济','排放']:
    
        df1=pd.read_excel(path1,0,index_col='UID')
        df2=pd.read_excel(path2,'消费情景',index_col='年份')
        dfPro=pd.read_excel(path3,0,index_col='UID')
        dfOut=pd.read_excel(path4,retire,index_col='UID')
        #清空中间表
        for a in dfOut.columns:
            for i in dfOut.index:
                dfOut.loc[i,a]=None
        #安全退出路径
        #engCon=df2.loc[2020:2050,'2度情景']#截取能源消耗序列
        years=range(2020,2051,1)
        #按照退出顺序反向排序
        #20220626, 这里排序应该是False, 为了算上限暂时改为True
        df1.sort_values(by=[retire],ascending=[False],inplace=True)#先排序
        
        idxList=df1.index.tolist()
    
        for sce in ['政策','强化','2度','1.5度']:
            for year in years:
                prod=0
                for idx in idxList:
                    if prod<df2.loc[year,sce]:
                        temp=dfPro.loc[idx,2019]*12/100000000
                        if pd.isna(temp):
                            continue
                        prod=prod+temp
                        continue
                    elif pd.isna(dfOut.loc[idx,sce]):
                        dfOut.loc[idx,sce]=year
                    else:
                        break
        dfOut.fillna(value=10000,inplace=True)#未退出煤矿填充空值
        dfOut.to_excel(writer,sheet_name=retire)
    writer.save()
    return

##  1108 不同路径下的校准因子
def CaliFactor():
    path1=r'煤炭退出/退出/退出顺序.xlsx'
    path2=r'煤炭退出/退出/四个情景数据.xlsx'
    path3=r'煤炭退出/退出/所有煤矿产量.xlsx'
    path4=r'煤炭退出/退出/校准因子.xlsx'
    writer=pd.ExcelWriter(path4,engine='openpyxl')
    
    for retire in ['规模','经济','排放']:
    
        df1=pd.read_excel(path1,retire,index_col='UID')
        df2=pd.read_excel(path2,'消费情景',index_col='年份')
        dfPro=pd.read_excel(path3,0,index_col='UID')
        dfFac=pd.read_excel(path4,retire,index_col='年份')
        for a in dfFac.columns:
            for i in dfFac.index:
                dfFac.loc[i,a]=None
        years=range(2020,2051,1)
        
        def getProdAbi(UIDList,year):
            prod=0
            for UID in UIDList:
                if year<=2019:
                    temp=dfPro.loc[UID,year]
                else:
                    temp=dfPro.loc[UID,2019]
                    
                if not pd.isna(temp):
                    prod=prod+temp
            return prod
        
        for year in years:
            for pathway in ['政策','强化','2度','1.5度']:
                dfT=df1[df1[pathway]>year]
                UIDList=dfT.index.tolist()
                factor=df2.loc[year,pathway]/(getProdAbi(UIDList,year)*12/100000000)
                dfFac.loc[year,pathway]=factor
        dfFac.to_excel(writer,sheet_name=retire)
    writer.save()
    return


###  价格变动和累积成本

def CostCalc():
    path1=r'煤炭退出/退出/总表.xlsx'
    
    path3=r'煤炭退出/退出/所有煤矿产量.xlsx'
    path4=r'煤炭退出/退出/校准因子.xlsx'
    path5=r'煤炭退出/退出/四个情景数据.xlsx'
    path2=r'煤炭退出/退出/退出顺序.xlsx'
    writer=pd.ExcelWriter(r'煤炭退出/退出/临时结果/成本.xlsx',engine='openpyxl')
    pathway=['政策','强化','2度','1.5度']
    for scenario in pathway:
        dfTot=pd.read_excel(path1,0,index_col='UID')
        dfPro=pd.read_excel(path3,0,index_col='UID')
        dfCsm=pd.read_excel(path5,'消费情景',index_col='年份')
        
        Years=list(np.arange(2020,2051,1))
        df=pd.DataFrame(index=Years,columns=['规模','经济','排放','规模成本','经济成本','排放成本'])
        
        for sce in ['规模','经济','排放']:
            dfPath=pd.read_excel(path2,sce,index_col='UID')
            dfFac=pd.read_excel(path4,sce,index_col='年份')
            hisCost=0
            for year in Years:
                Cost=0
                dfT=dfPath[dfPath[scenario]>year]
                UIDlist=dfT.index.tolist()
                for UID in UIDlist:
                    temp=dfPro.loc[UID,2019]
                    if not pd.isna(temp):
                        Cost=Cost+dfTot.loc[UID,'价格']*temp*dfFac.loc[year,scenario]*12#元
                hisCost=Cost+hisCost
                df.loc[year,sce+'成本']=hisCost/100000000#亿元  
                df.loc[year,sce]=Cost/(dfCsm.loc[year,scenario]*100000000)#元
        df.to_excel(writer,sheet_name=scenario)
    writer.save()
    return

def MarCurve():
    path1=r'煤炭退出/退出/总表.xlsx'
    path2=r'煤炭退出/退出/临时结果/成本曲线.xlsx'
    path3=r'煤炭退出/退出/临时结果/EF曲线.xlsx'
    
    df=pd.read_excel(path1,0,index_col='UID')
    for manner in ['价格','排放因子']:
        df.sort_values(by=[manner],ascending=[False],inplace=True)#先排序
        
        price=df[manner].unique()
        dfCost=pd.DataFrame(index=price,columns=['产能'])
        
        for pri in price:
            dfT=df[df[manner]==pri]
            dfCost.loc[pri,'产能']=dfT['月产煤量'].sum()*12/100000000#亿吨
        if manner=='价格':
            path=path2
        else:
            path=path3
        dfCost.to_excel(path,sheet_name='Sheet1')
    return


# ### 计算排放顺序
# import pandas as pd
# import math

# path1=r'煤炭退出/退出/所有煤矿产量.xlsx'
# path2=r'煤炭退出/退出/所有煤矿总表.xlsx'

# dfTot=pd.read_excel(path2,0,index_col='UID')
# dfPro=pd.read_excel(path1,0,index_col='UID')
# b=2.017
# Df=0.672#水淹
# Dd=0.302#干井
# #定义获取产量
# def getProdAbi(UIDList,year):
#     prod=0
#     for UID in UIDList:
#         if year<=2019:
#             temp=dfPro.loc[UID,year]
#         else:
#             temp=dfPro.loc[UID,2019]
            
#         if not pd.isna(temp):
#             prod=prod+temp
#     return prod

# ### 井工部分
# def getMineEmi(UID,year):
#     '''传入具体某个UID，在此基础上计算排放和'''
#     '''返回单位：kg'''
#     ef=dfTot.loc[UID,'排放因子']#unit kg/t
#     prod=getProdAbi([UID],year)*12#unit t
#     Emi=ef*prod
#     return Emi

# UIDList=dfTot[(dfTot['是否水淹']==0)&(dfTot['是否露天']==0)].index.tolist()
# for UID in UIDList:
#     if dfTot.loc[UID,'废弃时间'] != 10000 and dfTot.loc[UID,'废弃时间']!=dfTot.loc[UID,'新建时间']:
#         year=dfTot.loc[UID,'废弃时间']-1
#     elif dfTot.loc[UID,'废弃时间']==dfTot.loc[UID,'新建时间']:
#         year=dfTot.loc[UID,'废弃时间']
#     else:
#         year=2019
#     dfTot.loc[UID,'排放']=getMineEmi(UID,year)
# dfTot.to_excel(r'煤炭退出/退出/所有煤矿总表.xlsx',sheet_name='Sheet1')
