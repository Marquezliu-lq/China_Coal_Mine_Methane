# -*- coding: utf-8 -*-
"""
Created on Sat Dec 25 21:14:25 2021

@author: MarqueLiu
"""
'''trade-off surface'''
'''用于计算帕累托边界上的退出策略，与pre_exit_sp对应'''

import pandas as pd
import numpy as np
from openpyxl import load_workbook
import random

path1=r'煤炭退出/退出/总表.xlsx'
path2=r'煤炭退出/退出/四个情景数据.xlsx'
path3=r'煤炭退出/所有煤矿产量.xlsx'
path4=r'煤炭退出/退出/退出顺序.xlsx'
path5=r'煤炭退出/退出/校准因子.xlsx'
retire='SAMPLE'

dfTot=pd.read_excel(path1,0,index_col='UID')
dfSce=pd.read_excel(path2,'消费情景',index_col='年份')
dfPro=pd.read_excel(path3,0,index_col='UID')
dfSer=pd.read_excel(path4,retire,index_col='UID')
dfFac=pd.read_excel(path5,retire,index_col='年份')
    
def judgeFun(UID1,UID2,barrier):
    Emi1=dfTot.loc[UID1,'排放因子'];Emi2=dfTot.loc[UID2,'排放因子'];
    Cost1=dfTot.loc[UID1,'价格'];Cost2=dfTot.loc[UID2,'价格'];
    if (Emi1-Emi2)<barrier and (Cost1<Cost2):
        return True
    else:
        return False

def getSAMPLE(UIDList,barrier):
    #传入一个基本的序列：比如按排放因子调整过的UIDlist
    # S=[x for x in range(1, 3044)]
    # random.shuffle(S)
    #S=dfTot[['ECO-EMI','EMI-ECO']]
    #洗牌 但只洗一点点
    N=len(UIDList)
    for i in range(N):
 
        # Last i elements are already in place
        for j in range(0, N-i-1):
 
            if judgeFun(UIDList[j],UIDList[j+1],barrier) :
                UIDList[j], UIDList[j+1] = UIDList[j+1], UIDList[j]#python独有的交换方法
    return UIDList

### 退出顺序文件
def ExitOrder(barrier):    
    #writer=pd.ExcelWriter(path4,engine='openpyxl')
    dfTot.sort_values(by=['EMI-ECO'],ascending=[True],inplace=True)#先按照固定顺序获得列表
    List=dfTot.index.tolist()
    List=getSAMPLE(List,barrier)#获取调整后的UIDList
    i=1
    for UID in List:
        dfTot.loc[UID,'SAMPLE']=i
        i=i+1
    #dfTot['SAMPLE']=dfTot['EMI-ECO']
    for i in dfSer.index:
        dfSer.loc[i,'政策']=None
    #安全退出路径
    #engCon=df2.loc[2020:2050,'2度情景']#截取能源消耗序列
    years=range(2020,2051,1)
    #按照退出顺序反向排序
    dfTotT=dfTot.sort_values(by=[retire],ascending=[False])#先排序
    idxList=dfTotT.index.tolist()

    for sce in ['政策']:
        for year in years:
            prod=0
            for idx in idxList:
                if prod<dfSce.loc[year,sce]:
                    temp=dfPro.loc[idx,2019]*12/100000000
                    if pd.isna(temp):
                        continue
                    prod=prod+temp
                    continue
                elif pd.isna(dfSer.loc[idx,sce]):
                    dfSer.loc[idx,sce]=year
                else:
                    break
    dfSer.fillna(value=10000,inplace=True)#未退出煤矿填充空值
    #dfOut.to_excel(writer,sheet_name=retire)
    #writer.save()
    return dfSer

##  1108 不同路径下的校准因子
def CaliFactor(df1):    
    for i in dfFac.index:
        dfFac.loc[i,'政策']=None
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
        for pathway in ['政策']:
            dfT=df1[df1[pathway]>year]
            UIDList=dfT.index.tolist()
            factor=dfSce.loc[year,pathway]/(getProdAbi(UIDList,year)*12/100000000)
            dfFac.loc[year,pathway]=factor
    return dfFac


###  价格变动和累积成本

def CostCalc(dfPath,dfF):
    #writer=pd.ExcelWriter(r'煤炭退出/退出/临时结果/成本.xlsx',engine='openpyxl')
    pathway=['政策']
    for scenario in pathway:
        Years=list(np.arange(2020,2051,1))
        #df=pd.DataFrame(index=Years,columns=['SAMPLE','SAMPLE成本'])
        
        for sce in [retire]:
            hisCost=0
            for year in Years:
                Cost=0
                dfT=dfPath[dfPath[scenario]>year]
                UIDlist=dfT.index.tolist()
                for UID in UIDlist:
                    temp=dfPro.loc[UID,2019]
                    if not pd.isna(temp):
                        Cost=Cost+dfTot.loc[UID,'价格']*temp*dfF.loc[year,scenario]*12#元
                hisCost=Cost+hisCost
                # df.loc[year,sce+'成本']=hisCost/100000000#亿元  
                # df.loc[year,sce]=Cost/(dfCsm.loc[year,scenario]*100000000)#元
    hisCost=hisCost/100000000#亿元  
    return hisCost
        #df.to_excel(writer,sheet_name=scenario)
    #writer.save()
    #return

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
