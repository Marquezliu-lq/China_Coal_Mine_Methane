# -*- coding: utf-8 -*-
"""
Created on Sat Dec 25 21:14:25 2021
#Required code for the simulation of mine closure strategies
#The pre-required code for the 'mine_exit.py'

@author: MarqueLiu
"""

import pandas as pd
import numpy as np
from openpyxl import load_workbook

### refresh the timesheet of mine closure
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
        #clear sheet
        for a in dfOut.columns:
            for i in dfOut.index:
                dfOut.loc[i,a]=None
        years=range(2020,2051,1)
        #ranke the order of retirement
        df1.sort_values(by=[retire],ascending=[False],inplace=True)
        
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
        dfOut.fillna(value=10000,inplace=True)#unretirement mines
        dfOut.to_excel(writer,sheet_name=retire)
    writer.save()
    return

##calibrating annual production of that year
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


###calculating average production costs in the future and the cumulative production costs
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
