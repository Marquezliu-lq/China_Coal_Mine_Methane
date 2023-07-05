# -*- coding: utf-8 -*-
"""
Created on Mon Nov  8 09:40:37 2021

@author: MarqueLiu
"""
'''计算煤炭推出路径下的排放情况'''
'''基本是借鉴前面的代码，更改了：不考虑回收利用'''


import pandas as pd
import random
import math
import numpy as np
import copy
import preExit

pathTot=r'煤炭退出/退出/所有煤矿总表.xlsx'
pathPro=r'所有煤矿产量.xlsx'
pathOth=r'其他信息.xlsx'
pathExit=r'煤炭退出/退出/退出顺序.xlsx'
pathFac=r'煤炭退出/退出/校准因子.xlsx'
pathUti=r'煤炭退出/退出/回收利用率.xlsx'

source=['井工','露天','矿后','AMM','CMM利用','AMM利用']

dfTotOri=pd.read_excel(pathTot,0,index_col='UID')#所有信息总表
dfPro=pd.read_excel(pathPro,0,index_col='UID')#产量表
dfPro.fillna(0,inplace=True)
dfAbanEf=pd.read_excel(pathOth,sheet_name='2011各行政区排放因子',index_col='行政区')
dfAbanNum=pd.read_excel(pathOth,sheet_name='2011前废弃',index_col='行政区')
dfUtiRate=pd.read_excel(pathUti,0,index_col='年份')
#常量部分
factor=0.67
citys=list(dfTotOri['行政区'].unique())
timeSeries=list(range(2020,2051,1))
b=2.017
Df=0.672#水淹
Dd=0.302#干井
aveProd=37919#unit t  2011前废弃矿井平均产能(共退出5亿吨产能)
resEfH=3#Unit m³/t  高瓦斯和突出矿井瓦斯矿后排放因子
resEfL=0.94#Unit m³/t  瓦斯矿井瓦斯矿后排放因子
resPit=0.5#Unit m³/t  露天煤矿矿后排放因子
idxList=dfTotOri.index.tolist()#矿井级排放数据使用index


pathResult=r'煤炭退出/退出/临时结果/退出结果0626.xlsx'
writer=pd.ExcelWriter(pathResult,engine='openpyxl')

col=['政策-规模-高','政策-规模-中','政策-规模-低','政策-经济-高','政策-经济-中','政策-经济-低',\
     '政策-排放-高','政策-排放-中','政策-排放-低','强化-规模-高','强化-规模-中','强化-规模-低',\
     '强化-经济-高','强化-经济-中','强化-经济-低','强化-排放-高','强化-排放-中','强化-排放-低',\
     '2度-规模-高','2度-规模-中','2度-规模-低','2度-经济-高','2度-经济-中','2度-经济-低','2度-排放-高','2度-排放-中','2度-排放-低',\
     '1.5度-规模-高','1.5度-规模-中','1.5度-规模-低','1.5度-经济-高','1.5度-经济-中','1.5度-经济-低',\
    '1.5度-排放-高','1.5度-排放-中','1.5度-排放-低']
dfCMM=pd.DataFrame(index=timeSeries,columns=col)
dfPit=pd.DataFrame(index=timeSeries,columns=col)
dfPost=pd.DataFrame(index=timeSeries,columns=col)
dfAMM=pd.DataFrame(index=timeSeries,columns=col)
dfCMMUti=pd.DataFrame(index=timeSeries,columns=col)
dfAMMUti=pd.DataFrame(index=timeSeries,columns=col)
#定义获取产量
def getProdAbi(UIDList,year):
    prod=0
    for UID in UIDList:
        if year<=2019:
            temp=dfPro.loc[UID,year]
        else:
            temp=dfPro.loc[UID,2019]*dfFac.loc[year,pathway]
            
        if not pd.isna(temp):
            prod=prod+temp
    return prod

### 井工部分
def getMineEmi(UID,year):
    '''传入具体某个UID，在此基础上计算排放和'''
    '''返回单位：kg'''
    ef=dfTot.loc[UID,'排放因子']#unit kg/t
    prod=getProdAbi([UID],year)*12#unit t
    Emi=ef*prod
    return Emi#kg

### 露天部分
def getPitEmi(UID,year):
    '''传入具体某个露天煤矿的UID，在此基础上计算排放和'''
    ef=dfTot.loc[UID,'排放因子']#unit kg/t
    prod=getProdAbi([UID],year)*12#unit t   
    Emi=ef*prod
    return Emi#kg
### 废弃部分(注意用蒙特卡洛来多次模拟)
def getAbanEmi(UID,year,isFlood,flag):
    abanYear=dfTot.loc[UID,'废弃时间']
    if abanYear>2011 and not flag:#并非当年废弃，可以-1
        iniEmi=getMineEmi(UID,abanYear-1)
    else:
        iniEmi=getMineEmi(UID,abanYear)
    #flood
    if isFlood:
        emi=iniEmi*math.exp(-(year-abanYear)*Df)
    else:
        emi=iniEmi*math.pow(1+b*Dd*(year-abanYear),(-1/b))
    return emi

#2011前废弃煤矿排放
def iniAbanEmi(year):
    '''计算某个特定行政区，2011前废弃煤矿的排放'''
    '''针对整个行政区的调用函数'''
    totEmi=0
    for city in citys+['广东']:
        aveFac=dfAbanEf.loc[city,'排放因子']
        iniEmi=aveProd*aveFac#单个的初始排放
        for i in range(2005,2011):
            num=dfAbanNum.loc[city,i]
            #干井
            emiDry=num*iniEmi*math.pow(1+b*Dd*(year-i),(-1/b))*0.2
            #水淹
            emiFlood=num*iniEmi*math.exp(-(year-i)*Df)*0.8
            
            totEmi=totEmi+emiFlood+emiDry
    return totEmi

#矿后活动
def getResEmi(UID,year):
    '''传入具体煤矿的UID，计算对应的矿后活动排放'''
    '''注意只计算井工矿的矿后活动'''
    judge=dfTot.loc[UID,'是否露天']
    if judge:
        Efs=resPit*0.67
    else:
        flag=dfTot.loc[UID,'瓦斯等级']
        if flag=='瓦斯':
            Efs=resEfL*0.67
        else:
            Efs=resEfH*0.67
    prod=getProdAbi([UID],year)
    emi=prod*12*Efs
    return emi

#整理退出时间
def renewExitPath(road,retire):
    global dfTot
    dfExit=pd.read_excel(pathExit,retire,index_col='UID')
    for UID in idxList:
        if dfTotOri.loc[UID,'废弃时间']==10000:
            if UID in dfExit.index:
                dfTot.loc[UID,'废弃时间']=dfExit.loc[UID,road]

def getExitEmi(exitRule,scenario,UtiDegree):
    title=scenario+'-'+exitRule+'-'+UtiDegree
    global pathway
    pathway=scenario
    global dfTot
    dfTot=copy.deepcopy(dfTotOri)
    global dfFac
    dfFac=pd.read_excel(pathFac,exitRule,index_col='年份')
    renewExitPath(scenario,exitRule)#更新时间
    
    # dfMineLevel=pd.DataFrame(index=dfTot.index.tolist(),columns=['kg'])
    
    timeSeries=[2030]
    for year in timeSeries:
        dfTemp=dfTot[(dfTot['新建时间']<=year)&(dfTot['废弃时间']>year)&(dfTot['是否露天']==0)]#筛选每年井工表
        UIDlist=dfTemp.index.tolist()
        emi=0
        for UID in UIDlist:
            emiTemp=getMineEmi(UID, year)#返回单位为 kg
            emi=emi+emiTemp
            
            # dfMineLevel.loc[UID,'kg']=emiTemp
            
        dfCMM.loc[year,title]=emi/1000000000#unit Mt
    #再计算露天部分
    for year in timeSeries:
        dfTemp=dfTot[(dfTot['新建时间']<=year)&(dfTot['废弃时间']>year)&(dfTot['是否露天']==1)]#筛选每年露天表
        UIDlist=dfTemp.index.tolist()
        emi=0
        for UID in UIDlist:
            emiTemp=getPitEmi(UID, year)
            emi=emi+emiTemp
            
            # dfMineLevel.loc[UID,'kg']=emiTemp
            
        dfPit.loc[year,title]=emi/1000000000#unit Mt
    #计算废弃部分
    #以下为矿井级别
    for year in timeSeries:
        dfTemp=dfTot[(dfTot['废弃时间']<=year)&(dfTot['是否露天']==0)]#筛选废弃表
        UIDlist=dfTemp.index.tolist()
        emi=0
        for UID in UIDlist:
            flag=0#代表不是当年废弃
            if dfTemp.loc[UID,'新建时间']==dfTemp.loc[UID,'废弃时间']:
                flag=1
            isFlood=0
            if dfTemp.loc[UID,'是否水淹']==1:
                isFlood=1
            # elif dfTemp.loc[UID,'废弃时间']>2019:
            #     isSealed=1
            # elif dfTemp.loc[UID,'Sealed']<=year:
            #     isSealed=1
            emiTemp=getAbanEmi(UID, year,isFlood,flag)
            emi=emi+emiTemp
            
            # dfMineLevel.loc[UID,'kg']=emiTemp
            
        dfAMM.loc[year,title]=(iniAbanEmi(year)+emi)/1000000000#unit Mt
    #计算矿后活动
    for year in timeSeries:
        dfTemp=dfTot[(dfTot['新建时间']<=year)&(dfTot['废弃时间']>year)]#筛选每年井工+露天表
        UIDlist=dfTemp.index.tolist()
        emi=0
        for UID in UIDlist:
            emiTemp=getResEmi(UID, year)
            emi=emi+emiTemp
            
            # dfMineLevel.loc[UID,'kg']+=emiTemp
            
        dfPost.loc[year,title]=emi/1000000000#unit Mt
    # CMM-Utilization
    for year in timeSeries:
        dfCMMUti.loc[year,title]=dfCMM.loc[year,title]*(dfUtiRate.loc[year,UtiDegree+'抽采率']*dfUtiRate.loc[year,UtiDegree+'利用率'])
    ##  AMM-Utilizaition
    for year in timeSeries:
        dfT=dfTot[(dfTot['是否AMM']==1)&(dfTot['废弃时间']<=year)].copy()
        dfT.sort_values(by=['选择顺序'],ascending=[True],inplace=True)
        num=dfUtiRate.loc[year,UtiDegree+'AMM']
        count=0
        uti=0
        for UID in dfT.index:
            if count>=num:
                break
            if dfT.loc[UID,'废弃时间']==10000:
                continue
            count=count+1
            flag=0
            if dfTot.loc[UID,'新建时间']==dfTot.loc[UID,'废弃时间']:
                flag=1
            uti=uti+getAbanEmi(UID,year,0,flag)*(0.75*0.65)
        dfAMMUti.loc[year,title]=uti/1000000000#Tg
    # return dfMineLevel

if __name__=='__main__':
    preExit.ExitOrder()#先退出顺序
    preExit.CaliFactor()#校准因子
    for exitRule in ['规模','经济','排放']:
        for scenario in ['政策','强化','2度','1.5度']:
            for UtiDegree in ['高','中','低']:
                getExitEmi(exitRule,scenario,UtiDegree)
    with pd.ExcelWriter(pathResult) as writer:
        dfCMM.to_excel(writer,sheet_name='CMM')
        dfPit.to_excel(writer,sheet_name='Pit')
        dfPost.to_excel(writer,sheet_name='Post')
        dfAMM.to_excel(writer,sheet_name='AMM')
        dfCMMUti.to_excel(writer,sheet_name='CMMUti')
        dfAMMUti.to_excel(writer,sheet_name='AMMUti')
    preExit.CostCalc()#计算成本
    preExit.MarCurve()#边际曲线


# ## 计算不同情景下2030的排放分布   对应getExitEmi函数中被注释的部分
# if __name__=='__main__':
#     preExit.ExitOrder()#先退出顺序
#     preExit.CaliFactor()#校准因子
#     for exitRule in ['规模','经济','排放']:
#         for scenario in ['1.5度']:
#                 df=getExitEmi(exitRule,scenario,'高')
#                 df.to_excel(exitRule+'2030'+'.xlsx',sheet_name='sheet1'



    
    
    