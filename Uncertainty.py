# -*- coding: utf-8 -*-
"""
Created on Mon Dec  6 18:16:00 2021

@author: MarqueLiu
"""
import pandas as pd
import random
import math
import numpy as np
import copy

path=r'推断结果/'
pathTot=r'所有煤矿总表.xlsx'
pathPro=r'所有煤矿产量.xlsx'
pathOth=r'其他信息.xlsx'
pathRes=r'敏感性分析-新抽样方式.xlsx'

dfTot_O=pd.read_excel(pathTot,0,index_col='UID')#mine-specific information
dfPro_O=pd.read_excel(pathPro,0,index_col='UID')#mine-specific annual production
dfPro_O.fillna(0,inplace=True)
dfAbanEf=pd.read_excel(pathOth,sheet_name='2011各行政区排放因子',index_col='行政区')
dfAbanNum=pd.read_excel(pathOth,sheet_name='2011前废弃',index_col='行政区')
dfRevOri_O=pd.read_excel(pathOth,sheet_name='回收利用',index_col='行政区')

#number of sampling-先2000次
sampleNum=2000

#the uncertainty rage of given parameters
prod_U=0.1
util_U=0.1
EFmine_U=0.4#
EFsur_U=0.5
EFpost_U=0.5

flood_U=0.5
abanRate_U=0.5
para_U=0.4
hisAban_U=0.5
rev_U=0.1

###constant value
factor=0.67
citys=dfTot_O['行政区'].unique()
timeSeries=list(range(2011,2020))
b_O=2.017
D1_O=0.672
D2_O=0.302
floodRate_O=0.8
aveProd_O=37919
resEfH_O=3
resEfL_O=0.94
resPit_O=0.5
idxList=dfTot_O.index.tolist()
###dataframe of parameters after sampling
dfTot=copy.deepcopy(dfTot_O)
dfPro=copy.deepcopy(dfPro_O)
dfRevOri=copy.deepcopy(dfRevOri_O)

b=b_O
D1=D1_O
D2=D2_O
aveProd=aveProd_O
resEfH=resEfH_O
resEfL=resEfL_O
resPit=resPit_O
floodRate=floodRate_O

def getProdAbi(UIDList,year):
    prod=0
    for UID in UIDList:
        temp=dfPro.loc[UID,year]
        if not pd.isna(temp):
            prod=prod+dfPro.loc[UID,year]
    #Unit:t/month
    return prod

#CMM
def getMineEmi(UID,year):
    ef=dfTot.loc[UID,'排放因子']#unit kg/t
    prod=dfPro.loc[UID,year]*12#unit t
    Emi=ef*prod#kg
    return Emi

#Open-pit miens
def getPitEmi(UID,year):
    ef=dfTot.loc[UID,'排放因子']#unit kg/t
    prod=dfPro.loc[UID,year]*12#unit t   
    Emi=ef*prod
    return Emi
# AMM
def getAbanEmi(UID,year,isFlood,flag):

    abanYear=dfTot.loc[UID,'新废弃时间']
    if abanYear>2011 and not flag:
        iniEmi=getMineEmi(UID,abanYear-1)
    else:
        iniEmi=getMineEmi(UID,abanYear)
    #dry
    if not isFlood:
        emi=iniEmi*math.pow(1+b*D2*(year-abanYear),(-1/b))
    else:
        emi=iniEmi*math.exp(-(year-abanYear)*D1)
    return emi#kg

#AMM from mines closed before 2011
def iniAbanEmi(year):
    totEmi=0
    for city in list(citys)+['广东']:
        aveFac=dfAbanEf.loc[city,'排放因子']
        iniEmi=aveProd*aveFac
        for i in range(2005,2011):
            num=dfAbanNum.loc[city,i]
            #水淹
            emiDry=num*iniEmi*math.pow(1+b*D2*(year-i),(-1/b))*(1-floodRate)#
            #干井
            emiFlood=num*iniEmi*math.exp(-(year-i)*D1)*(floodRate)#
            totEmi=totEmi+emiFlood+emiDry#kg
    return totEmi#kg

#Post-mining
def getResEmi(UID,year):
    judge=dfTot.loc[UID,'是否露天']
    if judge:
        Efs=resPit*0.67
    else:
        flag=dfTot.loc[UID,'瓦斯等级']
        if flag=='瓦斯':
            Efs=resEfL*0.67
        else:
            Efs=resEfH*0.67
    prod=dfPro.loc[UID,year]
    emi=prod*12*Efs
    return emi

###calculating emissions with sampled parameters
def get_Uncertain():
    for year in timeSeries:
        dfTemp=dfTot[(dfTot['新建时间']<=year)&(dfTot['废弃时间']>year)&(dfTot['是否露天']==0)]#筛选每年井工表
        UIDlist=dfTemp.index.tolist()
        emi=0
        for UID in UIDlist:
            emiTemp=getMineEmi(UID, year)
            emi=emi+emiTemp#kg
        dfMineRes.loc[count,year]=emi/1000000000#unit Tg
    for year in timeSeries:
        dfTemp=dfTot[(dfTot['新建时间']<=year)&(dfTot['废弃时间']>year)&(dfTot['是否露天']==1)]
        UIDlist=dfTemp.index.tolist()
        emi=0
        for UID in UIDlist:
            emiTemp=getPitEmi(UID, year)
            emi=emi+emiTemp
        dfPitRes.loc[count,year]=emi/1000000000#unit Tg
    for year in timeSeries:
        dfTemp=dfTot[(dfTot['新废弃时间']<=year)&(dfTot['是否露天']==0)]
        UIDlist=dfTemp.index.tolist()
        emi=0
        for UID in UIDlist:
            flag=0
            if dfTemp.loc[UID,'新建时间']==dfTemp.loc[UID,'新废弃时间']:
                flag=1
            emiTemp=getAbanEmi(UID, year,dfTemp.loc[UID,'是否水淹'],flag)
            emi=emi+emiTemp
        dfAbanRes.loc[count,year]=(emi+iniAbanEmi(year))/1000000000#Tg
    for year in timeSeries:
        dfTemp=dfTot[(dfTot['新建时间']<=year)&(dfTot['废弃时间']>year)]
        UIDlist=dfTemp.index.tolist()
        emi=0
        for UID in UIDlist:
            emiTemp=getResEmi(UID, year)
            emi=emi+emiTemp
        dfResRes.loc[count,year]=emi/1000000000#unit Tg
    for year in timeSeries:
        dfRev.loc[count,year]=dfRevOri.loc['全国利用',year]#Tg
    for year in timeSeries:
        arr=np.array([dfMineRes.loc[count,year],dfPitRes.loc[count,year],dfAbanRes.loc[count,year],dfResRes.loc[count,year],-1*dfRev.loc[count,year]])
        dfAll.loc[count,year]=np.nansum(arr)

def initial():
    global dfMineRes
    dfMineRes=pd.DataFrame(index=idx,columns=timeSeries+['单位'])
    dfMineRes['单位']='Tg'
    global dfPitRes
    dfPitRes=pd.DataFrame(index=idx,columns=timeSeries+['单位'])
    dfPitRes['单位']='Tg'
    global dfAbanRes
    dfAbanRes=pd.DataFrame(index=idx,columns=timeSeries+['单位'])
    dfAbanRes['单位']='Tg'
    global dfResRes
    dfResRes=pd.DataFrame(index=idx,columns=timeSeries+['单位'])
    dfResRes['单位']='Tg'
    global dfRev
    dfRev=pd.DataFrame(index=idx,columns=timeSeries+['单位'])
    dfRev['单位']='Tg'
    global dfAll
    dfAll=pd.DataFrame(index=idx,columns=timeSeries+['单位'])
    dfAll['单位']='Tg'
    return

if __name__=='__main__':
    idx=range(0,sampleNum)
    initial()
    global count
    for count in idx:
        b=np.random.uniform((1-para_U)*b_O,(1+para_U)*b_O)
        D1=np.random.uniform((1-para_U)*D1_O,(1+para_U)*D1_O)
        D2=np.random.uniform((1-para_U)*D2_O,(1+para_U)*D2_O)
        aveProd=np.random.uniform((1-hisAban_U)*aveProd_O,(1+hisAban_U)*aveProd_O)
        resEfH=np.random.normal(resEfH_O,((EFpost_U)*resEfH_O)/3)
        resEfL=np.random.normal(resEfL_O,((EFpost_U)*resEfL_O)/3)
        resPit=np.random.normal(resPit_O,((EFpost_U)*resPit_O)/3)
        floodRate=np.random.normal(floodRate_O,((flood_U)*floodRate_O)/3)
        while floodRate>1:
            floodRate=np.random.normal(floodRate_O,((flood_U)*floodRate_O)/3)
        
        abanRate=np.random.uniform((1-abanRate_U)*0.5,(1+abanRate_U)*0.5)
        
        #EFminefac=np.random.uniform((1-EFmine_U),(1+EFmine_U))
        EFsurfac=np.random.normal(1,(EFsur_U)/3)
        prodfac=np.random.uniform((1-prod_U),(1+prod_U))
        UIDlist=dfTot_O.index.tolist()
        for UID in UIDlist:
            if random.random() < floodRate:
                dfTot.loc[UID,'是否水淹']=1
            else:
                dfTot.loc[UID,'是否水淹']=0
            EF=dfTot_O.loc[UID,'排放因子']
            if dfTot_O.loc[UID,'是否露天']==0:
                dfTot.loc[UID,'排放因子']=EF*np.random.normal((1),(EFmine_U)/3)
            else:
                dfTot.loc[UID,'排放因子']=EF*EFsurfac
            if dfTot_O.loc[UID,'废弃时间'] in [2013,2014]:
                if random.random() < abanRate:
                    dfTot.loc[UID,'新废弃时间']=2013
                else:
                    dfTot.loc[UID,'新废弃时间']=2014
            for year in timeSeries:
                prod=dfPro_O.loc[UID,year]
                if not pd.isna(prod):
                    dfPro.loc[UID,year]=prod*prodfac
        for year in timeSeries:
            rev=dfRevOri_O.loc['全国利用',year]
            dfRevOri.loc['全国利用',year]=np.random.uniform((1-rev_U)*rev,(1+rev_U)*rev)
        get_Uncertain()
         
    with pd.ExcelWriter(pathRes) as writer:
        dfMineRes.to_excel(writer,sheet_name='井工')
        dfPitRes.to_excel(writer,sheet_name='露天')
        dfAbanRes.to_excel(writer,sheet_name='废弃')
        dfResRes.to_excel(writer,sheet_name='矿后')
        dfRev.to_excel(writer,sheet_name='回收利用')
        dfAll.to_excel(writer,sheet_name='总排放')




        
        
