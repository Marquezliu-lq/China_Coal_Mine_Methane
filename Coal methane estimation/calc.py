# -*- coding: utf-8 -*-
"""
Created on Tue Aug 17 18:46:58 2021
#The core code for coal methane emission estimates

@author: 刘强
"""
'''code for emissiosn calculation'''
import pandas as pd
import random
import math
import numpy as np

path=r'推断结果/'
pathTot=r'所有煤矿总表.xlsx'
pathPro=r'所有煤矿产量.xlsx'
pathOth=r'其他信息.xlsx'
pathRes=r'国家和行政级别排放结果.xlsx'
pathLocRes=r'矿井级别排放结果.xlsx'

dfTot=pd.read_excel(pathTot,0,index_col='UID')#mine-specific information
dfPro=pd.read_excel(pathPro,0,index_col='UID')#mine-specific production
dfPro.fillna(0,inplace=True)
# dfRawCoal=pd.read_excel(pathOth,sheet_name='校准后原煤产量',index_col='行政区')#原煤产量表（包括露天）
# dfPitCoal=pd.read_excel(pathOth,sheet_name='校准后露天产量',index_col='行政区')#露天产量表
# dfMineCoal=pd.read_excel(pathOth,sheet_name='校准后井工产量',index_col='行政区')#井工产量表
dfAbanEf=pd.read_excel(pathOth,sheet_name='2011各行政区排放因子',index_col='行政区')
dfAbanNum=pd.read_excel(pathOth,sheet_name='2011前废弃',index_col='行政区')
dfRevOri=pd.read_excel(pathOth,sheet_name='回收利用',index_col='行政区')

#constant value
factor=0.67
citys=dfTot['行政区'].unique()
timeSeries=list(range(2011,2020))
b=2.017#2.017
Df=0.672#flooding mines
Dd=0.302#dry mines
aveProd=37919#unit t  average production capacity of mines closed before 2011
resEfH=3#Unit m³/t  post-mining EF of outburst and high gas content mines
resEfL=0.94#Unit m³/t  post-mining EF of low gas content mines
resPit=0.5#Unit m³/t  post-mining EF of open-pit mines
idxList=dfTot.index.tolist()#the list of UID(unique identification digit)
floodRate=0.8
#get annual production
def getProdAbi(UIDList,year):
    prod=0
    for UID in UIDList:
        temp=dfPro.loc[UID,year]
        if not pd.isna(temp):
            prod=prod+temp#t/month
    return prod

### CMM calculation
def getMineEmi(UID,year):
    ef=dfTot.loc[UID,'排放因子']#unit kg/t
    prod=dfPro.loc[UID,year]*12#unit t
    Emi=ef*prod
    return Emi

### Open-pit emissions
def getPitEmi(UID,year):
    ef=dfTot.loc[UID,'排放因子']#unit kg/t
    prod=dfPro.loc[UID,year]*12#unit t   
    Emi=ef*prod
    return Emi
# AMM emissions
def getAbanEmi(UID,year,isFlood,flag):
    abanYear=dfTot.loc[UID,'废弃时间']
    if abanYear>2011 and not flag:
        iniEmi=getMineEmi(UID,abanYear-1)
    else:
        iniEmi=getMineEmi(UID,abanYear)
    #dry
    if not isFlood:
        emi=iniEmi*math.pow(1+b*Dd*(year-abanYear),(-1/b))
    else:
        emi=iniEmi*math.exp(-(year-abanYear)*Df)
    return emi

#AMM from mines closed before 2011
def iniAbanEmi(year,city):
    aveFac=dfAbanEf.loc[city,'排放因子']
    iniEmi=aveProd*aveFac
    totEmi=0
    for i in range(2005,2011):
        num=dfAbanNum.loc[city,i]
        #dry
        emiDry=num*iniEmi*math.pow(1+b*Dd*(year-i),(-1/b))*((1-floodRate))
        #flooded
        emiFlood=num*iniEmi*math.exp(-(year-i)*Df)*floodRate
        
        totEmi=totEmi+emiFlood+emiDry
    return totEmi

#Post-mining emissions
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

#recovery and utilization
def getRevEmi(UID,year):
    city=dfTot.loc[UID,'行政区']
    global dfMineRes
    totEmi=dfMineRes.loc[city,year]#unit 吨
    fac=dfRevOri.loc[city,year]/totEmi
    oriEmi=getMineEmi(UID, year)
    return oriEmi*fac


if __name__=='__main__':
    global dfMineRes
    #provincial and national emissions
    dfMineRes=pd.DataFrame(index=citys,columns=timeSeries+['单位'])
    dfMineRes['单位']='吨'# unit t
    #mine-level emissions
    dfMine=pd.DataFrame(index=idxList,columns=timeSeries+['lng','lat','单位'])
    dfMine['lng']=dfTot['经度_百度_POI']
    dfMine['lat']=dfTot['纬度_百度_POI']
    dfMine['单位']='百吨'
    for year in timeSeries:
        dfTemp=dfTot[(dfTot['新建时间']<=year)&(dfTot['废弃时间']>year)&(dfTot['是否露天']==0)]#sheet of in-production mines in that year
        for city in citys:
            dfT=dfTemp[dfTemp['行政区']==city]
            UIDlist=dfT.index.tolist()
            emi=0
            for UID in UIDlist:
                emiTemp=getMineEmi(UID, year)#unit kg
                emi=emi+emiTemp
                dfMine.loc[UID,year]=emiTemp/100000#unit 100t
            dfMineRes.loc[city,year]=emi/1000#unit t
        dfMineRes.loc['全国',year]=dfMineRes[year].sum()/1000000#unit Mt或Tg
        dfMineRes.loc['全国','单位']='Tg'
    #open-pit mines
    dfPitRes=pd.DataFrame(index=citys,columns=timeSeries+['单位'])
    dfPitRes['单位']='吨'
    #mine-level
    dfPit=pd.DataFrame(index=idxList,columns=timeSeries+['lng','lat','单位'])
    dfPit['lng']=dfTot['经度_百度_POI']
    dfPit['lat']=dfTot['纬度_百度_POI']
    dfPit['单位']='百吨'
    for year in timeSeries:
        dfTemp=dfTot[(dfTot['新建时间']<=year)&(dfTot['废弃时间']>year)&(dfTot['是否露天']==1)]#sheet of open-pit mines
        for city in citys:
            dfT=dfTemp[dfTemp['行政区']==city]
            UIDlist=dfT.index.tolist()
            emi=0
            for UID in UIDlist:
                emiTemp=getPitEmi(UID, year)
                emi=emi+emiTemp
                dfPit.loc[UID,year]=emiTemp/100000#unit 100t
            dfPitRes.loc[city,year]=emi/1000#unit 吨
        dfPitRes.loc['全国',year]=dfPitRes[year].sum()/1000000#unit Mt或Tg
        dfPitRes.loc['全国','单位']='Tg'
    #AMM emissions
    dfAbanRes=pd.DataFrame(index=citys,columns=timeSeries+['单位'])
    dfAbanRes['单位']='吨'
    #mine-level
    dfAban=pd.DataFrame(index=idxList,columns=timeSeries+['lng','lat','单位'])
    dfAban['lng']=dfTot['经度_百度_POI']
    dfAban['lat']=dfTot['纬度_百度_POI']
    dfAban['单位']='百吨'
    for year in timeSeries:
        dfTemp=dfTot[(dfTot['废弃时间']<=year)&(dfTot['是否露天']==0)]#sheet of abandoned mines
        for city in citys:
            dfT=dfTemp[dfTemp['行政区']==city]
            UIDlist=dfT.index.tolist()
            emi=0
            for UID in UIDlist:
                flag=0
                if dfT.loc[UID,'新建时间']==dfT.loc[UID,'废弃时间']:
                    flag=1
                emiTemp=getAbanEmi(UID, year,dfT.loc[UID,'是否水淹'],flag)
                emi=emi+emiTemp
                dfAban.loc[UID,year]=emiTemp/100000#unit 100t
            dfAbanRes.loc[city,year]=(emi+iniAbanEmi(year,city))/1000#unit t
        dfAbanRes.loc['广东',year]=iniAbanEmi(year,'广东')/1000#unit t
        dfAbanRes.loc['全国',year]=dfAbanRes[year].sum()/1000000#unit Mt或Tg
        dfAbanRes.loc['全国','单位']='Tg'
    #post mining
    dfResRes=pd.DataFrame(index=citys,columns=timeSeries+['单位'])
    dfResRes['单位']='吨'
    #mine-level
    dfRes=pd.DataFrame(index=idxList,columns=timeSeries+['lng','lat','单位'])
    dfRes['lng']=dfTot['经度_百度_POI']
    dfRes['lat']=dfTot['纬度_百度_POI']
    dfRes['单位']='百吨'
    for year in timeSeries:
        dfTemp=dfTot[(dfTot['新建时间']<=year)&(dfTot['废弃时间']>year)]#inproduction underground and open-pit mines
        for city in citys:
            dfT=dfTemp[dfTemp['行政区']==city]
            UIDlist=dfT.index.tolist()
            emi=0
            for UID in UIDlist:
                emiTemp=getResEmi(UID, year)
                emi=emi+emiTemp
                dfRes.loc[UID,year]=emiTemp/100000 #100t
            dfResRes.loc[city,year]=emi/1000#unit t
        dfResRes.loc['全国',year]=dfResRes[year].sum()/1000000#unit Mt或Tg
        dfResRes.loc['全国','单位']='Tg'
    # recovery and utilization
    dfRev=pd.DataFrame(index=idxList,columns=timeSeries+['lng','lat','单位'])
    dfRev['lng']=dfTot['经度_百度_POI']
    dfRev['lat']=dfTot['纬度_百度_POI']
    dfRev['单位']='百吨'
    for year in timeSeries:
        dfTemp=dfTot[(dfTot['新建时间']<=year)&(dfTot['废弃时间']>year)&(dfTot['是否露天']==0)]#in-production underground mines
        for city in citys:
            dfT=dfTemp[dfTemp['行政区']==city]
            UIDlist=dfT.index.tolist()
            emi=0
            for UID in UIDlist:
                emiTemp=getRevEmi(UID, year)
                dfRev.loc[UID,year]=emiTemp/100000 #100t
    #Overall emissions
    dfAll=pd.DataFrame(index=idxList,columns=timeSeries+['lng','lat','单位'])
    dfAll['lng']=dfTot['经度_百度_POI']
    dfAll['lat']=dfTot['纬度_百度_POI']
    dfAll['单位']='百吨'
    for year in timeSeries:
        for UID in idxList:
            arr=np.array([dfMine.loc[UID,year],dfPit.loc[UID,year],dfAban.loc[UID,year],dfRes.loc[UID,year],-dfRev.loc[UID,year]])
            dfAll.loc[UID,year]=np.nansum(arr)
            
    #written
    with pd.ExcelWriter(pathRes) as writer:
        dfMineRes.to_excel(writer,sheet_name='井工')
        dfPitRes.to_excel(writer,sheet_name='露天')
        dfAbanRes.to_excel(writer,sheet_name='废弃')
        dfResRes.to_excel(writer,sheet_name='矿后')
        dfRevOri.to_excel(writer,sheet_name='利用')
    with pd.ExcelWriter(pathLocRes) as writer:
        dfMine.to_excel(writer,sheet_name='井工')
        dfPit.to_excel(writer,sheet_name='露天')
        dfAban.to_excel(writer,sheet_name='废弃')
        dfRes.to_excel(writer,sheet_name='矿后')
        dfRev.to_excel(writer,sheet_name='回收利用')
        dfAll.to_excel(writer,sheet_name='总排放')
