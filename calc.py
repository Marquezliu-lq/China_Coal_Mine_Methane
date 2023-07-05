# -*- coding: utf-8 -*-
"""
Created on Tue Aug 17 18:46:58 2021

@author: 刘强
"""

'''计算排放的主体程序'''
'''0906版本更新,增加了露天矿后的计算'''
'''0908修改了废弃排放的计算'''
'''0911增加了总排放计算'''

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

dfTot=pd.read_excel(pathTot,0,index_col='UID')#所有信息总表
dfPro=pd.read_excel(pathPro,0,index_col='UID')#产量表
dfPro.fillna(0,inplace=True)
# dfRawCoal=pd.read_excel(pathOth,sheet_name='校准后原煤产量',index_col='行政区')#原煤产量表（包括露天）
# dfPitCoal=pd.read_excel(pathOth,sheet_name='校准后露天产量',index_col='行政区')#露天产量表
# dfMineCoal=pd.read_excel(pathOth,sheet_name='校准后井工产量',index_col='行政区')#井工产量表
dfAbanEf=pd.read_excel(pathOth,sheet_name='2011各行政区排放因子',index_col='行政区')
dfAbanNum=pd.read_excel(pathOth,sheet_name='2011前废弃',index_col='行政区')
dfRevOri=pd.read_excel(pathOth,sheet_name='回收利用',index_col='行政区')

#常量部分
factor=0.67
citys=dfTot['行政区'].unique()
timeSeries=list(range(2011,2020))
b=2.017#2.017
Df=0.672#水淹0.672
Dd=0.302#干井0.302
aveProd=37919#unit t  2011前废弃矿井平均产能(共退出5亿吨产能)
resEfH=3#Unit m³/t  高瓦斯和突出矿井瓦斯矿后排放因子
resEfL=0.94#Unit m³/t  瓦斯矿井瓦斯矿后排放因子
resPit=0.5#Unit m³/t  露天煤矿矿后排放因子
idxList=dfTot.index.tolist()#矿井级排放数据使用index
floodRate=0.8
#定义获取产量
def getProdAbi(UIDList,year):
    prod=0
    for UID in UIDList:
        temp=dfPro.loc[UID,year]
        if not pd.isna(temp):
            prod=prod+temp#t/month
    return prod

### 井工部分
def getMineEmi(UID,year):
    '''传入具体某个UID，在此基础上计算排放和'''
    '''返回单位：kg'''
    ef=dfTot.loc[UID,'排放因子']#unit kg/t
    prod=dfPro.loc[UID,year]*12#unit t
    Emi=ef*prod
    return Emi

### 露天部分
def getPitEmi(UID,year):
    '''传入具体某个露天煤矿的UID，在此基础上计算排放和'''
    ef=dfTot.loc[UID,'排放因子']#unit kg/t
    prod=dfPro.loc[UID,year]*12#unit t   
    Emi=ef*prod
    return Emi
### 废弃部分(注意用蒙特卡洛来多次模拟)

#在外部蒙特卡洛
def getAbanEmi(UID,year,isFlood,flag):
    '''传入UID和当前年份，计算某个煤矿当前的废弃排放'''
    '''注意80%的水淹比例'''
    '''每次调用前要先在外部蒙特卡洛判断是否水淹'''
    '''0--未被水淹；1--已被水淹'''
    '''假设没有密封这种技术使用'''
    '''还要注意2011年以前的部分要加进来'''
    '''针对单个煤矿的调用函数'''
    '''废弃当年也有排放，t=0，不要忽视'''
    '''返回单位kg'''
    abanYear=dfTot.loc[UID,'废弃时间']
    if abanYear>2011 and not flag:#并非当年废弃，可以-1
        iniEmi=getMineEmi(UID,abanYear-1)
    else:
        iniEmi=getMineEmi(UID,abanYear)
    #dry
    if not isFlood:
        emi=iniEmi*math.pow(1+b*Dd*(year-abanYear),(-1/b))
    else:
        emi=iniEmi*math.exp(-(year-abanYear)*Df)
    return emi

#2011前废弃煤矿排放
def iniAbanEmi(year,city):
    '''计算某个特定行政区，2011前废弃煤矿的排放'''
    '''针对整个行政区的调用函数'''
    '''返回单位kg'''
    aveFac=dfAbanEf.loc[city,'排放因子']
    iniEmi=aveProd*aveFac#单个的初始排放
    totEmi=0
    for i in range(2005,2011):
        num=dfAbanNum.loc[city,i]
        #dry
        emiDry=num*iniEmi*math.pow(1+b*Dd*(year-i),(-1/b))*((1-floodRate))
        #水淹
        emiFlood=num*iniEmi*math.exp(-(year-i)*Df)*floodRate
        
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
    prod=dfPro.loc[UID,year]
    emi=prod*12*Efs
    return emi

#回收利用
def getRevEmi(UID,year):
    '''传入UID，根据年度数据调整具体到煤矿的回收利用'''
    city=dfTot.loc[UID,'行政区']
    global dfMineRes
    totEmi=dfMineRes.loc[city,year]#unit 吨
    fac=dfRevOri.loc[city,year]/totEmi
    oriEmi=getMineEmi(UID, year)
    return oriEmi*fac

#定义初始化程序
def init():
    yearList=[2011,2012,2013,2014,2015,2016,2017,2018,2019]
    for year in yearList:
        dfTemp=dfTot[(dfTot['新建时间']<=year)&(dfTot['废弃时间']>=year)&(dfTot['是否露天']==0)]
        for key in ['突出','高瓦斯','瓦斯']:
            for city in citys:
                dfTemp1=dfTemp[(~pd.isna(dfTemp['排放因子']))&(dfTemp['瓦斯等级']==key)&(dfTemp['行政区']==city)&(~pd.isna(dfTemp['月产煤量']))]#排放因子不为空
                dfTemp2=dfTemp[(pd.isna(dfTemp['排放因子']))&(dfTemp['瓦斯等级']==key)&(dfTemp['行政区']==city)]#排放因子为空
                if dfTemp2.empty:
                    continue
                totEmi=0
                UIDList=dfTemp1.index.tolist()
                for UID in UIDList:
                    totEmi=totEmi+dfTot.loc[UID,'排放因子']*dfPro.loc[UID,year]
                totPro=getProdAbi(UIDList,year)
                if totPro==0:
                    continue
                emiFac=totEmi/totPro
                idxList=dfTemp2.index.tolist()
                for i in idxList:
                    dfTot.loc[i,'排放因子']=emiFac
    dfTot.to_excel(r'所有煤矿总表.xlsx',sheet_name='Sheet1')

# #主体计算程序
if __name__=='__main__':
    '''主体计算程序'''
    #根据情况是否初始化
    #init()#初始化填补空缺Efs
    #先计算井工部分
    global dfMineRes
    #以下为国家级和省级排放计算
    dfMineRes=pd.DataFrame(index=citys,columns=timeSeries+['单位'])
    dfMineRes['单位']='吨'
    #以下为矿井级别的排放计算
    dfMine=pd.DataFrame(index=idxList,columns=timeSeries+['lng','lat','单位'])
    dfMine['lng']=dfTot['经度_百度_POI']
    dfMine['lat']=dfTot['纬度_百度_POI']
    dfMine['单位']='百吨'
    for year in timeSeries:
        dfTemp=dfTot[(dfTot['新建时间']<=year)&(dfTot['废弃时间']>year)&(dfTot['是否露天']==0)]#筛选每年井工表
        for city in citys:
            dfT=dfTemp[dfTemp['行政区']==city]
            UIDlist=dfT.index.tolist()
            emi=0
            for UID in UIDlist:
                emiTemp=getMineEmi(UID, year)#返回单位为 kg
                emi=emi+emiTemp
                dfMine.loc[UID,year]=emiTemp/100000#单位  百吨
            dfMineRes.loc[city,year]=emi/1000#unit 吨
        dfMineRes.loc['全国',year]=dfMineRes[year].sum()/1000000#unit Mt或Tg
        dfMineRes.loc['全国','单位']='Tg'
    #再计算露天部分
    dfPitRes=pd.DataFrame(index=citys,columns=timeSeries+['单位'])
    dfPitRes['单位']='吨'
    #矿井级别
    dfPit=pd.DataFrame(index=idxList,columns=timeSeries+['lng','lat','单位'])
    dfPit['lng']=dfTot['经度_百度_POI']
    dfPit['lat']=dfTot['纬度_百度_POI']
    dfPit['单位']='百吨'
    for year in timeSeries:
        dfTemp=dfTot[(dfTot['新建时间']<=year)&(dfTot['废弃时间']>year)&(dfTot['是否露天']==1)]#筛选每年露天表
        for city in citys:
            dfT=dfTemp[dfTemp['行政区']==city]
            UIDlist=dfT.index.tolist()
            emi=0
            for UID in UIDlist:
                emiTemp=getPitEmi(UID, year)
                emi=emi+emiTemp
                dfPit.loc[UID,year]=emiTemp/100000#单位  百吨
            dfPitRes.loc[city,year]=emi/1000#unit 吨
        dfPitRes.loc['全国',year]=dfPitRes[year].sum()/1000000#unit Mt或Tg
        dfPitRes.loc['全国','单位']='Tg'
    #计算废弃部分
    
    ###首先随机生成判断是否水淹
    ##原表中有提供默认的是否水淹
    # idxList=dfTot.index.tolist()
    # for i in idxList:
    #     flag=random.random()
    #     if flag>floodRate:
    #         #未被水淹
    #         dfTot.loc[i,'是否水淹']=0#表示未被水淹
    #     else:
    #         dfTot.loc[i,'是否水淹']=1#表示被水淹
    #上面这部分可以在后期改成固定的水淹序列（先要蒙特卡洛）
    dfAbanRes=pd.DataFrame(index=citys,columns=timeSeries+['单位'])
    dfAbanRes['单位']='吨'
    #以下为矿井级别
    dfAban=pd.DataFrame(index=idxList,columns=timeSeries+['lng','lat','单位'])
    dfAban['lng']=dfTot['经度_百度_POI']
    dfAban['lat']=dfTot['纬度_百度_POI']
    dfAban['单位']='百吨'
    for year in timeSeries:
        dfTemp=dfTot[(dfTot['废弃时间']<=year)&(dfTot['是否露天']==0)]#筛选废弃表
        for city in citys:
            dfT=dfTemp[dfTemp['行政区']==city]
            UIDlist=dfT.index.tolist()
            emi=0
            for UID in UIDlist:
                flag=0#代表不是当年废弃
                if dfT.loc[UID,'新建时间']==dfT.loc[UID,'废弃时间']:
                    flag=1
                emiTemp=getAbanEmi(UID, year,dfT.loc[UID,'是否水淹'],flag)
                emi=emi+emiTemp
                dfAban.loc[UID,year]=emiTemp/100000#单位  百吨只考虑该矿井的排放，不考虑11年以前的AMM
            dfAbanRes.loc[city,year]=(emi+iniAbanEmi(year,city))/1000#unit 吨
        dfAbanRes.loc['广东',year]=iniAbanEmi(year,'广东')/1000#unit 吨
        dfAbanRes.loc['全国',year]=dfAbanRes[year].sum()/1000000#unit Mt或Tg
        dfAbanRes.loc['全国','单位']='Tg'
    #计算矿后活动
    dfResRes=pd.DataFrame(index=citys,columns=timeSeries+['单位'])
    dfResRes['单位']='吨'
    #以下为矿井级别
    dfRes=pd.DataFrame(index=idxList,columns=timeSeries+['lng','lat','单位'])
    dfRes['lng']=dfTot['经度_百度_POI']
    dfRes['lat']=dfTot['纬度_百度_POI']
    dfRes['单位']='百吨'
    for year in timeSeries:
        dfTemp=dfTot[(dfTot['新建时间']<=year)&(dfTot['废弃时间']>year)]#筛选每年井工+露天表
        for city in citys:
            dfT=dfTemp[dfTemp['行政区']==city]
            UIDlist=dfT.index.tolist()
            emi=0
            for UID in UIDlist:
                emiTemp=getResEmi(UID, year)
                emi=emi+emiTemp
                dfRes.loc[UID,year]=emiTemp/100000#单位  百吨
            dfResRes.loc[city,year]=emi/1000#unit 吨
        dfResRes.loc['全国',year]=dfResRes[year].sum()/1000000#unit Mt或Tg
        dfResRes.loc['全国','单位']='Tg'
    #计算回收利用
    #此处只计算矿井级，行政区域级别的已经计算得到
    dfRev=pd.DataFrame(index=idxList,columns=timeSeries+['lng','lat','单位'])
    dfRev['lng']=dfTot['经度_百度_POI']
    dfRev['lat']=dfTot['纬度_百度_POI']
    dfRev['单位']='百吨'
    for year in timeSeries:
        dfTemp=dfTot[(dfTot['新建时间']<=year)&(dfTot['废弃时间']>year)&(dfTot['是否露天']==0)]#筛选每年井工表
        for city in citys:
            dfT=dfTemp[dfTemp['行政区']==city]
            UIDlist=dfT.index.tolist()
            emi=0
            for UID in UIDlist:
                emiTemp=getRevEmi(UID, year)
                dfRev.loc[UID,year]=emiTemp/100000#单位  百吨
    #总排放
    dfAll=pd.DataFrame(index=idxList,columns=timeSeries+['lng','lat','单位'])
    dfAll['lng']=dfTot['经度_百度_POI']
    dfAll['lat']=dfTot['纬度_百度_POI']
    dfAll['单位']='百吨'
    for year in timeSeries:
        for UID in idxList:
            arr=np.array([dfMine.loc[UID,year],dfPit.loc[UID,year],dfAban.loc[UID,year],dfRes.loc[UID,year],-dfRev.loc[UID,year]])
            dfAll.loc[UID,year]=np.nansum(arr)
            
    #结果写入
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


##随机验证废弃煤矿80%的假设
# if __name__=='__main__':
#     pathVarify=r'不确定性/随机验证.xlsx'
#     dfSeriesSave=pd.read_excel(pathVarify,sheet_name='序列',index_col='UID')
#     dfOutcomeSave=pd.read_excel(pathVarify,sheet_name='结果',index_col='序号')
#     count=2000#蒙特卡洛次数
#     for i in range(count):
#         idxList=dfTot.index.tolist()
#         #每次蒙特卡洛先生成水淹序列
#         for j in idxList:
#             flag=random.random()
#             if flag>0.8:
#                 #未被水淹
#                 dfSeriesSave.loc[j,i]=0#表示未被水淹
#             else:
#                 dfSeriesSave.loc[j,i]=1#表示被水淹 
#         #逐渐计算AMM排放
#         for year in timeSeries:
#             dfTemp=dfTot[(dfTot['废弃时间']<=year)&(dfTot['是否露天']==0)]#筛选废弃表
#             UIDlist=dfTemp.index.tolist()
#             emi=0
#             for UID in UIDlist:
#                 flag=0#代表不是当年废弃
#                 if dfTot.loc[UID,'新建时间']==dfTot.loc[UID,'废弃时间']:
#                     flag=1
#                 emiTemp=getAbanEmi(UID, year,dfSeriesSave.loc[UID,i],flag)
#                 emi=emi+emiTemp
#             dfOutcomeSave.loc[i,year]=emi/1000000000#unit Tg
#     with pd.ExcelWriter(pathVarify) as writer:
#         dfSeriesSave.to_excel(writer,sheet_name='序列')
#         dfOutcomeSave.to_excel(writer,sheet_name='结果')

###加上2011前AMM排放
# if __name__=='__main__':
#     citys=list(citys)+['广东']
#     for i in range(2011,2020):
#         Emi=0
#         for city in citys:
#             Emi=Emi+iniAbanEmi(i,city)
#         Emi=Emi/1000000000
#         out=str(i)+'='+str(Emi)
#         print(out)


# ###计算锁定排放
# if __name__=='__main__':
#     citys=list(citys)
#     time=list(range(2020,2061))
#     dfLockEmi=pd.DataFrame(index=citys+['广东'],columns=time+['单位'])
#     for year in time:
#         #情景1  现有废弃煤矿
#         dfTemp=dfTot[(dfTot['废弃时间']<=2019)&(dfTot['是否露天']==0)]#筛选废弃表
#         for city in citys:
#             dfT=dfTemp[dfTemp['行政区']==city]
#             UIDidx=dfT.index.tolist()
#             emi=0
#             for UID in UIDidx:
#                 flag=0#代表不是当年废弃
#                 if dfTot.loc[UID,'新建时间']==dfTot.loc[UID,'废弃时间']:
#                     flag=1
#                 emiTemp=getAbanEmi(UID, year,dfT.loc[UID,'是否水淹'],flag)
#                 emi=emi+emiTemp
#             dfLockEmi.loc[city,year]=(emi+iniAbanEmi(year,city))/1000#unit 吨
#             dfLockEmi.loc[city,'单位']='吨'
#         dfLockEmi.loc['广东',year]=iniAbanEmi(year,'广东')/1000#unit 吨
#         dfLockEmi.loc['广东','单位']='吨'
#         dfLockEmi.loc['全国',year]=dfLockEmi[year].sum()/1000000#unit Mt或Tg
#         dfLockEmi.loc['全国','单位']='Tg'
#     dfLockEmi.to_excel('锁定排放.xlsx',sheet_name='Sheet1')

# ### 每年历史排放和新增排放

# if __name__=='__main__':
#       # for year in timeSeries:
#       #   dfTemp=dfTot[(dfTot['废弃时间']<year)&(dfTot['是否露天']==0)]#筛选废弃表_不含当年
#       #   Emi=0
#       #   for UID in dfTemp.index:
#       #       emiTemp=getAbanEmi(UID, year,dfTemp.loc[UID,'是否水淹'],0)
#       #       Emi=Emi+emiTemp
#       #   Emi=Emi/1000000000#Tg
#       #   for city in list(citys)+['广东']:
#       #       Emi=Emi+iniAbanEmi(year,city)/1000000000#Tg
#     #2011前废弃煤矿的排放
#     for year in timeSeries:
#         Emi=0
#         for city in list(citys)+['广东']:
#             Emi=Emi+iniAbanEmi(year,city)/1000000000#Tg
#         print(Emi)
        