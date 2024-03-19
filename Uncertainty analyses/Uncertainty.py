# -*- coding: utf-8 -*-
"""
Created on Mon Dec  6 18:16:00 2021
#this code is used for uncertainty analysis of the coal methane emissions

@author: MarqueLiu
"""
import pandas as pd
import random
import math
import numpy as np
import copy
import Sample_Fun

path=r'推断结果/'
pathTot=r'所有煤矿总表.xlsx'
pathPro=r'所有煤矿产量.xlsx'
pathOth=r'其他信息.xlsx'
pathRes=r'敏感性分析V231213.xlsx'
pathPar=r'参数抽样.xlsx'

dfTot_O=pd.read_excel(pathTot,0,index_col='UID')#所有信息总表
dfPro_O=pd.read_excel(pathPro,0,index_col='UID')#产量表
dfPro_O.fillna(0,inplace=True)
dfAbanEf=pd.read_excel(pathOth,sheet_name='2011各行政区排放因子',index_col='行政区')
dfAbanNum=pd.read_excel(pathOth,sheet_name='2011前废弃',index_col='行政区')
dfRevOri_O=pd.read_excel(pathOth,sheet_name='回收利用',index_col='行政区')
dfPar=pd.read_excel(pathPar,0,index_col=0)
dfSample=pd.read_excel(r'参数抽样.xlsx',0,index_col=0)

###常量
sampleNum=2000
factor=0.67
citys=dfTot_O['行政区'].unique()
timeSeries=list(range(2011,2020))
b_O=2.017
D1_O=0.672#水淹
D2_O=0.302#干井
floodRate_O=0.8
aveProd_O=37919#unit t  2011前废弃矿井平均产能(共退出5亿吨产能)
resEfH_O=3#Unit m³/t  高瓦斯和突出矿井瓦斯矿后排放因子
resEfL_O=0.94#Unit m³/t  瓦斯矿井瓦斯矿后排放因子
resPit_O=0.5#Unit m³/t  露天煤矿矿后排放因子
idxList=dfTot_O.index.tolist()#矿井级排放数据使用index


b=b_O
D1=D1_O
D2=D2_O
aveProd=aveProd_O
resEfH=resEfH_O
resEfL=resEfL_O
resPit=resPit_O
floodRate=floodRate_O
####定义获取产量
def getProdAbi(UIDList,year):
    prod=0
    for UID in UIDList:
        temp=dfPro.loc[UID,year]
        if not pd.isna(temp):
            prod=prod+dfPro.loc[UID,year]
    #Unit:t/month
    return prod

### 井工部分
def getMineEmi(UID,year):
    '''传入具体某个UID，在此基础上计算排放和'''
    '''返回单位：kg'''
    ef=dfTot.loc[UID,'排放因子']#unit kg/t
    prod=dfPro.loc[UID,year]*12#unit t
    Emi=ef*prod#kg
    return Emi

### 露天部分
def getPitEmi(UID,year):
    '''传入具体某个露天煤矿的UID，在此基础上计算排放和'''
    ef=dfTot.loc[UID,'排放因子']#unit kg/t
    prod=dfPro.loc[UID,year]*12#unit t   
    Emi=ef*prod
    return Emi
# AMM
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
    abanYear=dfTot.loc[UID,'新废弃时间']
    if abanYear>2011 and not flag:#并非当年废弃，可以-1
        iniEmi=getMineEmi(UID,abanYear-1)
    else:
        iniEmi=getMineEmi(UID,abanYear)
    #dry
    if not isFlood:
        emi=iniEmi*math.pow(1+b*D2*(year-abanYear),(-1/b))
    else:
        emi=iniEmi*math.exp(-(year-abanYear)*D1)
    return emi#kg

#2011前废弃煤矿排放
def iniAbanEmi(year):
    '''计算全国，2011前废弃煤矿的排放'''
    '''针对整个行政区的调用函数'''
    '''返回单位kg'''
    totEmi=0
    for city in list(citys)+['广东']:
        aveFac=dfAbanEf.loc[city,'排放因子']
        iniEmi=aveProd*aveFac#单个的初始排放
        for i in range(2005,2011):
            num=dfAbanNum.loc[city,i]
            #水淹
            emiDry=num*iniEmi*math.pow(1+b*D2*(year-i),(-1/b))*(1-floodRate)#
            #干井
            emiFlood=num*iniEmi*math.exp(-(year-i)*D1)*(floodRate)#
            totEmi=totEmi+emiFlood+emiDry#kg
    return totEmi#kg

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

###函数调用部分，输入抽样的参数
def get_Uncertain():
    #井工
    for year in timeSeries:
        dfTemp=dfTot[(dfTot['新建时间']<=year)&(dfTot['废弃时间']>year)&(dfTot['是否露天']==0)]#筛选每年井工表
        UIDlist=dfTemp.index.tolist()
        emi=0
        for UID in UIDlist:
            emiTemp=getMineEmi(UID, year)#返回单位为 kg
            emi=emi+emiTemp#kg
        dfMineRes.loc[count,year]=emi/1000000000#unit Tg
    #再计算露天部分
    for year in timeSeries:
        dfTemp=dfTot[(dfTot['新建时间']<=year)&(dfTot['废弃时间']>year)&(dfTot['是否露天']==1)]#筛选每年露天表
        UIDlist=dfTemp.index.tolist()
        emi=0
        for UID in UIDlist:
            emiTemp=getPitEmi(UID, year)
            emi=emi+emiTemp
        dfPitRes.loc[count,year]=emi/1000000000#unit Tg
    #计算废弃部分
    for year in timeSeries:
        dfTemp=dfTot[(dfTot['新废弃时间']<=year)&(dfTot['是否露天']==0)]#筛选废弃表
        UIDlist=dfTemp.index.tolist()
        emi=0
        for UID in UIDlist:
            flag=0#代表不是当年废弃
            if dfTemp.loc[UID,'新建时间']==dfTemp.loc[UID,'新废弃时间']:
                flag=1
            emiTemp=getAbanEmi(UID, year,dfTemp.loc[UID,'是否水淹'],flag)
            emi=emi+emiTemp
        dfAbanRes.loc[count,year]=(emi+iniAbanEmi(year))/1000000000#Tg
    #计算矿后活动
    #以下为矿井级别
    for year in timeSeries:
        dfTemp=dfTot[(dfTot['新建时间']<=year)&(dfTot['废弃时间']>year)]#筛选每年井工+露天表
        UIDlist=dfTemp.index.tolist()
        emi=0
        for UID in UIDlist:
            emiTemp=getResEmi(UID, year)
            emi=emi+emiTemp
        dfResRes.loc[count,year]=emi/1000000000#unit Tg
    #回收利用
    for year in timeSeries:
        dfRev.loc[count,year]=dfRevOri.loc['全国利用',year]#Tg
    #总排放
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
    
#抽样主体程序
if __name__=='__main__':
    idx=range(0,sampleNum)
    initial()
    global count
    for count in idx:##每一次循环
        ###抽样后建立的新表和参数
        lobal dfTot,dfPro,dfRevOri
        dfTot=copy.deepcopy(dfTot_O)
        dfPro=copy.deepcopy(dfPro_O)
        dfRevOri=copy.deepcopy(dfRevOri_O)
        ##抽样参数
        b=dfPar.loc['b',count]*b_O
        D1=dfPar.loc['Di',count]*D1_O
        D2=dfPar.loc['Di',count]*D2_O
        aveProd=dfPar.loc['before 2011',count]*aveProd_O
        resEfH=dfPar.loc['Post-EF-High',count]*resEfH_O
        resEfL=dfPar.loc['Post-EF-Low',count]*resEfL_O
        resPit=resPit_O
        floodRate=dfPar.loc['Flooding Rate',count]
        abanRate=dfPar.loc['between13-14',count]*0.5

        ##抽样表-更新水淹状态和废弃状态
        UIDlist=dfTot_O.index.tolist()
        for UID in UIDlist:
            if random.random() < floodRate:
                dfTot.loc[UID,'是否水淹']=1
            else:
                dfTot.loc[UID,'是否水淹']=0
            if dfTot_O.loc[UID,'废弃时间'] in [2013,2014]:
                if random.random() < abanRate:
                    dfTot.loc[UID,'新废弃时间']=2013#该因素只考虑对AMM的影响
                else:
                    dfTot.loc[UID,'新废弃时间']=2014
        ##更新排放因子
        index=dfTot[dfTot['是否露天']==1].index;dfTot.loc[index,'排放因子']*=Sample_Fun.all_sample(dfSample, 'Sur-EF', 1)
        # 定义行政区列表
        regions = ['安徽', '北京', '重庆', '福建', '甘肃', '广西', '贵州', '河北', '黑龙江', '河南', '湖北', '湖南', '内蒙古', '江苏', '江西', '吉林', '辽宁', '宁夏', '陕西', '山东', '山西', '四川', '新疆', '云南']
        # 遍历每个行政区
        for region in regions:
            index_high_gas = dfTot[(dfTot['是否露天']==0)&(dfTot['行政区'] == region) & (dfTot['瓦斯等级'].isin(['高瓦斯', '突出']))].index
            index_low_gas = dfTot[(dfTot['是否露天']==0)&(dfTot['行政区'] == region) & (~dfTot['瓦斯等级'].isin(['高瓦斯', '突出']))].index
            
            # 更新 '高瓦斯' 和 '突出' 的排放因子
            if not (region=='北京' or region=='福建'):
                dfTot.loc[index_high_gas, '排放因子'] = Sample_Fun.all_sample(dfSample, f'{region}-High', len(index_high_gas))

            # 更新非 '高瓦斯' 和 '突出' 的排放因子
            dfTot.loc[index_low_gas, '排放因子'] = Sample_Fun.all_sample(dfSample, f'{region}-Low', len(index_low_gas))
      ## 更新产量表
        for year in timeSeries:
            dfPro.loc[:,year]*=Sample_Fun.triangle_sample(1,0.9,1.1,1)

        #回收利用
        for year in timeSeries:
            dfRevOri.loc[:,year]*=Sample_Fun.triangle_sample(1,0.9,1.1,1)
          #计算
        get_Uncertain()
         
    with pd.ExcelWriter(pathRes) as writer:
        dfMineRes.to_excel(writer,sheet_name='井工')
        dfPitRes.to_excel(writer,sheet_name='露天')
        dfAbanRes.to_excel(writer,sheet_name='废弃')
        dfResRes.to_excel(writer,sheet_name='矿后')
        dfRev.to_excel(writer,sheet_name='回收利用')
        dfAll.to_excel(writer,sheet_name='总排放')










        
        
