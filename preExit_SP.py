# -*- coding: utf-8 -*-
"""
Created on Sat Dec 25 21:14:25 2021

@author: MarqueLiu
"""
'''get the retirement sequence of the points on the frontier curve'''

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
    ## input the sequence of retirement under the emission factor driven strategy, and the constraint on emissions
    N=len(UIDList)
    for i in range(N):
 
        # Last i elements are already in place
        for j in range(0, N-i-1):
 
            if judgeFun(UIDList[j],UIDList[j+1],barrier) :
                UIDList[j], UIDList[j+1] = UIDList[j+1], UIDList[j]
    return UIDList
