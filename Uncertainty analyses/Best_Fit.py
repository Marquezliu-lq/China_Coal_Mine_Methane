# -*- coding: utf-8 -*-
"""
Created on Mon Dec 11 20:02:51 2023

@author: MarqueLiu
"""
# Use this code to generate the best fit of mine EF at provincial level

import pandas as pd
import numpy as np
from scipy.stats import norm, lognorm, weibull_min, logistic, gamma
import matplotlib.pyplot as plt
from scipy.stats import kstest


def fit_best(data):
    # 选择要尝试的分布类型
    distributions = [norm, lognorm, weibull_min, logistic, gamma]
    
    # 初始化最佳拟合的变量
    best_fit_name = ""
    best_fit_params = None
    best_fit_ks_statistic = np.inf
    
    # 尝试每种分布类型并找到最适合的
    for distribution in distributions:
        # 估计分布的参数
        params = distribution.fit(data)
        
        # 使用Kolmogorov-Smirnov检验进行分布拟合的质量检查
        ks_statistic, p_value = kstest(data, distribution.cdf, args=params)
        
        # 如果p值小于0.05，说明拟合不够好，跳过该分布
        if p_value < 0.05:
            continue
        
        # 计算Kolmogorov-Smirnov统计量，值越小越好
        if ks_statistic < best_fit_ks_statistic:
            best_fit_name = distribution.name
            best_fit_params = params
            best_fit_ks_statistic = ks_statistic
    return best_fit_name, best_fit_params

dfSample=pd.read_excel(r'参数抽样.xlsx',0,index_col=0)
dfTot=pd.read_excel(r'所有煤矿总表.xlsx',0,index_col='UID')#所有信息总表

# 定义行政区列表
regions = ['安徽', '北京', '重庆', '福建', '甘肃', '广西', '贵州', '河北', '黑龙江', '河南', '湖北', '湖南', '内蒙古', '江苏', '江西', '吉林', '辽宁', '宁夏', '陕西', '山东', '山西', '四川', '新疆', '云南']
# 遍历每个行政区
for region in regions:
    index_high_gas = dfTot[(dfTot['新建时间']==2011)&(dfTot['是否露天']==0)&(dfTot['行政区'] == region) & (dfTot['瓦斯等级'].isin(['高瓦斯', '突出']))].index
    index_low_gas = dfTot[(dfTot['新建时间']==2011)&(dfTot['是否露天']==0)&(dfTot['行政区'] == region) & (~dfTot['瓦斯等级'].isin(['高瓦斯', '突出']))].index
    name="";param=None
    # 更新 '高瓦斯' 和 '突出' 的排放因子分布
    if not (region=='北京' or region=='福建'):
        data=dfTot.loc[index_high_gas,'排放因子'].tolist()
        name,param=fit_best(data)
        print(f"{region}: high")
        print(f"Best fit distribution: {name}")
        print(f"Best fit parameters: {param}")
        # 根据拟合的分布和参数进行随机抽样
        sample_size =5000  # 指定抽样大小
        samples=[]
        # 根据拟合的分布进行抽样
        if name == 'norm':
            samples = norm.rvs(loc=param[0], scale=param[1], size=sample_size)
        elif name == 'lognorm':
            samples = lognorm.rvs(s=param[0], loc=param[1], scale=param[2], size=sample_size)
        elif name == 'weibull_min':
            samples = weibull_min.rvs(c=param[0], loc=param[1], scale=param[2], size=sample_size)
        elif name == 'logistic':
            samples = logistic.rvs(loc=param[0], scale=param[1], size=sample_size)
        elif name == 'gamma':
            samples = gamma.rvs(a=param[0], loc=param[1], scale=param[2], size=sample_size)
        dfSample.loc[f'{region}-High','distribution']=name
        if len(samples):
            samples=np.where(samples < 0, 0, samples)
            dfSample.loc[f'{region}-High',4:5005]=samples
    # 更新非 '高瓦斯' 和 '突出' 的排放因子
    name="";param=None;samples=[]
    data=dfTot.loc[index_low_gas,'排放因子'].tolist()
    name,param=fit_best(data)
    print(f"{region}: low")
    print(f"Best fit distribution: {name}")
    print(f"Best fit parameters: {param}")
    # 根据拟合的分布进行抽样
    if name == 'norm':
        samples = norm.rvs(loc=param[0], scale=param[1], size=sample_size)
    elif name == 'lognorm':
        samples = lognorm.rvs(s=param[0], loc=param[1], scale=param[2], size=sample_size)
    elif name == 'weibull_min':
        samples = weibull_min.rvs(c=param[0], loc=param[1], scale=param[2], size=sample_size)
    elif name == 'logistic':
        samples = logistic.rvs(loc=param[0], scale=param[1], size=sample_size)
    elif name == 'gamma':
        samples = gamma.rvs(a=param[0], loc=param[1], scale=param[2], size=sample_size)
    dfSample.loc[f'{region}-Low','distribution']=name
    if len(samples):
        samples=np.where(samples < 0, 0, samples)
        dfSample.loc[f'{region}-Low',4:5005]=samples
dfSample.to_excel(r'参数抽样.xlsx',sheet_name='sheet1')
