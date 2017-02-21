# -*- coding: utf-8 -*-
__author__ = 'Wang Bin'

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
from pylab import *
import datetime as dt

mpl.rcParams['font.sans-serif'] = ['SimHei']
mpl.rcParams['axes.unicode_minus'] = False

date = u'20160712'

def ret_cal(dt,range1,range2):
    zs1 = float(test.ix[dt[0], 0][range1:range2])
    zs2 = float(test.ix[dt[1], 0][range1:range2])
    zs3 = float(test.ix[dt[2], 0][range1:range2])
    zs4 = float(test.ix[dt[3], 0][range1:range2])
    zs5 = float(test.ix[dt[4], 0][range1:range2])

    return ((1 + zs1/100) * (1 + zs2/100) * (1 + zs3/100) * (1 + zs4/100) * (1 + zs5/100) - 1) * 100

return_list = list()

def return_rotation(length_temp,text, para_a, para_b):
    for i in range(length_temp):
        if text in test.ix[i][0]:
            range_temp = range(i+6, i+11)  #当文件为周末计算出来时，设为8和13，当文件是周二计算出来时，设为6和11
            return_list.append(ret_cal(range_temp, para_a, para_b))
        else:
            continue

test = pd.read_csv(u'E:\\研究生资料\\路透实验室\\量华资本\\微信平台\\' + date + u'\\Total_Rotation_HY_ALL_20150601.txt')
#test = pd.read_csv(u'E:\\研究生资料\\路透实验室\\量华资本\\微信平台\\' + date + u'\\d__weixh_PyGR_data-monthly_Total_Rotation_HY_ALL_20150601.txt')
#test = pd.read_csv(u'E:\\研究生资料\\路透实验室\\量华资本\\微信平台\\' + date + u'\\c__weixh_PyGR_data_Total_Rotation_HY_ALL_20150201.txt')
#test = pd.read_csv(u'E:\\研究生资料\\路透实验室\\量华资本\\微信平台\\' + date + u'\\c__weixh_PyGR_data_Total_Rotation_HY_ALL_20151231.txt')
length = len(test)

return_rotation(length, 'GRY_0_1', 17, 26)
return_rotation(length, 'GRY_1_1', 17, 26)
return_rotation(length, 'GRY_2_1', 17, 26)
return_rotation(length, 'GRM_10_4', 17, 26)
return_list.append(' ')

return_rotation(length, 'GRY_0_1', 39, 48)
return_rotation(length, 'GRY_1_1', 39, 48)
return_rotation(length, 'GRY_2_1', 39, 48)
return_rotation(length, 'GRM_10_4', 39, 48)
return_list.append(' ')

'''
range1 = np.array(range(len(test)-5, len(test)))
range2 = range1 - 13
range3 = range2 - 13
range4 = range3 - 13

return_list.append(ret_cal(range1, 17, 26))
return_list.append(ret_cal(range2, 17, 26))
return_list.append(ret_cal(range3, 17, 26))
return_list.append(ret_cal(range4, 17, 26))
return_list.append(' ')

return_list.append(ret_cal(range1, 39, 48))
return_list.append(ret_cal(range2, 39, 48))
return_list.append(ret_cal(range3, 39, 48))
return_list.append(ret_cal(range4, 39, 48))
return_list.append(' ')
'''

test = pd.read_csv(u'E:\\研究生资料\\路透实验室\\量华资本\\微信平台\\' + date + u'\\Total_Rotation_Hedged_GR_ALL_20150601.txt')
#test = pd.read_csv(u'E:\\研究生资料\\路透实验室\\量华资本\\微信平台\\' + date + u'\\d__weixh_PyGR_data-monthly_Total_Rotation_Hedged_GR_ALL_20150601.txt')
#test = pd.read_csv(u'E:\\研究生资料\\路透实验室\\量华资本\\微信平台\\' + date + u'\\c__weixh_PyGR_data_Total_Rotation_Hedged_GR_ALL_20150201.txt')
#test = pd.read_csv(u'E:\\研究生资料\\路透实验室\\量华资本\\微信平台\\' + date + u'\\c__weixh_PyGR_data_Total_Rotation_Hedged_HY_ALL_20151231.txt')
length = len(test)

return_rotation(length, 'GRY_0_1', 17, 26)
#return_rotation(length, 'GRY_1_1', 17, 26)
return_rotation(length, 'GRY_2_1', 17, 26)
return_rotation(length, 'GRM_10_4', 17, 26)
return_list.append(' ')

return_rotation(length, 'GRY_0_1', 39, 48)
#return_rotation(length, 'GRY_1_1', 39, 48)
return_rotation(length, 'GRY_2_1', 39, 48)
return_rotation(length, 'GRM_10_4', 39, 48)

'''
range1 = np.array(range(len(test)-5, len(test)))
range2 = range1 - 13
range3 = range2 - 13
range4 = range3 - 13

return_list.append(ret_cal(range1, 17, 26))
return_list.append(ret_cal(range2, 17, 26))
return_list.append(ret_cal(range3, 17, 26))
return_list.append(ret_cal(range4, 17, 26))
return_list.append(' ')

return_list.append(ret_cal(range1, 39, 48))
return_list.append(ret_cal(range2, 39, 48))
return_list.append(ret_cal(range3, 39, 48))
return_list.append(ret_cal(range4, 39, 48))
'''

#print return_list

name_list = [u'GR320成长指数', u'GR320价值指数', u'GR320价值成长指数', u'GR320动量指数', u' ', u'GR320择时成长指数', u'GR320择时价值指数',
             u'GR320择时价值成长指数', u'GR320择时动量指数', u' ', u'GR320对冲成长指数', u'GR320对冲价值成长指数',
             u'GR320对冲动量指数', u' ', u'GR320对冲择时成长指数', u'GR320对冲择时价值成长指数', u'GR320对冲择时动量指数',]

for i in range(len(return_list)):
    if return_list[i] != ' ':
        temp = str(round(return_list[i], 2)) + '%'
        return_list[i] = temp
    else:
        continue

data = pd.DataFrame(return_list, index=name_list)
data.to_csv(u'E:\\研究生资料\\路透实验室\\量华资本\\微信平台\\' + date + u'\\模拟投资产品.csv', encoding='utf8')

data = pd.read_excel(u'E:\\研究生资料\\路透实验室\\量华资本\\微信平台\\' + date + u'\\Total_Rotation_HY_ALL_20161231.xls')
#data = pd.read_excel(u'E:\\研究生资料\\路透实验室\\量华资本\\微信平台\\' + date + u'\\d__weixh_PyGR_data-monthly_Total_Rotation_HY_ALL_20161231.xls')
#data = pd.read_excel(u'E:\\研究生资料\\路透实验室\\量华资本\\微信平台\\' + date + u'\\c__weixh_PyGR_data_Total_Rotation_HY_ALL_20161231.xls')
date_ini = str(data['trade_dt'][0])
date_end = str(data['trade_dt'][len(data)-1])
#data = pd.read_excel(u'E:\\研究生资料\\路透实验室\\量华资本\\微信平台\\' + date + u'\\c__weixh_PyGR_data_Total_Rotation_HY_ALL_20150101.xls')
#data = pd.DataFrame(np.array(data.ix[1:,(-4):]), index=pd.to_datetime(data.index[1:]), columns=[u'GR320择时动量', u'GR320择时价值成长', u'GR320择时价值', u'GR320择时成长'])
data = pd.DataFrame(np.array(data[['DowDailyYield_GRM_10_4', 'DowDailyYield_GRY_2_1', 'DowDailyYield_GRY_1_1', 'DowDailyYield_GRY_0_1']]),
                    index=pd.to_datetime(data['trade_dt'].apply(lambda x: str(x))), columns=[u'GR320择时动量', u'GR320择时价值成长', u'GR320择时价值', u'GR320择时成长'])
data = data / 100 + 1
data = np.cumproduct(data)

fig = plt.figure(figsize=[18,10])
ax = fig.add_subplot(1, 1, 1)
plt.subplots_adjust(wspace=0, hspace=0)
plt.plot(data[[3]][:-2], 'r', label=u'GR320择时成长', linewidth=3)
plt.plot(data[[2]][:-2], 'y', label=u'GR320择时价值', linewidth=3)
plt.plot(data[[1]][:-2], 'b', label=u'GR320择时价值成长', linewidth=3)
plt.plot(data[[0]][:-2], 'g', label=u'GR320择时动量', linewidth=3)
plt.legend(loc=0, numpoints=1)
leg = plt.gca().get_legend()
ltext = leg.get_texts()
plt.setp(ltext, fontsize=35)
ax.set_xticks(pd.date_range(date_ini, date, freq='M'))
ax.set_xticklabels(pd.date_range(date_ini, date, freq='M').strftime('%Y-%m'), rotation=-30, fontsize=30)

max_y = int((float(str(np.max(np.max(data)))[:3]) + 0.2) * 10)
min_y = int((float(str(np.min(np.min(data)))[:3]) - 0.1) * 10)
yrange = list(np.array(range(min_y, max_y, 1)) / 10.0)

ax.set_yticks(yrange)
ax.set_yticklabels(yrange, fontsize=30)
ax.set_title(u'择时型产品近一年（'+date_ini+u'-'+date_end+u'）单位净值', fontsize=45)

fig.savefig(u'E:\\研究生资料\\路透实验室\\量华资本\\微信平台\\' + date + u'\\zs-' + date + u'.png')

data = pd.read_excel(u'E:\\研究生资料\\路透实验室\\量华资本\\微信平台\\' + date + u'\\Total_Rotation_Hedged_GR_ALL_20161231.xls')
#data = pd.read_excel(u'E:\\研究生资料\\路透实验室\\量华资本\\微信平台\\' + date + u'\\d__weixh_PyGR_data-monthly_Total_Rotation_Hedged_GR_ALL_20161231.xls')
#data = pd.read_excel(u'E:\\研究生资料\\路透实验室\\量华资本\\微信平台\\' + date + u'\\c__weixh_PyGR_data_Total_Rotation_Hedged_GR_ALL_20161231.xls')
date_ini = str(data['trade_dt'][0])
date_end = str(data['trade_dt'][len(data)-1])
#data = pd.read_excel(u'E:\\研究生资料\\路透实验室\\量华资本\\微信平台\\' + date + u'\\c__weixh_PyGR_data_Total_Rotation_Hedged_HY_ALL_20150101.xls')
#data = pd.DataFrame(np.array(data.ix[1:,(-4):]), index=pd.to_datetime(data.index[1:]), columns=[u'GR320对冲择时动量', u'GR320对冲择时价值成长', u'GR320对冲择时价值', u'GR320对冲择时成长'])
data = pd.DataFrame(np.array(data[['DowDailyYield_GRM_10_4', 'DowDailyYield_GRY_2_1', 'DowDailyYield_GRY_0_1']]),
                    index=pd.to_datetime(data['trade_dt'].apply(lambda x: str(x))), columns=[u'GR320择时动量', u'GR320择时价值成长', u'GR320择时成长'])
data = data / 100 + 1
data = np.cumproduct(data)

fig = plt.figure(figsize=[18,10])
ax = fig.add_subplot(1, 1, 1)
plt.subplots_adjust(wspace=0, hspace=0)
plt.plot(data[[2]][:-2], 'r', label=u'GR320对冲择时成长', linewidth=3)
#plt.plot(data[[2]], 'y', label=u'GR320对冲择时价值', linewidth=3)
plt.plot(data[[1]][:-2], 'b', label=u'GR320对冲择时价值成长', linewidth=3)
plt.plot(data[[0]][:-2], 'g', label=u'GR320对冲择时动量', linewidth=3)
plt.legend(loc=0, numpoints=1)
leg = plt.gca().get_legend()
ltext = leg.get_texts()
plt.setp(ltext, fontsize=35)
ax.set_xticks(pd.date_range(date_ini, date, freq='M'))
ax.set_xticklabels(pd.date_range(date_ini, date, freq='M').strftime('%Y-%m'), rotation=-30, fontsize=30)

max_y = int((float(str(np.max(np.max(data)))[:3]) + 0.2) * 10)
min_y = int((float(str(np.min(np.min(data)))[:3]) - 0.1) * 10)
yrange = list(np.array(range(min_y, max_y, 1)) / 10.0)

ax.set_yticks(yrange)
ax.set_yticklabels(yrange, fontsize=30)
ax.set_title(u'对冲择时型产品近一年（'+date_ini+u'-'+date_end+u'）单位净值', fontsize=45)

fig.savefig(u'E:\\研究生资料\\路透实验室\\量华资本\\微信平台\\' + date + u'\\dczs-' + date + u'.png')

