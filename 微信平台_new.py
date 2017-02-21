# -*- coding: utf-8 -*-
"""
@author: Wang Bin
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
from pylab import *
import datetime as dt
from WindPy import *
import xlrd
import xlwt
from xlutils.copy import copy
w.start()

date = u'20170217' #上周五的日期
lh1_swith = 1
zd1_swith = 1
cz1_swith = 1
wj1_swith = 1
start_date = (dt.datetime.strptime(date, '%Y%m%d') - 4 * dt.timedelta(1)).strftime('%Y%m%d')
tdays_list = w.tdays(start_date, date).Times
last_tday = w.tdaysoffset(-1, start_date)

mpl.rcParams['font.sans-serif'] = ['SimHei']
mpl.rcParams['axes.unicode_minus'] = False


file_path = u'E:/研究生资料/路透实验室/量华资本/微信平台/'

'''
量华1号
'''
if lh1_swith == 1:
    file_name = u'lh1.xlsx'

    lh1_data = pd.read_excel(file_path + file_name, header=0)
    lh1_data.index = pd.to_datetime(lh1_data['trade_dt'].astype(str))

    try:
        lh1_return = lh1_data['Close'][date]/lh1_data['Close'][last_tday] - 1
    except:
        lh1_return = lh1_data['Close'][-1]/lh1_data['Close'][-2] - 1
    lh1_netvalue = lh1_data['Close'][-1]

    date_end = str(lh1_data.index[-1])[:10]

    fig = plt.figure(figsize=[18, 10])
    ax = fig.add_subplot(1, 1, 1)
    plt.subplots_adjust(wspace=0, hspace=0)
    ax.yaxis.grid(True)
    ax.xaxis.grid(True)
    plt.plot(lh1_data['Close'].resample('W').last().dropna(), 'b', linewidth=3)
    ax.set_title(u'量华1号单位净值(截至%s)' % date_end, fontsize=45)
    ax.set_xticks(lh1_data['trade_dt'].resample('Q').last().index)
    ax.set_xticklabels(lh1_data['trade_dt'].resample('Q').last().astype(str), rotation=-30, fontsize=15)

    max_y = int((float(str(np.max(lh1_data['Close']))) + 3))
    min_y = int((float(str(np.min(lh1_data['Close']))) - 3))
    yrange = list(np.array(range(min_y, max_y, 3)))

    ax.set_yticks(yrange)
    ax.set_yticklabels(yrange, fontsize=16)

    fig.savefig(u'E:/研究生资料/路透实验室/量华资本/微信平台/' + date + u'/lh1.png')

else:
    lh1_return = 0.0
    lh1_netvalue = 100.0


'''
证道1号
'''
if zd1_swith == 1:
    file_name = u'zd1.xls'
    oldwb = xlrd.open_workbook(file_path + file_name, formatting_info=True)
    oldwb_scan = oldwb.sheets()[0]
    nrows = oldwb_scan.nrows
    newwb = copy(oldwb)
    new_edit = newwb.get_sheet(0)

    new_row = nrows
    for i in range(len(tdays_list)):
        try:
            each = tdays_list[i]
            temp_file_name = u'%s/证券投资基金估值表_外贸信托-证道1号证券投资集合资金信托计划_%s.xls' % (date, str(each)[0:10])
            temp_data = pd.read_excel(file_path + temp_file_name)
            temp_location = list(temp_data.iloc[:, 0]).index(u'今日单位净值:')
            temp_netvalue = temp_data.iloc[temp_location, 1]
            new_edit.write(new_row, 0, str(each)[0:4] + str(each)[5:7] + str(each)[8:10])
            new_edit.write(new_row, 1, str(temp_netvalue))
            new_row = new_row + 1
        except:
            try:
                each = tdays_list[i + 1]
                temp_file_name = u'%s/证券投资基金估值表_外贸信托-证道1号证券投资集合资金信托计划_%s.xls' % (date, str(each)[0:10])
                temp_data = pd.read_excel(file_path + temp_file_name)
                temp_location = list(temp_data.iloc[:, 0]).index(u'今日单位净值:')
                temp_netvalue = temp_data.iloc[temp_location, 1]
                each = tdays_list[i]
                new_edit.write(new_row, 0, str(each)[0:4] + str(each)[5:7] + str(each)[8:10])
                new_edit.write(new_row, 1, str(temp_netvalue))
                new_row = new_row + 1
            except:
                print str(each)[0:10] + ' no data!'

    newwb.save(file_path + file_name)

    zd1_data = pd.read_excel(file_path + file_name, header=0)
    zd1_data.index = pd.to_datetime(zd1_data['trade_dt'].astype(str))

    try:
        zd1_return = zd1_data['Close'][date]/zd1_data['Close'][last_tday] - 1
    except:
        zd1_return = zd1_data['Close'][-1]/zd1_data['Close'][-6] - 1
    zd1_netvalue = zd1_data['Close'][-1] * 100

    date_end = str(zd1_data.index[-1])[:10]

    fig = plt.figure(figsize=[18, 10])
    ax = fig.add_subplot(1, 1, 1)
    plt.subplots_adjust(wspace=0, hspace=0)
    ax.yaxis.grid(True)
    ax.xaxis.grid(True)
    plt.plot((zd1_data['Close'] * 100).resample('W').last().dropna(), 'b', linewidth=3)
    ax.set_title(u'量华证道1号单位净值(截至%s)' % date_end, fontsize=45)
    ax.set_xticks(zd1_data['trade_dt'].resample('W').last().index)
    ax.set_xticklabels(zd1_data['trade_dt'].resample('W').last().astype(str), rotation=-30, fontsize=15)

    max_y = int((float(str(np.max(zd1_data['Close']))) + 0.03) * 100)
    min_y = int((float(str(np.min(zd1_data['Close']))) - 0.03) * 100)
    yrange = list(np.array(range(min_y, max_y, 1)))

    ax.set_yticks(yrange)
    ax.set_yticklabels(yrange, fontsize=16)

    fig.savefig(u'E:/研究生资料/路透实验室/量华资本/微信平台/' + date + u'/zd1.png')

else:
    zd1_return = 0.0
    zd1_netvalue = 100.0

'''
成长1号
'''
if cz1_swith == 1:
    file_name = u'cz1.xls'
    oldwb = xlrd.open_workbook(file_path + file_name, formatting_info=True)
    oldwb_scan = oldwb.sheets()[0]
    nrows = oldwb_scan.nrows
    newwb = copy(oldwb)
    new_edit = newwb.get_sheet(0)

    new_row = nrows
    for i in range(len(tdays_list)):
        try:
            each = tdays_list[i]
            temp_file_name = u'%s/SH8431%s年%s月%s日量华成长1号基金委托资产资产估值表.xls' % (date, str(each)[0:4], str(each)[5:7], str(each)[8:10])
            temp_data = pd.read_excel(file_path + temp_file_name)
            temp_location = list(temp_data.iloc[:, 0]).index(u'今日单位净值：')
            temp_netvalue = temp_data.iloc[temp_location, 1]
            new_edit.write(new_row, 0, str(each)[0:4] + str(each)[5:7] + str(each)[8:10])
            new_edit.write(new_row, 1, str(temp_netvalue))
            new_row = new_row + 1
        except:
            try:
                each = tdays_list[i + 1]
                temp_file_name = u'%s/SH8431%s年%s月%s日量华成长1号基金委托资产资产估值表.xls' % (date, str(each)[0:4], str(each)[5:7], str(each)[8:10])
                temp_data = pd.read_excel(file_path + temp_file_name)
                temp_location = list(temp_data.iloc[:, 0]).index(u'昨日单位净值：')
                temp_netvalue = temp_data.iloc[temp_location, 1]
                each = tdays_list[i]
                new_edit.write(new_row, 0, str(each)[0:4] + str(each)[5:7] + str(each)[8:10])
                new_edit.write(new_row, 1, str(temp_netvalue))
                new_row = new_row + 1
            except:
                print str(each)[0:10] + ' no data!'

    newwb.save(file_path + file_name)
    cz1_data = pd.read_excel(file_path + file_name, header=0)
    cz1_data.index = pd.to_datetime(cz1_data['trade_dt'].astype(str))

    try:
        cz1_return = cz1_data['Close'][date]/cz1_data['Close'][last_tday] - 1
    except:
        cz1_return = cz1_data['Close'][-1]/cz1_data['Close'][-6] - 1
    cz1_netvalue = cz1_data['Close'][-1] * 100

    date_end = str(cz1_data.index[-1])[:10]

    fig = plt.figure(figsize=[18, 10])
    ax = fig.add_subplot(1, 1, 1)
    plt.subplots_adjust(wspace=0, hspace=0)
    ax.yaxis.grid(True)
    ax.xaxis.grid(True)
    plt.plot((cz1_data['Close'] * 100).resample('W').last().dropna(), 'b', linewidth=3)
    ax.set_title(u'量华成长1号单位净值(截至%s)' % date_end, fontsize=45)
    ax.set_xticks(cz1_data['trade_dt'].resample('W').last().index)
    ax.set_xticklabels(cz1_data['trade_dt'].resample('W').last().astype(str), rotation=-30, fontsize=15)

    max_y = int((float(str(np.max(cz1_data['Close']))) + 0.03) * 100)
    min_y = int((float(str(np.min(cz1_data['Close']))) - 0.03) * 100)
    yrange = list(np.array(range(min_y, max_y, 1)))

    ax.set_yticks(yrange)
    ax.set_yticklabels(yrange, fontsize=16)

    fig.savefig(u'E:/研究生资料/路透实验室/量华资本/微信平台/' + date + u'/cz1.png')

else:
    cz1_return = 0.0
    cz1_netvalue = 100.0


'''
稳健1号
'''
if wj1_swith == 1:
    file_name = u'%s/量华稳健基金净值.xls' % date
    wj1_data = pd.read_excel(file_path + file_name, header=0)
    wj1_data.index = pd.to_datetime(wj1_data[u'业务日期'].astype(str))

    try:
        wj1_return = wj1_data[u'单位净值'][date]/wj1_data[u'单位净值'][last_tday] - 1
    except:
        wj1_return = wj1_data[u'单位净值'][-1]/wj1_data[u'单位净值'][-6] - 1
    wj1_netvalue = wj1_data[u'单位净值'][-1] * 100

    date_end = str(wj1_data.index[-1])[:10]

    fig = plt.figure(figsize=[18, 10])
    ax = fig.add_subplot(1, 1, 1)
    plt.subplots_adjust(wspace=0, hspace=0)
    ax.yaxis.grid(True)
    ax.xaxis.grid(True)
    plt.plot((wj1_data[u'单位净值'] * 100).resample('W').last(), 'b', linewidth=3)
    ax.set_title(u'量华稳健1号单位净值(截至%s)' % date_end, fontsize=45)
    ax.set_xticks(wj1_data[u'业务日期'].resample('W').last().index)
    ax.set_xticklabels(wj1_data[u'业务日期'].resample('W').last().astype(str), rotation=-30, fontsize=15)

    max_y = int((float(str(np.max(wj1_data[u'单位净值']))) + 0.03) * 100)
    min_y = int((float(str(np.min(wj1_data[u'单位净值']))) - 0.03) * 100)
    yrange = list(np.array(range(min_y, max_y, 1)))

    ax.set_yticks(yrange)
    ax.set_yticklabels(yrange, fontsize=16)

    fig.savefig(u'E:/研究生资料/路透实验室/量华资本/微信平台/' + date + u'/wj1.png')

else:
    wj1_return = 0.0
    wj1_netvalue = 100.0


print lh1_return, zd1_return, cz1_return, wj1_return

return_data = pd.DataFrame(np.array([[str(round(lh1_return * 100, 2)) + '%', str(round(cz1_return * 100, 2)) + '%', str(round(zd1_return * 100, 2)) + '%',
                            str(round(wj1_return * 100, 2)) + '%'], [lh1_netvalue, cz1_netvalue, zd1_netvalue, wj1_netvalue]]).T, index=['lh1', 'cz1', 'zd1', 'wj1'], columns=['weekly_return', 'netvalue'])
return_data.to_csv(u'E:\\研究生资料\\路透实验室\\量华资本\\微信平台\\' + date + u'\\产品收益.csv', encoding='utf8')
