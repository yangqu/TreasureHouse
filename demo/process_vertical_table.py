#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import numpy as np
from datetime import datetime


def read():
    shutdown = pd.read_excel(u'D:/tmp/xxx.xlsx', dtype={'电子序列号':str})
    print(shutdown.dtypes)
    shutdown['数据周期'] = shutdown['数据周期'].apply(lambda x: str(x).replace('_', '-'))
    shutdown.dropna(subset=['数据周期'])
    pivot_shutdown = shutdown.pivot_table(index='电子序列号', columns='数据周期', values='解析后开关机时间', aggfunc='first')

    duration = pd.read_excel(u'D:/tmp/zzz.xlsx', dtype={'电子序列号':str})
    print(duration.dtypes)
    duration.dropna(subset=['日期'])
    duration['日期'] = duration[pd.notnull(duration["日期"])]
    # duration['日期'] = duration['日期'].apply(lambda x:  x.strftime('%Y-%m-%d'))
    left = duration[['城市', '项目id', '项目名称', '点位id', '电子序列号']].drop_duplicates()
    pivot_palylog_duration = duration.pivot_table(index='电子序列号', columns='日期', values='分钟统计监播时长', aggfunc='first')
    pivot_heartbeat_duration = duration.pivot_table(index='电子序列号', columns='日期', values='心跳时长', aggfunc='first')
    pivot_palylog_duration.fillna('')
    pivot_heartbeat_duration.fillna('')

    first = pd.merge(left, pivot_shutdown, on='电子序列号', how='left')
    second = pd.merge(first, pivot_palylog_duration, on='电子序列号', how='left')
    third = pd.merge(second, pivot_heartbeat_duration, on='电子序列号', how='left')
    third.dropna(subset=['2021-01-20'])
    print(third.dtypes)
    # third = third[pd.notnull(third[['2021-01-20']])]
    third.to_excel(u'D:/tmp/yyy.xlsx')
    print(third)


if __name__ == '__main__':
    read()
