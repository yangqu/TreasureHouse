# -*- coding: utf-8 -*-
"""
Created on Mon Aug 17 15:46:13 2020

@author: Focusmedia
"""

import jieba
import numpy as np
from collections import Counter
import difflib
import threading
import pandas as pd
from pandas import DataFrame
import os

lock = threading.Lock()

#input = r'C:/Users/Focusmedia/Desktop/word/testsim.xlsx'
#output = r'C:/Users/Focusmedia/Desktop/wordresults'

input = r'D:\source\testsim.xlsx'
output = r'D:\source'


def edit_similar(str1, str2):
    len_str1 = len(str1)
    len_str2 = len(str2)
    taglist = np.zeros((len_str1 + 1, len_str2 + 1))
    for a in range(len_str1):
        taglist[a][0] = a
    for a in range(len_str2):
        taglist[0][a] = a
    for i in range(1, len_str1 + 1):
        for j in range(1, len_str2 + 1):
            if str1[i - 1] == str2[j - 1]:
                temp = 0
            else:
                temp = 1
            taglist[i][j] = min(taglist[i - 1][j - 1] + temp, taglist[i][j - 1] + 1, taglist[i - 1][j] + 1)
    return 1 - taglist[len_str1][len_str2] / max(len_str1, len_str2)


def cos_sim(str1, str2):
    co_str1 = (Counter(str1))
    co_str2 = (Counter(str2))
    p_str1 = []
    p_str2 = []
    for temp in set(str1 + str2):
        p_str1.append(co_str1[temp])
        p_str2.append(co_str2[temp])
    p_str1 = np.array(p_str1)
    p_str2 = np.array(p_str2)
    return p_str1.dot(p_str2) / (np.sqrt(p_str1.dot(p_str1)) * np.sqrt(p_str2.dot(p_str2)))


def user_cos_sim(list1, list2):
    str1 = []
    str2 = []
    for i in list1:
        str1.extend(jieba.lcut(i))
    for i in list2:
        str2.extend(jieba.lcut(i))
    co_str1 = (Counter(str1))
    co_str2 = (Counter(str2))
    p_str1 = []
    p_str2 = []
    for temp in set(str1 + str2):
        p_str1.append(co_str1[temp])
        p_str2.append(co_str2[temp])
    p_str1 = np.array(p_str1)
    p_str2 = np.array(p_str2)
    return p_str1.dot(p_str2) / (np.sqrt(p_str1.dot(p_str1)) * np.sqrt(p_str2.dot(p_str2)))


def compare(str1, str2):
    if str1 == str2:
        return 1.0
    diff_result = difflib.SequenceMatcher(None, str1, str2).ratio()
    str1 = jieba.lcut(str1)
    str2 = jieba.lcut(str2)
    cos_result = cos_sim(str1, str2)
    edit_reslut = edit_similar(str1, str2)
    return cos_result * 0.4 + edit_reslut * 0.3 + 0.3 * diff_result


def parellel_compare(row1, row2, p, result):
    for i in row1['address1'].index.values:

        row_data1 = row1.loc[i]
        max_simi = 0
        address1 = row_data1['address1']
        address2 = ''
        samecity = row2[row2['城市2'] == row_data1['城市']]
        print(p, i ,samecity)
        for j in row2['address2'].index.values:
            if row2['城市2'].loc[j] == row_data1['城市']:
                # row_data2 = samecity['address2'].iloc[j]
                row_data2 = row2['address2'].loc[j]
                if compare(address1, row_data2) >= max_simi:
                    max_simi = compare(address1, row_data2)
                    address2 = row_data2
        if max_simi > 0.7:
            write_to_file(address1, address2, max_simi, result)
            # print({'Thread': p, 'address1': address1, 'address2': address2, 'similarity': max_simi})
    # return result
    # result_path = output + '\\result.xlsx'
    #


#
def write_to_file(address1, address2, max_simi, result):
    lock.acquire()  # thread blocks at this line until it can obtain lock

    # in this section, only one thread can be present at a time.
    #result.append({'address1': address1, 'address2': address2, 'similarity': max_simi}, ignore_index=True)
    result.append({'address1': address1, 'address2': address2, 'similarity': max_simi})
    #print(result)
    lock.release()


def main():
    number = 5
    threads = []
    data = pd.read_excel(input, sheet_name=0)
    # address1 = pd.read_excel(input, sheet_name=0)['address1']
    # address2 = pd.read_excel(input, sheet_name=0)['address2']
    # result = pd.DataFrame()
    result = []
    step = int(data['address1'].index.values.__len__() / number)
    # print(data['address1'].index.values.__len__())
    for i in range(number):
        #print(i * step)
        q = None
        if i == number - 1:
            q = data.iloc[i * step:]
        else:
            q = data.iloc[i * step:(i + 1) * step]
        #print(data)
        # samecity =  data[data['城市2'] == row_data1['城市']]
        t1 = threading.Thread(target=parellel_compare, args=(q, data, i, result))
        threads.append(t1)

        t1.start()
        t1.join()
    result_path = output + '\\test_result.xlsx'
    #print(result)
    DataFrame(result).to_excel(result_path, sheet_name='result', index=False, header=True)


if __name__ == '__main__':
    main()