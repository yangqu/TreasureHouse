import jieba
import numpy as np
from collections import Counter
import difflib
import pandas as pd
from pandas import DataFrame
import os

input = r'D:\source'
output = r'D:\result'


def edit_similar(str1, str2):
    len_str1 = len(str1)
    len_str2 = len(str2)
    taglist = np.zeros((len_str1+1, len_str2+1))
    for a in range(len_str1):
        taglist[a][0] = a
    for a in range(len_str2):
        taglist[0][a] = a
    for i in range(1, len_str1+1):
        for j in range(1, len_str2+1):
            if str1[i - 1] == str2[j - 1]:
                temp = 0
            else:
                temp = 1
            taglist[i][j] = min(taglist[i - 1][j - 1] + temp, taglist[i][j - 1] + 1, taglist[i - 1][j] + 1)
    return 1-taglist[len_str1][len_str2] / max(len_str1, len_str2)


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


def compare(str1, str2):
    if str1 == str2:
        return 1.0
    diff_result = difflib.SequenceMatcher(None, str1, str2).ratio()
    str1 = jieba.lcut(str1)
    str2 = jieba.lcut(str2)
    cos_result = cos_sim(str1, str2)
    edit_reslut = edit_similar(str1, str2)
    return cos_result*0.4+edit_reslut*0.3+0.3*diff_result


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


def main():
    for file in os.listdir(input):
        if file.endswith('xlsx'):
            io = input + '\\' + file
            data = pd.read_excel(io, sheet_name=0)
            data['similarity'] = None
            result = pd.DataFrame()
            for i in data.index.values:
                    row_data1 = data.iloc[i].to_dict()
                    max_simi = 0
                    address1 = row_data1['address1']
                    address2 = ''
                    for j in data.index.values:
                        row_data2 = data.iloc[j].to_dict()
                        if compare(address1, row_data2['address2']) >= max_simi:
                            max_simi = compare(address1, row_data2['address2'])
                            address2 = row_data2['address2']
                    if max_simi > 0.7:
                        result= result.append({'address1': address1, 'address2': address2, 'similarity': max_simi}, ignore_index=True)
                        print({'address1': address1, 'address2': address2, 'similarity': max_simi})

            result_path = output + '\\' + file
            DataFrame(result).to_excel(result_path, sheet_name='result', index=False, header=True)


if __name__ == '__main__':
    main()

