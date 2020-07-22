import pandas as pd
import numpy as np


# init algorithm Greedy class
# 算法类Greedy
class Greedy:
    # init function，five parameters will be initialized
    # 初始化方法，参数被赋值
    def __init__(self, config):
        self.path = config['path']
        self.index = config['index']
        self.max_epoch = config['max_epoch']
        self.max_sample = config['max_sample']
        self.building_limit = config['building_limit']
        self.must_go = config['must_go']

    # how to choose buildings
    # 选楼函数
    def pick(self):
        path = self.path
        index = self.index
        max_epoch = self.max_epoch
        max_sample = self.max_sample
        building_limit = self.building_limit
        must_go = self.must_go

        # 读取building+brand文件为一个dataframe，取名为df_building_brand
        df_building_brand = pd.read_excel(path)
        # 把df_building_brand里的building_id作为index
        df_building_brand = df_building_brand.set_index(index)
        # 新建一个set()集合,building_set用来存放挑中的楼
        building_set = set()
        # 新建一个remain向量，存放11个为5的元素，并把最后一个元素化为0，作为已剩余数量向量的初始化
        remain = [building_limit] * df_building_brand.shape[1]
        # 选取df_building_brand表中选中的楼（即必须取的楼）的行
        rowdata = df_building_brand.loc[must_go]
        # 把这些上述行的值相加
        deletedata = np.sum(rowdata, axis=0)
        # 更新remain
        remain = [i - j for i, j in zip(remain, deletedata)]
        remain[-1] = 0
        print(remain)
        # 新建一个vector向量，作为品牌空缺向量的初始化（把remain向量进行init_vector()函数处理,再化成list，再化成array）
        vector = np.array(list(map(self.init_vector, remain)))
        df_building_brand = df_building_brand.drop(must_go)
        # print(df_building_brand)
        # building_set.add(must_go.toSet)
        for building in must_go:
            building_set.add(building)
        j = 0
        # 循环条件，品牌空缺（即有1）和循环次数不大于max_epoch
        while sum(vector) > 0 and j <= max_epoch:
            # 把df_building_brand表和vector向量点积，并把结果表进行降序排序
            ser = df_building_brand.dot(vector).sort_values(0, ascending=False)
            # print( df_building_brand.dot(vector))
            # print(ser)
            # 把ser里得分为0 的删除
            ser = ser[ser > 0]
            if ser.size > 0:
                # 最大得分，ser截取最大得分的行
                max_value = ser.max()
                ser = ser[ser == max_value]
                # 截取ser最大得分里的前5行，在里面任选一行，再取该行的index
                sort = ser.head(max_sample).sample(n=1).index.tolist()[0]
                # 取df_building_brand表里的上述index对应的行值，并把最后一列即all值化为0
                add = df_building_brand.loc[sort].values
                add[-1] = 0
                # 更新remain已剩余数量向量和vector品牌空缺向量
                remain = [i - j for i, j in zip(remain, add)]
                vector = np.array(list(map(self.init_vector, remain)))
                # 在df_building_brand表里，把选中的index楼删除
                df_building_brand.drop(sort, inplace=True)
                # 把选中的楼加入building_set里
                building_set.add(sort)
                print("\n")
                print("-" * 50 + "start" + "-" * 50)
                print("第{0}轮---已挑选楼宇集合：{1}".format(j, building_set))
                print("第{0}轮---已品牌空缺向量：{1}".format(j, vector))
                print("第{0}轮---已剩余数量向量：{1}".format(j, remain))
                print("-" * 50 + " end " + "-" * 50)
                j = j + 1
            else:
                break

        return building_set, remain

    @staticmethod
    # 定义的init_vector函数，大于0的化为1，其他为0
    def init_vector(x):
        result = 1 if x > 0 else 0
        return result


if __name__ == '__main__':
    # parameters dictionary
    # 参数字典
    dict_config = {'path': 'D:/data/building+brand10.xlsx',
                   'index': 'building_id',
                   'max_epoch': 30,
                   'max_sample': 5,
                   'building_limit': 5,
                   'plan_numbers': 10,
                   'must_go': [3823, 4146, 12911, 17016, 18332, 63348, 72016, 80832, 2000109, 3000663, 3000771, 3009764,
                               3030612, 3030704, 5003538, 5011982, 5018989, 5030449]
                   }

    # transform class to instance
    # 实例化类
    module = Greedy(dict_config)
    # result container
    # 结果容器
    plan_set = set()
    # call Greedy.pick in a loop
    # 循环选楼
    for num in range(1, dict_config['plan_numbers']):
        print("#" * 50 + "开始方案" + str(num) + "#" * 50)
        # 把module.pick()选楼函数l里返回的building_set和remain存入result_set和remain_set里
        (result_set, remain_set) = module.pick()
        list_result = list(result_set)
        list_result.sort()
        plan_set.add("方案{0}:{1}; 损失向量:{2}; 楼宇数量: {3}"
                     .format(str(num), str(list_result), str(remain_set), str(result_set.__len__())))
        print("#" * 50 + "结束方案" + str(num) + "#" * 50)
        print("\n")
    for element in plan_set:
        print(element)
