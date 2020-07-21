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

    # how to choose buildings
    # 选楼函数
    def pick(self):
        path = self.path
        index = self.index
        max_epoch = self.max_epoch
        max_sample = self.max_sample
        building_limit = self.building_limit

        df_building_brand = pd.read_excel(path)
        df_building_brand = df_building_brand.set_index(index)
        building_set = set()
        remain = [building_limit] * df_building_brand.shape[1]
        remain[-1] = 0
        vector = np.array(list(map(self.init_vector, remain)))
        j = 0
        while sum(vector) > 0 and j <= max_epoch:
            ser = df_building_brand.dot(vector).sort_values(0, ascending=False)
            ser = ser[ser > 0]
            if ser.size > 0:
                max_value = ser.max()
                ser = ser[ser == max_value]
                sort = ser.head(max_sample).sample(
                    n=1).index.tolist()[0]
                # print(ser.size)
                add = df_building_brand.ix[sort].values
                add[-1] = 0
                remain = [i - j for i, j in zip(remain, add)]
                vector = np.array(list(map(self.init_vector, remain)))
                df_building_brand.drop(sort, inplace=True)
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
                   'plan_numbers': 10
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
        (result_set, remain) = module.pick()
        list_result = list(result_set)
        list_result.sort()
        plan_set.add("方案{0}:{1}; 损失向量:{2}; 楼宇数量: {3}".format(str(num),
                                                             str(list_result), str(remain), str(result_set.__len__())))

    print("\n")
    for colletion in plan_set:
        print(colletion)
