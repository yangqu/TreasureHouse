import pandas as pd
import numpy as np


# init data
def createData(brand_list, building_total):
    # dataframe [0,2] as type init,brand_list length equals total number of columns
    df = pd.DataFrame(np.random.randint(0, 2, (building_total, brand_list.__len__())), columns=brand_list)
    # column array
    cl = df.columns.values
    # initialize a container called dictionary {brand : building array}
    brand_building_pool = {}
    # every column
    for brand in cl:
        # filter brand column by value is 1 ,store into tmp container
        tmp = df[brand].loc[df[brand] == 1]
        # insert into brand_building_pool,key is brand  name ,value is building array extract by value==1
        brand_building_pool[brand] = tmp.index.values
    print(brand_building_pool)
    # return result
    return df, brand_building_pool


def sortData(df):
    # new column all is the accumulation of brand score every building
    df['all'] = df.apply(lambda x: x.sum(), axis=1)
    # sort the building by column all
    df.sort_values(by=['all'], ascending=False, inplace=False)
    print(df)
    return df


def pickBuilding(df, building_limit, brand_list):
    # initialize the score array filled zero
    score = [0] * brand_list.__len__()
    # initialize the result container set
    building_list_picked = set()
    for indexs in df.index:
        # get every row
        rowData = df.loc[indexs].values[0: brand_list.__len__()]
        # transform the row type to list type
        rowData = rowData.tolist()
        # accumulate the new list to the score
        score = [i + j for i, j in zip(score, rowData)]
        for x in zip(score, rowData):
            print(x)
        # add the picked building
        building_list_picked.add(indexs)
        print(score)
        # judge whether the standard building_limit is met
        if (np.array(score) >= building_limit).all():
            return building_list_picked
    return building_list_picked


if __name__ == '__main__':
    # Kinds of Brand,it is the list('A','B'...) for short ,you can also code as list('brand1','brand2')
    brand_list = list('ABCDEFGH')
    # Create data including how many buildings
    building_total = 100
    # Total of buildings you have to visit by every brand
    building_limit = 5

    # Function mimic data
    data = createData(brand_list, building_total)[0]
    # Function sort by total brands of building
    sort = sortData(data)
    # Result set buildings
    result = pickBuilding(sort, building_limit, brand_list)
    print("Pick Result : " + str(result))
