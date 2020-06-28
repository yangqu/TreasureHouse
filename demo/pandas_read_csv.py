# import the function pandas to format the dataset as a table called dataframe
# 导入pandas包，起一个别名pd
import pandas as pd
# import the package os to read some folder to list files
# 导入os包
import os


def main():
    # input dir ,if your files in this folder,the name will be listed,you can change it
    # input为输入路径，r代表字符串里边没有转义符号
    input = r'D:\format'
    # in this case we use the for-loop function for file list
    # 使用OS包来读取输入的路径下的所有文件，file来引用读取到的文件，这是一个循环操作
    for file in os.listdir(input):
        # if the file name end with txt
        # 如果文件以txt为结尾，那么继续运行
        if file.endswith('txt'):
            # get the full path of the file
            # file为该文件的文件名，和input拼成全路径
            io = input + '\\' + file
            # read csv file.It is not only the csv ,but also excel(pd.read_excel(io, sheet_name=0))
            # the parameter io is the stream of the file content,error_bad_lines means skip the parsing error
            # 使用pandas的方法读取csv，并且用\t间隔
            source = pd.read_csv(io, error_bad_lines=False, sep='\t')
            # extract the columns you want by the column name
            # 读取文件里的响应字段
            source_extract = source[['tbb.locationid', 'push.stat', 'playlog.stat', 'heartbeat.stat']]
            # transform Null into 0
            # 把null转化成0
            source_extract.fillna(0, inplace=True)
            # print the dataframe
            # 打印
            print(source_extract)


# main function ,the python program entrance,it is the start of the dream
# 函数入口，代表程序执行
# 所有的程序都是从上往下来执行，先是包的引入，然后遇到函数例如main（），先读进内存，但是不是运行，主程序看到if __name__ == '__main__'，这个算是通用写法，主程序执行的入口，执行
# if中的部分，if中调用了main（）函数，于是，运行main（）函数里边的内容已经从上往下运行，运行到print截至，程序运行完毕
if __name__ == '__main__':
    main()
