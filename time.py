import pandas as pd
# path = '‪D:\pycharm\Mathematical_Modeling_20190919\data111.xlsx'
data = pd.DataFrame(pd.read_excel('data111.xlsx'))
print(data.index)#获取行的索引名称
print(data.columns)#获取列的索引名称
print(data['时间'])