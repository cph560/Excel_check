import pandas as pd
import numpy as np
import xlrd 

df1 = pd.read_excel('data1.xlsx')
df2 = pd.read_excel('data2.xlsx')

print(df1)
print(df2)
diff = pd.DataFrame.items(df1)
comparsion = df1.values==df2.values
df1 = df1.fillna(value='NONE')
df2 = df2.fillna(value='NONE')
print(comparsion)
# for i in range(len)
# diff = df1[df1!=df2]

# dict = {}
max1 = df1['wdawd'].tolist()
max2 = df2['dsadsadsa'].tolist()
print(len(max1))
print(max1,max2)
print(max1==max2)

dic = {}

for i in range(df1.shape[0]):
    data1 = df1.loc[i]
    dic[data1[0]] = []
    print(data1)

print(dic)
# print(len(max1['wdawd']))
# print(len(max2['dsadsadsa']))
# list1 = []
# list2 = []
# for x, y in max2['dsadsadsa'].items():
#     list2.append(y)

# for x,y in max1['wdawd'].items():
#     list1.append(y)
# print(list1,list2)
# for i in range(len(list1)):
#     if list1[i]!=list2[i]:
#         dict[i] =  