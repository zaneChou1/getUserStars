#!/usr/bin/python
#coding:utf-8
import requests
import xlrd
import xlwt
import random

# read userName data
a = xlrd.open_workbook("githubUsername.xlsx")
jsonUser = a.sheet_by_index(0).col_values(0,1)
print(jsonUser)

# write first row
f = xlwt.Workbook()
sheet1 = f.add_sheet("result",cell_overwrite_ok=True)
row0 = ["用户名", "star数量"]
for i in range(0, len(row0)):
    sheet1.write(0, i, row0[i])

# # 写第一列
# for i in range(0, len(colum0)):
#     sheet1.write(i + 1, 0, colum0[i], set_style('Times New Roman', 220, True))

list = ['apache/servicecomb-service-center', 'apache/servicecomb-java-chassis', 'apache/servicecomb-saga-actuator',
        'apache/servicecomb-fence', 'apache/servicecomb-kie', 'apache/servicecomb-samples', 'apache/servicecomb-mesher',
        'apache/servicecomb-docs', 'apache/servicecomb-website', 'apache/servicecomb-pack',
        'apache/servicecomb-toolkit']

# Data handle
for i in range(len(jsonUser)):
    result = requests.get("https://api.github.com/users/"+jsonUser[i]+"/starred")

    userlist = []
    for j in range(len(result.json())):
        full_name = result.json()[j]["full_name"]
        if full_name in list:
            userlist.append(full_name)

    sheet1.write(i+1,0,jsonUser[i])
    sheet1.write(i+1,1,len(userlist))
    userlist.clear()

# output star result data
num = random.randint(0,100)
print(num)
f.save("result" + str(num) + ".xls")
