from time import sleep
import requests
import pandas as pd

"""
:author: Yancey
:date: 2024-2-28

-------------
:version: 2.0
:date: 2024-8-28
-- 优化代码逻辑，提高运行速度
-- 修改保存为 CSV 时可能出现的中文乱码问题
-- 增加全局配置，导出自定义
-- 导出结果默认删除身份证与准考证号
"""

# 全局配置
# CSV 文件是否保留未参加的学生（无需保存请改为 False）
retainAbsentStudents = True

# CSV 文件是否保留未通过考试的学生（无需保存请改为 False）
retainFailedStudents = True

# 四六级成绩查询 api, 如果程序无法使用，请自行检查 api 是否有变化
url = "http://cachecloud.neea.cn/api/latest/results/cet"

# 包含身份证号和姓名的学生信息表 data.xlsx，与 .py文件放在同一目录下
data = pd.read_excel('data.xlsx', usecols=['姓名', '身份证号'])
data['身份证号'] = data['身份证号'].astype(str)
# 添加几个字段，用于存储爬取到的信息，
# 索引从 0 开始，因为表中已经有2个字段(姓名+身份证号)，所以从 3 开始，这里需要根据自己的信息表调整

data.insert(2, '准考证号', '')
data.insert(3, 'CET 等级', '')
data.insert(4, '成绩', '')
data.insert(5, '听力部分', '')
data.insert(6, '阅读部分', '')
data.insert(7, '写作部分', '')

print(data.columns)

# 添加请求头（请添加自己的信息）
headerss = {
  }
# 用列表存储所有考生的参数
params_list = []
for i in range(len(data)):
    params_list.append({
        'km': 2,  # 默认查询 六级 成绩， 不用修改，没有六级成绩会查询 四级成绩
        'xm': data.iloc[i, 0],
        'no': data.iloc[i, 1]
    })

# 遍历每一个考生的数据
for i in range(len(data)):
    flag = False
    response = requests.get(url, params_list[i], headers=headerss)
    print("*******************************************************************")
    print("开始爬取%s的六级成绩......" % params_list[i]['xm'])
    json_data = response.json()

    if json_data['code'] != 0:  # 如果没有查到六级成绩，尝试查询四级成绩
        flag = True
        params_list[i]['km'] = 1  # 修改 km 的值为 1，查询四级成绩
        print('再次查询：是否参加四级....')

        # 重新发起请求
        response = requests.get(url, params_list[i], headers=headerss)
        json_data = response.json()

    if json_data['code'] == 0:
        data.iloc[i, 2] = json_data['zkzh']  # 准考证号
        data.iloc[i, 3] = 'CET6' if flag == False else 'CET4'
        data.iloc[i, 4] = int(json_data['score'])  # 分数
        data.iloc[i, 5] = json_data['sco_lc']  # 听力部分
        data.iloc[i, 6] = json_data['sco_rd']  # 阅读部分
        data.iloc[i, 7] = json_data['sco_wt']  # 写作部分

        info = '%d：%s查询成功，成绩为：%s' % (
            (i + 1), params_list[i]['xm'], json_data['score'])
        print(info)

    # 四级六级均未参加
    else:
        # print('%d：%s未参加CET 考试...' % ((i + 1), params_list[i]['xm']))
        data.iloc[i, 3] = '未参加'
        data.iloc[i, 4] = 0

    print("*******************************************************************")
    sleep(0.3)

# 导出至csv文件中，会覆写本地文件
print("已全部查询完毕，正在导出到csv......")
data.drop(columns=['身份证号', '准考证号'], inplace=True)  # 保护隐私删除身份证号和准考证号

if (retainAbsentStudents == False):
    data = data[data['CET 等级'] != '未参加']

if (retainFailedStudents == False):
    data = data[data['成绩'] >= 425]

data = data.sort_values(by='成绩', ascending=False) # 按成绩降序排列
data.to_csv('grade.csv', mode='w',encoding='utf_8_sig',index=False) #使用 utf_8_sig 避免可能出现的中文乱码
print("导出完毕.")
