from time import sleep
import requests
import pandas as pd
 
"""
:author: Yancey
:date: 2024-02-27
:apiNote: 查询四六级成绩
"""

#包含身份证号和姓名的学生信息表 data.xlsx，与 .py文件放在同一目录下
data = pd.read_excel('data.xlsx', usecols=['姓名', '身份证号'])

# 添加几个字段，用于存储爬取到的信息，
# 索引从 0 开始，因为表中已经有2个字段(姓名+身份证号)，所以从 3 开始，这里需要根据自己的信息表调整

data.insert(2, '准考证号', '')
data.insert(3, 'CET 等级','')
data.insert(4, '成绩', '')
data.insert(5, '听力部分', '')
data.insert(6, '阅读部分', '')
data.insert(7, '写作部分', '')

print(data.columns)
 
# 添加请求头 
headerss = {  
    'Accept': '*/*',  
    'Accept-Encoding': 'gzip, deflate',  
    'Accept-Language': 'zh-CN,zh;q=0.9',  
    'Connection': 'keep-alive',  
    'Host': 'cachecloud.neea.cn',
    'Cookie': '_abfpc=10989d0e3b06a9fc0d0ef2ecac7992d7cd8c1f28_2.0; cna=b76c2d88a93a4f674ce016da72b4da0a; HMF_CI=f7c03c09abe92467573d5dd9c5e2c4c4809ed449269d25378c629ee9b1eaa665b3da0befef3b9f77088fa7d4b9bd23c1e5da279e786b62bbc4eb0b06a5f3d6eca2; community=Home; language=1',  
    #cookie 很重要，没有 cookie，网站大概率报 403
    'If-Modified-Since': 'Fri, 23 Feb 2024 10:35:10 GMT',  
    'If-None-Match': 'W/"65d874de-5d7b"',  
    'Referer': 'https://cjcx.neea.edu.cn/html1/folder/21083/9970-1.htm',  
    'Sec-Ch-Ua': '"Chromium";v="122", "Not(A:Brand";v="24", "Google Chrome";v="122"',  
    'Sec-Ch-Ua-Mobile': '?1',  
    'Sec-Ch-Ua-Platform': '"Android"',  
    'Sec-Fetch-Dest': 'document',  
    'Sec-Fetch-Mode': 'navigate',  
    'Sec-Fetch-Site': 'same-origin',  
    'Sec-Fetch-User': '?1',  
    'Upgrade-Insecure-Requests': '1',  
    'User-Agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Mobile Safari/537.36'  
}
# 用列表存储所有考生的参数
params_list = []
for i in range(len(data)):
    params_list.append({
        'km': 2, #默认查询 六级 成绩， 不用修改，没有六级成绩会查询 四级成绩
        'xm': data.iloc[i, 0],
        'no': data.iloc[i, 1]
    })


# 四六级成绩查询 api
url = "http://cachecloud.neea.cn/api/latest/results/cet"

# 遍历每一个考生的数据
for i in range(len(data)):
    flag = False
    response = requests.get(url, params_list[i], headers=headerss)
    print("*******************************************************************")
    print("开始爬取%s的六级成绩......" % params_list[i]['xm'])
    json_data = response.json()

    if json_data['code'] != 0: #如果没有查到六级成绩，尝试查询四级成绩
        flag = True
        params_list[i]['km'] = 1  # 修改 km 的值为 1，查询四级成绩
        print('再次查询：是否参加四级....')

        # 重新发起请求
        response = requests.get(url, params_list[i], headers=headerss)
        json_data = response.json()

    if json_data['code'] == 0:
        data.iloc[i, 2] = json_data['zkzh']  # 准考证号
        data.iloc[i, 3] =  'CET6' if flag == False else 'CET4'
        data.iloc[i, 4] = json_data['score']  # 分数
        data.iloc[i, 5] = json_data['sco_lc']  # 听力部分
        data.iloc[i, 6] = json_data['sco_rd']  # 阅读部分
        data.iloc[i, 7] = json_data['sco_wt']  # 写作部分

        info = '%d：%s查询成功，成绩为：%s' % (
            (i + 1), params_list[i]['xm'], json_data['score'])
        print(info)

    #四级六级均为参加
    else: 
        print('%d：%s未参加CET 考试...' % ((i + 1), params_list[i]['xm']))
        data.iloc[i, 3] = '未参加'
    
    print("*******************************************************************")
    sleep(0.3)
 

# 导出至csv文件中，会覆写本地文件
print("已全部查询完毕，正在导出到csv......")
data.to_csv('grade.csv', mode='w', encoding='gbk')
print("导出完毕.")
 