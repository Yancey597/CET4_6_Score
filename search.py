from time import sleep
import requests
import pandas as pd

"""
:author: Yancey
:date: 2024-2-28


-------------
:version: 3.0
:date: 2024-8-28
-- 新语法大幅优化代码逻辑，提高运行速度

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
from time import sleep
import requests
import pandas as pd


def fetch_cet_score(name, id_number, km, headers):
    params = {'km': km, 'xm': name, 'no': id_number}
    response = requests.get(url, params=params, headers=headers)
    if response.status_code == 200:
        json_data = response.json()
        if json_data['code'] == 0:
            return json_data
    return None


def process_data(data, headers, retain_absent=True, retain_failed=True):
    columns_to_add = ['写作部分', '阅读部分', '听力部分', '成绩','CET 等级']

    for col in columns_to_add:
        data.insert(2, col, '')

    for i in range(len(data)):
        name = data.iloc[i, 0]
        id_number = data.iloc[i, 1]

        # Try CET-6 first
        result_6 = fetch_cet_score(name, id_number, 2, headers)
        if result_6:
            data.iloc[i, 2:] = ['CET6', int(result_6['score']), result_6['sco_lc'],
                                result_6['sco_rd'], result_6['sco_wt']]
        else:
            # Try CET-4
            result_4 = fetch_cet_score(name, id_number, 1, headers)
            if result_4:
                data.iloc[i, 2:] = ['CET4', int(result_4['score']), result_4['sco_lc'],
                                    result_4['sco_rd'], result_4['sco_wt']]
            else:
                data.iloc[i, 2] = '未参加'
                data.iloc[i, 3] = 0

        print(f"正在查询 {name} 的成绩......")
        sleep(0.3)

        # Post-processing
    data.drop(columns=['身份证号'], inplace=True)

    if not retain_absent:
        data = data[data['CET 等级'] != '未参加']
    if not retain_failed:
        data = data[data['成绩'] >= 425]

    data.sort_values(by='成绩', ascending=False, inplace=True)


# Load data
data = pd.read_excel('data.xlsx', usecols=['姓名', '身份证号'])

# Headers
headers = {
    'Accept': '*/*',
    'Accept-Encoding': 'gzip, deflate',
    'Accept-Language': 'zh-CN,zh;q=0.9',
    'Connection': 'keep-alive',
    'Host': 'cachecloud.neea.cn',
    'Cookie': '改成自己的',
    # cookie 很重要，没有 cookie，网站大概率报 403
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

# Process data
process_data(data, headers, retain_absent=retainAbsentStudents, retain_failed=retainFailedStudents)

# Save to CSV
data.to_csv('grade.csv', encoding='utf_8_sig', index=False)

print("导出完毕.")