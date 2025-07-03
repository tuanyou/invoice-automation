# -*- coding: utf-8 -*-
import json
import re
import requests


def get_access_token():
    '''获取访问凭证'''
    url = 'https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal'
    data = {
        "app_id": "xx",
        "app_secret": "xx"
    }
    ret = requests.post(url=url, data=json.dumps(data, ensure_ascii=False))
    data = ret.json()
    return data.get("tenant_access_token")


def get_sheet_info(spreadsheet_id, access_token):
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    url = f"https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/{spreadsheet_id}/sheets/query"
    params = {
        "valueRenderOption": "ToString",
        "dateTimeRenderOption": "FormattedString"
    }
    res = requests.get(url, headers=headers, params=params)
    data = res.json()
    sheets_info = []
    for sheet in data['data']['sheets']:
        sheet_name = sheet['title']
        sheet_range = sheet['sheet_id']
        sheets_info.append({
            "sheet_name": sheet_name,
            "sheet_range": sheet_range
        })
    return sheets_info


access_token = get_access_token()
# 电子表格运费计算器id（每三个月会更新一次）
spreadsheet_id = 'NDLHsXUy4hC4JmtH2wTcP9mWncV'
sheets_info = get_sheet_info(spreadsheet_id, access_token)

weritu_count = 0
yinghe_count = 0
desu_count = 0
weritu_info = []
yinghe_info = []
desu_info = []
current_sheet_name = sheets_info[1]['sheet_name']
current_date = re.search(r'(\d+\.\d+)', current_sheet_name).group(1)
print('当前处理{}发票'.format(current_date))

for info in sheets_info:
    if current_date in info['sheet_name'] and '为途' in info['sheet_name']:
        weritu_count += 1
        weritu_info.append(info)
    elif current_date in info['sheet_name'] and '盈合' in info['sheet_name']:
        yinghe_count += 1
        yinghe_info.append(info)
    elif current_date in info['sheet_name'] and '德速' in info['sheet_name']:
        desu_count += 1
        desu_info.append(info)

print(f"本周待处理为途发票总数：{weritu_count}，本周待处理盈和发票总数：{yinghe_count}，本周待处理德速发票总数：{desu_count}")
for info in weritu_info:
    print(info)
for info in yinghe_info:
    print(info)
for info in desu_info:
    print(info)

