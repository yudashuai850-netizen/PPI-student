import requests
import json

url = "https://data.stats.gov.cn/easyquery.htm"

params = {
    "m": "QueryData",
    "dbcode": "hgyd",
    "rowcode": "zb",
    "colcode": "sj",
    "wds": "[]",
    "dfwds": '[{"wdcode":"sj","valuecode":"202608"}]'
}

response = requests.get(url, params=params, timeout=20)

print("状态码：", response.status_code)
print("网址：", response.url)
print("前500个字符：")
print(response.text[:500])