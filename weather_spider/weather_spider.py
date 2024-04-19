import random
import time
import requests
from bs4 import BeautifulSoup
import xlwings as xw
import json
from config import info
from city_num import city
import os


def write(a, b, c, value, wb):
    wb.sheets[a - 1][b - 1, c - 1].value = value


def generate_targets():
    targets = []
    if info["from_year_month"][0] == info["to_year_month"][0]:
        for month in range(info["from_year_month"][1], info["to_year_month"][1] + 1):
            targets.append((info["from_year_month"][0], month))
        return targets
    else:
        for year in range(info["from_year_month"][0], info["to_year_month"][0] + 1):
            if year == info["from_year_month"][0]:
                for month in range(info["from_year_month"][1], 13):
                    targets.append((year, month))
            if year == info["to_year_month"][0]:
                for month in range(1, info["to_year_month"][1] + 1):
                    targets.append((year, month))
            else:
                for month in range(1, 13):
                    targets.append((year, month))
        return targets


def get_data(info, wb, row):
    url = "https://tianqi.2345.com/Pc/GetHistory?"
    headers = {
        "authority": "tianqi.2345.com",
        "accept": "application/json, text/javascript, */*; q=0.01",
        "accept-language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
        "referer": "https://tianqi.2345.com/wea_history/45011.htm",
        "sec-ch-ua": "\"Chromium\";v=\"122\", \"Not(A:Brand\";v=\"24\", \"Microsoft Edge\";v=\"122\"",
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": "\"Windows\"",
        "sec-fetch-dest": "empty",
        "sec-fetch-mode": "cors",
        "sec-fetch-site": "same-origin",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36 Edg/122.0.0.0",
        "x-requested-with": "XMLHttpRequest"
    }
    cookies = {
        "Hm_lvt_a3f2879f6b3620a363bec646b7a8bcdd": "1712069986,1712118821,1712293376",
        "Hm_lpvt_a3f2879f6b3620a363bec646b7a8bcdd": "1712293513"
    }
    for target in generate_targets():
        params = {
            "areaInfo[areaId]": city[info['city_name']],
            "areaInfo[areaType]": "2",
            "date[year]": target[0],
            "date[month]": target[1]
        }
        # print(params)
        time.sleep(random.randint(5, 10) / 10)
        response = requests.get(url, headers=headers, cookies=cookies, params=params)
        html = json.loads(response.text)['data']
        soup = BeautifulSoup(html, "html.parser")
        trs = soup.find_all("tr")
        for tr in trs[1::]:
            tds = tr.find_all("td")
            index_col = 1
            # print(tds)
            for td in tds:
                # print(td)
                print(td.get_text(), end=' ')
                write(1, row, index_col, td.get_text(), wb)
                index_col += 1
            print()
            row += 1


def check(city):
    city = list(city.keys())
    if info['city_name'] not in city:
        return False
    if not os.path.exists(f"{info['city_name']}.xlsx"):
        return False
    return True


def start():
    if check(city):
        app = xw.App(visible=True, add_book=False)
        wb = app.books.open(f"{info['city_name']}.xlsx")
        index = 1
        row = 2
        for i in ["日期", "最高温", "最低温", "天气", "风力风向", "空气质量指数"]:
            write(1, 1, index, i, wb)
            index += 1
        get_data(info, wb, row)
        wb.save()
        wb.close()
        app.quit()
    else:
        print("city or file error")


start()
