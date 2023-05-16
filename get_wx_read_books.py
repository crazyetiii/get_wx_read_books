from openpyxl import Workbook
import requests
import json
import time
import random

# 生成随机的休眠时间（100毫秒到500毫秒之间）
sleep_time = random.randint(500, 5000) / 1000.0

wx_list = [
    "精品小说",  # 100000
    "历史",
    "文学",
    "艺术",
    "人物传记",
    "哲学宗教",  # 600000
    "计算机",
    "心里",
    "社会文化",
    "个人成长",  # 1000000
    "经济理财",
    "政治军事",
    "童书",
    "教育学习",
    "科学技术",
    "生活百科",  # 1600000
    "期刊杂志",
    "原版书",
    "医学健康",
]

category_code_map = {
    "文学": 300000,
    "心里": 800000,
    "计算机": 700000,
    "社会文化": 900000,
    "个人成长": 1000000,
    "科学技术": 1500000

}

base_url = "https://weread.qq.com/web/bookListInCategory/"


def get(name, category_code):
    compute_url = str.format("{}{}?maxIndex=", base_url, category_code)

    result = []
    cur_url = ""

    result_name = name
    file_name = result_name+".json"
    excel_name = result_name+".xlsx"

    with open(file_name, 'w', encoding="utf8") as file:

        for i in range(10000):
            time.sleep(sleep_time)
            cur_url = compute_url+str(i*20)
            print(cur_url)
            # 发送GET请求，并设置请求头
            response = requests.get(cur_url, verify=False)
            data = response.json()
            if '502' in response.text:
                continue
            book_list = data.get('books')
            if len(book_list) == 0:
                break

            for book in book_list:

                book_item = book.get('bookInfo')
                result.append(book_item)
        json.dump(result, file, ensure_ascii=False)

    with open(file_name, 'r', encoding='utf-8') as file:
        loaded_list = json.load(file)
        # 创建一个新的工作簿
        workbook = Workbook()

        # 获取默认的工作表（Sheet）
        sheet = workbook.active

        # 将数据写入工作表
        for item in loaded_list:
            books = []
            # 筛选需要的信息
            books.append(item.get("title"))
            books.append(item.get("author"))
            books.append(item.get("category"))
            books.append(item.get("newRating")/10)
            books.append(item.get("newRatingDetail").get('good'))
            books.append(item.get("newRatingDetail").get('fair'))
            books.append(item.get("newRatingDetail").get('poor'))
            books.append(item.get("newRatingDetail").get('title'))
            sheet.append(books)
        # 保存工作簿到文件
        workbook.save(excel_name)


for k, v in category_code_map.items():
    print(f'正在请求{k}类别')
    get(k, v)
    print(f'请求{k}类别结束')

print(f'finish!!!')
