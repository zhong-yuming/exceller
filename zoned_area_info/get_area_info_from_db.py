#!/usr/bin/env python
# _*_ coding: utf-8 _*_
# Create by Jack on 2021/03/02

"""
    从数据库中读取国家统一区划信息，写入 Excel 表格中，格式如下：
    梅州市         xxx县        xxx镇
    xxx县         xxx镇        xxx居委会
    xxx县         xxx镇        xxx居委会

    这样的格式是为了实现 Excel 中地区的多级联动。
"""
from pymysql import *
import xlwings as xw
from collections import deque

data_cols = []
area_list = deque()
area_list.append(['梅州市', '441400000000'])


def get_col_data(header_title, area_code):
    """主函数"""
    try:
        # 创建数据库连接
        conn = connect(host='localhost', port=3306, user='root', password='root', database='demo', charset='utf8')
        # 通过连接获取 cursor 对象
        cursor = conn.cursor()
        cursor.execute('select * from area_info where title = %s and area_code = %s', (header_title, area_code))
        result = cursor.fetchone()
        temp_list = []
        temp_list.append(result[1])
        cursor.execute('select title, area_code from area_info where parent_area_code = %s', result[0])
        result = cursor.fetchall()
        if len(result) == 0:
            return
        for item in result:
            temp_list.append(item[0])
            area_list.append([item[0], item[1]])
        data_cols.append(temp_list)
    except Exception as e:
        e.with_traceback()
    finally:
        cursor.close()
        conn.close()


if __name__ == '__main__':
    while True:
        if len(area_list) == 0:
            break
        item = area_list.popleft()
        get_col_data(item[0], item[1])

    print(data_cols)
    app = xw.App(visible=False, add_book=False)
    # 设置打开时不显示提示信息，加快打开速度
    app.display_alerts = False
    app.screen_updating = False

    workbook = app.books.open('地区信息表.xlsx')
    worksheet = workbook.sheets['Sheet1']
    # 开始写入数据
    for index, data_col in enumerate(data_cols):
        worksheet.range((1, index + 3)).options(transpose=True).value = data_col

    workbook.save()
    workbook.close()
    app.quit()
