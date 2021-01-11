#!/usr/bin/env python
# _*_ coding: utf-8 _*_
# Create by Jack on 2021/01/06
"""
    该模块为工具模块
"""
from excel_logger import logger
import xlwings as xw


def is_empty(obj):
    """对对象进行判空和None处理"""
    return obj == '' or (obj is None)


def write_data2excel(filename, sheet_header, body_data):
    """将数据写入 Excel 文件中"""
    # 新建 Excel 文件
    app = xw.App(add_book=False, visible=False)
    workbook = app.books.add()
    worksheet = workbook.sheets['Sheet1']
    try:
        # 写入数据
        worksheet.range('A1').value = sheet_header  # 在第一行中写入表头
        worksheet.range('A2').value = body_data  # 在第二行中写入主体数据
        worksheet.autofit()  # 自动调整行列的宽高
        workbook.save(filename)  # 保存，filename为file的全路径
    except Exception as e:
        logger.error('将数据写入 Excel 文件失败，报错：%s', e)
        raise Exception
    finally:
        # 无论是否有异常，都需要释放资源，否则会很容易出现内存溢出
        workbook.close()
        app.quit()
