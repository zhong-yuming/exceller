#!/usr/bin/env python
# _*_ coding: utf-8 _*_
# Create by Jack on 2021/01/06
"""
    该模块为工具模块
"""
from excel_logger import logger
import xlwings as xw
import xlwt



def is_empty(obj):
    """对对象进行判空和None处理"""
    return obj == '' or (obj is None)


def write_data2excel(filename, sheet_header, body_data):
    """将数据写入 Excel 文件中"""
    # 新建 Excel 文件
    workbook = xlwt.Workbook(encoding='UTF-8')
    worksheet = workbook.add_sheet('Sheet1')
    try:
        # 写入数据
        # 在第一行中写入表头
        for col_num, cell_data in enumerate(sheet_header):
            worksheet.write(0, col_num, cell_data)
        # 在第二行中写入主体数据
        for row_num, row_data in enumerate(body_data):
            for col_num, cell_data in enumerate(row_data):
                worksheet.write(row_num + 1, col_num, cell_data)

        # 调整列宽
        worksheet.col(0).width = 256 * 10
        worksheet.col(1).width = 256 * 20
        worksheet.col(2).width = 256 * 40
        worksheet.col(3).width = 256 * 20
        worksheet.col(4).width = 256 * 40
        # 设置表头行高
        worksheet.row(0).set_style(xlwt.easyxf('font:height 360;'))
        workbook.save(filename)  # 保存，filename为file的全路径
    except Exception as e:
        logger.error('将数据写入 Excel 文件失败，报错：%s', e)
        raise Exception
