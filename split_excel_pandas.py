#!/usr/bin/env python
# _*_ coding: utf-8 _*_
# Create by Jack on 2021/01/06

"""
    需求：
        将数据量为 1W 的 Excel 文件，根据某列的值进行分组。值相同的分为一组，并将组各自写入一个新的 Excel 文件
        新 Excel 文件名：列值名+时间
"""
import xlwings as xw
from tkinter import filedialog
from tkinter import messagebox
import os
import time
import pandas

messagebox.showinfo(title="提示信息", message="请选择要打开的文件")
# 弹出文件打开窗口，初始化路径为 C:/Users
# input_file_path = filedialog.askopenfilename(initialdir="C:\\Users")
file_path = filedialog.askopenfilename(initialdir="%USERPROFILE%\\桌面")
print(file_path)

messagebox.showinfo(title="提示信息", message="请选择保存目录")
# 弹出目录选择窗口，用以后续保存 Excel 的根目录
input_folder = filedialog.askdirectory()
print(input_folder)

if file_path == '' or file_path is None:
    print("请选择要打开的文件")
    exit()

if input_folder == '' or input_folder is None:
    print("请选择保存目录")
    exit()

# 创建存放的根目录，目录名为“日期+时间”
root_folder = time.strftime('%Y年%m月%d日%H%M%S', time.localtime(time.time()))
root_folder_path = os.path.join(input_folder, root_folder)  # 根目录的绝对路径
os.mkdir(root_folder_path)

# 表示从第四行开始读，去读 A 列到 E 列的数据，并设置空不为 NaN
data = pandas.read_excel(file_path, header=3, usecols='A:E', keep_default_na=False)
# 分组
groups = data.groupby(['开户网点'])
groups.

print(groups)

# # 设置为不显示地打开 Excel，读取 Excel 文件
# app = xw.App(visible=False, add_book=False)
# workbook = app.books.open(input_file_path)  # 打开 Excel 文件
# worksheet = workbook.sheets[0]  # 选择第一个 sheet 表单
#
# # 获取表单的数据列数、行数
# # num_col = worksheet.api.UsedRange.Columns.count
# # num_row = worksheet.api.UsedRange.Rows.count
#
# i = 5  # 从第五行开始遍历
# groups_dict = dict()  # 空字典，存放分组
# while True:
#     col_cell_list = worksheet.range('A' + str(i) + ':E' + str(i)).value
#     group_name = col_cell_list[0]
#     if group_name == "" or group_name is None:  # 如果没有值为空，则说明遍历可以结束
#         break
#
#     if group_name in groups_dict:
#         groups_dict.get(group_name).append(col_cell_list)
#     else:
#         groups_dict[group_name] = [col_cell_list]
#
#     i = i + 1
#
# workbook.close()  # 释放资源
#
# # 遍历字典 groups_dict
# for group_name, data in groups_dict.items():
#     # 新建 Excel 文件
#     temp_workbook = app.books.add()
#     temp_worksheet = temp_workbook.sheets['Sheet1']
#     # 将数据写入
#     # temp_worksheet.range('A1').options(expand='table').value = data
#     temp_worksheet.range('A1').value = ['序号', '姓名', '银行卡号', '开户日期', '开户网点']
#     temp_worksheet.range('A2').value = data
#     temp_worksheet.autofit()
#
#     save_path = os.path.join(root_folder_path, group_name + '.xlsx')  # 保存路径
#     temp_workbook.save(save_path)  # 保存
#     temp_workbook.close()  # 释放资源
#
# app.quit()
