#!/usr/bin/env python
# _*_ coding: utf-8 _*_
# Create by Jack on 2021/01/06
# 内置模块
from tkinter import filedialog, messagebox, simpledialog
import os
import time
import re
# 第三方模块
import xlwings as xw
# 自建模块
from excel_logger import logger
from utils import is_empty


class SplitExcel(object):
    def __init__(self):
        self.input_file_path = None
        self.input_folder_path = None
        self.root_folder_path = None
        self.workbook = None
        self.worksheet = None
        self.groups_dict = dict()
        self.app = xw.App(visible=False, add_book=False)
        self.sheet_header = []

    def __confirm_file_and_folder(self):
        """让用户选择文件和保存目录"""
        logger.info('开始让用户选择文件和保存目录')
        messagebox.showinfo(title="提示信息", message="请选择要打开的文件")
        # 弹出文件打开窗口，初始化路径为 用户桌面
        self.input_file_path = filedialog.askopenfilename(initialdir="%USERPROFILE%/桌面")
        logger.info('用户选择的文件的绝对路径为：%s', self.input_file_path)

        messagebox.showinfo(title="提示信息", message="请选择保存目录")
        # 弹出目录选择窗口，用以后续保存 Excel 的根目录
        self.input_folder_path = filedialog.askdirectory()
        logger.info('用户选择的保存目录的绝对路径为：%s', self.input_folder_path)

        if is_empty(self.input_file_path):
            logger.warn("没有选择要处理的 Excel 文件")
            raise Exception('没有选择要处理的 Excel 文件')

        if is_empty(self.input_folder_path):
            logger.warn("没有选择保存目录")
            raise Exception('没有选择保存目录')

    @staticmethod
    def __select_area():
        """与用户交互，由用户选择要读取的 Excel 区域"""
        selected_area = str(simpledialog.askstring(title='选择区域', prompt='请选择Excel读取区域，如 A5:E6000 或 a5:e6000')).strip()
        logger.info('用户输入的 Excel 区域为：%s', selected_area)
        # 选择的区域不能为空，且要合法
        if not is_empty(selected_area):
            if re.match(r'^[a-zA-Z]\d+:[a-zA-Z]\d+$', selected_area):
                return selected_area.upper()  # 将 a5:e6000 转换为 A5:E6000
        logger.error('没有选择区域或输入的区域格式不正确')
        messagebox.showwarning(title='提示信息', message='没有选择区域或输入的区域格式不正确')
        raise Exception

    @staticmethod
    def __select_group_by():
        """选择根据哪一列来进行分组"""
        group_by = str(simpledialog.askstring(title='选择分组列', prompt='请选择根据哪列进行分组，如 A 或 a')).strip()
        logger.info('用户选择了 %s 列来进行分组', group_by)
        # 不能为空，且要合法
        if not is_empty(group_by):
            if re.match(r'^[a-zA-Z]$', group_by):
                return ord(str.lower(group_by)) - 97  # 将 a-z 或 A-Z 装换为数字 0-25
        logger.error('没有输入列或输入列的格式不正确')
        messagebox.showwarning(title='提示信息', message='没有输入列或输入列的格式不正确')
        raise Exception

    def __make_root_folder(self):
        """创建存放处理后得到的 Excel 文件的根目录"""
        logger.info('开始创建存放导出的 Excel 的根目录')
        # 创建存放的根目录，目录名为“日期+时间”
        root_folder = time.strftime('%Y年%m月%d日%H%M%S', time.localtime(time.time()))
        self.root_folder_path = os.path.join(self.input_folder_path, root_folder)  # 根目录的绝对路径
        try:
            os.mkdir(self.root_folder_path)
        except Exception as e:
            logger.error('创建根目录失败，报错：%s', e)
            raise Exception

        return self.root_folder_path

    def __open_excel(self):
        """打开 Excel 文件，获取要操作的 sheet 表单"""
        logger.info('打开 Excel 源文件')
        try:
            # 设置为不显示地打开 Excel，读取 Excel 文件
            self.workbook = self.app.books.open(self.input_file_path)
            self.worksheet = self.workbook.sheets[0]  # 选择第一个 sheet 表单
        except Exception as e:
            logger.error('打开源 Excel 文件失败，报错：%s', e)
            raise Exception

    def __read_excel_data(self, selected_area, group_by):
        """根据传入的 Excel 区域，从 Excel 文件中读取数据，并分组存放到一个字典中"""
        logger.info('开始从源 Excel 文件中读取数据')
        try:
            data = self.worksheet.range(selected_area).value
            self.sheet_header = data[0]
            for col_cell_list in data[1:]:
                group_name = str(col_cell_list[group_by]).strip()  # 剔除字符串两边空格
                if is_empty(group_name):
                    continue
                # 如果不为空，对其进行切片处理，取前 5 个字符来作为组名
                # group_name = group_name[:5]
                if group_name in self.groups_dict:
                    self.groups_dict.get(group_name).append(col_cell_list)
                else:
                    self.groups_dict[group_name] = [col_cell_list]
        except Exception as e:
            logger.error('读取 Excel 数据失败，原因：%s', e)
            raise Exception

    def __split2write_data(self):
        """将从 Excel 读取出来的数据分写到不同的 Excel 文件中"""
        logger.info('开始将数据分写到 Excel 文件中')
        # 遍历字典 groups_dict
        # todo 可以考虑使用多进程来进行读写，加快速度
        for group_name, data in self.groups_dict.items():
            # 新建 Excel 文件
            self.workbook = self.app.books.add()
            self.worksheet = self.workbook.sheets['Sheet1']
            # 将数据写入
            try:
                self.worksheet.range('A1').value = self.sheet_header
                self.worksheet.range('A2').value = data
                self.worksheet.autofit()

                save_path = os.path.join(self.root_folder_path, group_name + '.xlsx')  # 保存路径
                self.workbook.save(save_path)  # 保存
            except Exception as e:
                logger.error('将数据写入 Excel 文件失败，报错：%s', e)
                raise Exception

    def __release_src(self):
        logger.info('程序即将结束，开始释放资源')
        try:
            self.workbook.close()
            self.app.quit()
        except Exception as e:
            logger.error('空指针异常，报错：%s', e)
        finally:
            messagebox.showinfo(title='提示信息', message='Excel处理已结束')

    def run(self):
        """主程序"""
        try:
            self.__confirm_file_and_folder()
            selected_area = self.__select_area()
            group_by = self.__select_group_by()
            self.__make_root_folder()
            self.__open_excel()
            self.__read_excel_data(selected_area, group_by)
            self.__split2write_data()
        except Exception as e:
            e.with_traceback()
        finally:
            self.__release_src()


if __name__ == '__main__':
    se = SplitExcel()
    se.run()
