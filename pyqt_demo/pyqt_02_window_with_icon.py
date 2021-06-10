#!/usr/bin/env python
# _*_ coding: utf-8 _*_
# Create by Jack on 2021/01/15
from PyQt5.QtWidgets import QApplication, QWidget
import sys
from PyQt5.QtGui import QIcon


class Example(QWidget):
    def __init__(self):
        super(Example, self).__init__()  # 指执行 Example 父级的 __init__ 方法

        self._initUI()

    def _initUI(self):
        """Initial UI."""
        self.setWindowTitle("The Window with icon")
        self.setGeometry(250, 150, 500, 500)
        self.setWindowIcon(QIcon('../jia.ico'))
        self.show()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())