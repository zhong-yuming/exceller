#!/usr/bin/env python
# _*_ coding: utf-8 _*_
# Create by Jack on 2021/01/07
from PyQt5.QtWidgets import QWidget, QApplication
import sys

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = QWidget()
    window.resize(300, 250)
    window.setWindowTitle("简单窗口")
    window.move(300, 300)
    window.show()

    sys.exit(app.exec_())