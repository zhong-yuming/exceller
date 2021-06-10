#!/usr/bin/env python
# _*_ coding: utf-8 _*_
# Create by Jack on 2021/01/15
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QToolTip
from PyQt5.QtGui import QIcon, QFont
import sys


class Example(QWidget):
    def __init__(self):
        super(Example, self).__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowIcon(QIcon("../jia.ico"))
        # 暂时可以理解为，一旦创建了 QWidget 窗口，窗口就会有 QToolTip 属性
        # 设置 QToolTip 字体(全局设置)
        QToolTip.setFont(QFont('Jetbrains Mono', 12))

        # 设置 QWidget 窗口的全局文本内容
        self.setToolTip('This is a <b>QWidget</b> widget')

        btn = QPushButton('Button', self)
        # 暂时可以理解为，一旦创建了 QPushButton 后，里面就包含了 QToolTip 属性
        # Button 设置的 QToolTip 的值后会覆盖全局的 QToolTip
        btn.setToolTip('This is a <b>QPushButton</b> widget')
        btn.resize(btn.sizeHint())  # sizeHint() 是一个默认大小的按钮
        btn.move(50, 50)
        self.setGeometry(300, 300, 500, 500)
        self.setWindowTitle('Tooltips')
        self.show()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())
