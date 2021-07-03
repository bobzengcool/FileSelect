#-*- encodeing = utf-8 -*-
#@time : 2021/7/3 14:34
#@filename : FileSplitMain.py
#@product : PyCharm
import sys
from PyQt5.uic import *
from PyQt5.QtWidgets import *
from WindowTool.FileSplitUiTool import FileSplitQMainWindow

if __name__ == '__main__':
    app = QApplication(sys.argv)
    MainWindow = FileSplitQMainWindow()
    MainWindow.show()
    sys.exit(app.exec_())
