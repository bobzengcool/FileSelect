#-*- encodeing = utf-8 -*-
#@time : 2021/7/3 14:58
#@filename : FileSplitUiTool.py
#@product : PyCharm
from PyQt5.QtCore import QDir
from PyQt5.QtGui import QIntValidator
from PyQt5.QtWidgets import QMainWindow, QFileDialog

from Frame.FileSplitUi import FileSplit_MainWindow


class FileSplitQMainWindow(QMainWindow, FileSplit_MainWindow):
    def __init__(self, parent=None):
        super(FileSplitQMainWindow, self).__init__(parent)
        self.setupUi(self)
        pIntvalidator = QIntValidator(self)
        pIntvalidator.setRange(1, 65535)
        self.lineEdit_hangshu.setValidator(pIntvalidator)
        self.pushButton_wenjian.clicked.connect(self.selectFile)
        self.pushButton_lujin.clicked.connect(self.selectPath)
        self.pushButton_caifen.clicked.connect(self.FileSplitTool)

    def selectFile(self):
        self.textout_log.append("开始选择拆分文件")
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.AnyFile)
        dialog.setFilter(QDir.Files)
        if dialog.exec():
            filenames = dialog.selectedFiles()
            self.textout_log.append(filenames[0])
            self.lineEdit_wenjian.setText(filenames[0])

    def selectPath(self):
        self.textout_log.append("开始选择文件路径")
        dialog = QFileDialog()
        OutFilePath=dialog.getExistingDirectory(self,"选取文件夹","C:/")
        self.textout_log.append(OutFilePath)
        self.lineEdit_lujin.setText(OutFilePath)

    def FileSplitTool(self):
        self.textout_log.append("开始对文件进行拆分工作.....")
        limit = self.lineEdit_hangshu.text()
        inputFile = self.lineEdit_wenjian.text()
        outputPath = self.lineEdit_lujin.text()
        self.textout_log.append("每个excel行数为："+limit)
        self.textout_log.append("需要拆分的文件为：" + inputFile)
        self.textout_log.append("拆分后文件存放路径为：" + outputPath)
        if len(limit)<=0 or len(inputFile)<=0 or len(outputPath) :
            self.textout_log.append("参数不正确")


