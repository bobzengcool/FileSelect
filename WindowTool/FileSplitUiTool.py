#-*- encodeing = utf-8 -*-
#@time : 2021/7/3 14:58
#@filename : FileSplitUiTool.py
#@product : PyCharm
import xlrd
import xlwt
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
        limit = self.lineEdit_hangshu.text()
        inputFile = self.lineEdit_wenjian.text()
        outputPath = self.lineEdit_lujin.text()
        self.textout_log.append("每个excel行数为：{limit}".format(limit=limit))
        self.textout_log.append("需要拆分的文件为：" + inputFile)
        self.textout_log.append("拆分后文件存放路径为：" + outputPath)
        self.excelFileSplitTool()

    def excelFileSplitTool(self):
        self.textout_log.append("开始对文件进行拆分工作.....")
        limit = self.lineEdit_hangshu.text()
        inputFile = self.lineEdit_wenjian.text()
        outputPath = self.lineEdit_lujin.text()
        self.textout_log.append("每个excel行数为：{limit}".format(limit=limit))
        self.textout_log.append("需要拆分的文件为：" + inputFile)
        self.textout_log.append("拆分后文件存放路径为：" + outputPath)
        print("测试",len(limit.strip()),len(inputFile.strip()),len(outputPath.strip()))
        if len(limit.strip()) <= 0 or len(inputFile.strip()) <= 0 or len(outputPath.strip())<=0 :
            self.textout_log.append("参数不正确")
            return
        IntLimit=int(limit)
        file = inputFile
        rb = xlrd.open_workbook(filename=file)  # 打开文件
        print(rb.sheet_names())  # 获取所有表格名字
        sheet1 = rb.sheet_by_index(0)  # 通过索引获取表格
        # 读取表中的数据
        nrow = sheet1.nrows
        ncol = sheet1.ncols  # 找到行列总数
        self.textout_log.append("文件总行数为:{nrow}".format(nrow=nrow))
        self.textout_log.append("文件总列数为:{ncol}".format(ncol=ncol))

        FileNumber=nrow // IntLimit +1
        self.textout_log.append("文件将会被拆分为:{FileNumber}个文件".format(FileNumber=FileNumber))

        beginLine=0;

        for i in range(0,FileNumber):
            print("开始生成第:{i}个文件".format(i=i))
            self.textout_log.append("开始生成第:{i}个文件".format(i=i))
            endLine = IntLimit * i + IntLimit
            print("开始行为{beginLine}结束行为{endLine}".format(beginLine=beginLine,endLine=endLine-1))
            rows=[]
            for row in range(beginLine,endLine):
                if( row >= nrow):
                    break
                rows.append(sheet1.row_values(row, 0, ))
            print(rows)
            wb = xlwt.Workbook()  # 创建文件
            ws = wb.add_sheet("sheet{num}".format(num=i))  # 增加sheet

            row_idx = 0
            for new_r in rows:  # 这个循环用来在新的文件中按行、列写入数据
                col_idx = 0
                for v in new_r:
                    ws.write(row_idx, col_idx, v)
                    col_idx = col_idx + 1
                row_idx = row_idx + 1
            wb.save("{outputPath}/sheet{num}.xlsx".format(outputPath=outputPath,num=i))  # 将写入数据的workbook对象保存为文件
            beginLine = endLine
        self.textout_log.append("恭喜你，文件拆分完成!!!!!!!!!!")




