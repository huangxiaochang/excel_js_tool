#!usr/bin/env python
# -*- coding: utf-8 -*-

from PyQt5 import QtWidgets
import PyQt5.sip
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QPushButton
from PyQt5.QtWidgets import QGridLayout
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtWidgets import QComboBox
import sys

from excel_to_js import Excel2JS
from js_to_excel import JS2Excel
 
class MyWindow(QtWidgets.QWidget):

    def __init__(self):
        super(MyWindow,self).__init__()

        grid = QGridLayout()
        self.setLayout(grid)

        self.excel2js_type = 'es'
        self.js2excel_type = 'es'

        self.combos = []
        self.combos_info = [
            {
                'choose_list':['common', 'es', 'normal'],
                'func':self.onActivated1,
            },
            {
                'choose_list':['common', 'es', 'normal'],
                'func':self.onActivated2,
            }
        ]

        for i in range(len(self.combos_info)):
            myCombo = QComboBox(self)
            for j in self.combos_info[i]['choose_list']:
                myCombo.addItem(j)
            myCombo.activated[str].connect(self.combos_info[i]['func'])
            myCombo.setFixedSize(80, 30)
            grid.addWidget(myCombo, i, 1)
  
        self.buttons = []
        self.buttons_info = [
            {
                'text':"Excel转Js",
                'func':self.change_excel_to_js,
                'tip':'选择输入的Excel文件，再选择输出js的文件夹',
            },
            {
                'text':"Js转Excel",
                'func':self.change_js_to_excel,
                'tip':'选择输入js的文件夹，再选择输出Excel文件所在的文件夹',
            }
        ]

        for i in range(len(self.buttons_info)):
            myButton = QPushButton('Button', self)
            myButton.setObjectName("myButton")
            myButton.setText(self.buttons_info[i]['text'])
            myButton.clicked.connect(self.buttons_info[i]['func'])
            myButton.setToolTip(self.buttons_info[i]['tip'])
            myButton.setFixedSize(80, 30)
            myButton.move(50, 50)
            self.buttons.append(myButton)
            grid.addWidget(myButton, i, 0)

    def onActivated1(self, text):
        print(text)
        self.excel2js_type = text

    def onActivated2(self, text):
        print(text)
        self.js2excel_type = text

    def suc_info(self, text):
        reply = QMessageBox.information(self,'提示', text, QMessageBox.Ok | QMessageBox.Close, QMessageBox.Close)
        if reply == QMessageBox.Ok:
            print('你选择了Ok！')
        else:
            print('你选择了Close！')

    def change_excel_to_js(self):
        input_file, filetype = self.choose_excel_file('选择输入的Excel文件')
        if input_file == "":
            print("\n取消选择")
            return
        output_dir = self.choose_dir('选择输出文件夹')
        if output_dir == "":
            print("\n取消选择")
            return
        Excel2JS(input_file, output_dir, self.excel2js_type)
        self.suc_info('Excel转换到JS成功')

    def change_js_to_excel(self):
        input_dir = self.choose_dir('选择输入文件夹')
        if input_dir == "":
            print("\n取消选择")
            return
        output_dir = self.choose_dir('选择输出文件所在文件夹')
        if output_dir == "":
            print("\n取消选择")
            return
        output_file = output_dir + '\\results.xlsx'
        JS2Excel(input_dir, output_file, self.js2excel_type)
        self.suc_info('JS转换到Excel成功')


    def choose_excel_file(self, text):
        fileName, filetype = QFileDialog.getOpenFileName(self, text, "./", "All Files (*);;Excel Files (*.xlsx)")  #设置文件扩展名过滤,注意用双分号间隔
        print(fileName,filetype)
        return fileName, filetype

    def choose_file(self):
        fileName1, filetype = QFileDialog.getOpenFileName(self, "选取文件", "./", "All Files (*);;Text Files (*.txt)")  #设置文件扩展名过滤,注意用双分号间隔
        print(fileName1,filetype)

    def choose_dir(self, text):
        directory = QFileDialog.getExistingDirectory(self, text, "./")                 #起始路径
        print(directory)
        return directory

    # def msg(self):
    #     directory1 = QFileDialog.getExistingDirectory(self,
    #                   "选取文件夹",
    #                   "./")                 #起始路径
    #     print(directory1)

    #     fileName1, filetype = QFileDialog.getOpenFileName(self,
    #                   "选取文件",
    #                   "./",
    #                   "All Files (*);;Text Files (*.txt)")  #设置文件扩展名过滤,注意用双分号间隔
    #     print(fileName1,filetype)

    #     files, ok1 = QFileDialog.getOpenFileNames(self,
    #                   "多文件选择",
    #                   "./",
    #                   "All Files (*);;Text Files (*.txt)")
    #     print(files,ok1)
    #     fileName2, ok2 = QFileDialog.getSaveFileName(self,
    #                   "文件保存",
    #                   "./",
    #                   "All Files (*);;Text Files (*.txt)")

if __name__=="__main__":
    app = QtWidgets.QApplication(sys.argv) 
    myshow=MyWindow()
    myshow.resize(640, 480)
    myshow.setWindowTitle('多语言转换工具')
    myshow.show()
    sys.exit(app.exec_()) 
