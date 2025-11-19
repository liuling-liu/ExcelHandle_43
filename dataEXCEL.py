# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'dataEXCEL.ui'
#
# Created by: PyQt5 UI code generator 5.11.2
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import qApp,QFileDialog

import sys
import pandas as pd
import os
import glob
import numpy as np
import matplotlib.pyplot as plt
root=""
fileNum = 0
myrow=0

#自定义函数SaveExcel用于保存数据到Excel
def SaveExcel(df,isChecked):
    # 将提取后的数据保存到Excel
    if (isChecked):
        writer = pd.ExcelWriter('mycell.xlsx', engine='openpyxl')
    else:
        global temproot
        writer = pd.ExcelWriter(temproot + '/mycell.xlsx')
    df.to_excel(writer, 'sheet1')
    writer.close()
class Ui_MainWindow(QtWidgets.QWidget):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(838, 596)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.list1 = QtWidgets.QListView(self.centralwidget)
        self.list1.setGeometry(QtCore.QRect(1, 1, 171, 401))
        self.list1.setObjectName("list1")
        self.text1 = QtWidgets.QTextEdit(self.centralwidget)
        self.text1.setGeometry(QtCore.QRect(110, 450, 631, 21))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.text1.sizePolicy().hasHeightForWidth())
        self.text1.setSizePolicy(sizePolicy)
        self.text1.setObjectName("text1")
        self.viewButton = QtWidgets.QPushButton(self.centralwidget)
        self.viewButton.setGeometry(QtCore.QRect(746, 450, 75, 23))
        self.viewButton.setObjectName("viewButton")

        self.rButton1 = QtWidgets.QRadioButton(self.centralwidget)
        self.rButton1.setGeometry(QtCore.QRect(10, 450, 111, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.rButton1.setFont(font)
        self.rButton1.setObjectName("rButton1")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setEnabled(False)
        self.label.setGeometry(QtCore.QRect(7, 409, 61, 16))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.rButton2 = QtWidgets.QRadioButton(self.centralwidget)
        self.rButton2.setGeometry(QtCore.QRect(10, 430, 131, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.rButton2.setFont(font)
        self.rButton2.setCheckable(True)
        self.rButton2.setChecked(True)
        self.rButton2.setObjectName("rButton2")

        self.textEdit = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit.setGeometry(QtCore.QRect(170, 0, 661, 401))
        #水平滚动条
        self.textEdit.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOn)
        self.textEdit.setObjectName("textEdit")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 838, 23))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.toolBar = QtWidgets.QToolBar(MainWindow)
        self.toolBar.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.toolBar.sizePolicy().hasHeightForWidth())
        self.toolBar.setSizePolicy(sizePolicy)
        self.toolBar.setMouseTracking(False)
        self.toolBar.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.toolBar.setAcceptDrops(True)
        self.toolBar.setAutoFillBackground(False)
        self.toolBar.setInputMethodHints(QtCore.Qt.ImhNone)
        self.toolBar.setAllowedAreas(QtCore.Qt.AllToolBarAreas)
        self.toolBar.setIconSize(QtCore.QSize(48, 48))
        self.toolBar.setToolButtonStyle(QtCore.Qt.ToolButtonTextUnderIcon)
        self.toolBar.setFloatable(False)
        self.toolBar.setObjectName("toolBar")
        MainWindow.addToolBar(QtCore.Qt.TopToolBarArea, self.toolBar)
        self.button1 = QtWidgets.QAction(MainWindow)
        icon = QtGui.QIcon.fromTheme("导入EXCEL")
        icon.addPixmap(QtGui.QPixmap("image/图标-01.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.button1.setIcon(icon)
        self.button1.setObjectName("button1")
        self.button2 = QtWidgets.QAction(MainWindow)
        #self.button2.setCheckable(True)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("image/图标-02.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.button2.setIcon(icon)
        self.button2.setObjectName("button2")
        self.button3 = QtWidgets.QAction(MainWindow)
        #self.button3.setCheckable(True)
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap("image/图标-03.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.button3.setIcon(icon1)
        self.button3.setObjectName("button3")
        self.button4 = QtWidgets.QAction(MainWindow)
        #self.button4.setCheckable(True)
        icon2 = QtGui.QIcon()
        icon2.addPixmap(QtGui.QPixmap("image/图标-04.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.button4.setIcon(icon2)
        self.button4.setObjectName("button4")
        self.button5 = QtWidgets.QAction(MainWindow)
        #self.button5.setCheckable(True)
        icon3 = QtGui.QIcon()
        icon3.addPixmap(QtGui.QPixmap("image/图标-05.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.button5.setIcon(icon3)
        self.button5.setObjectName("button5")
        self.button6 = QtWidgets.QAction(MainWindow)
        #self.button6.setCheckable(True)
        icon4 = QtGui.QIcon()
        icon4.addPixmap(QtGui.QPixmap("image/图标-06.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.button6.setIcon(icon4)
        self.button6.setObjectName("button6")
        self.button7 = QtWidgets.QAction(MainWindow)
        #self.button7.setCheckable(True)
        icon5 = QtGui.QIcon()
        icon5.addPixmap(QtGui.QPixmap("image/图标-07.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.button7.setIcon(icon5)
        self.button7.setObjectName("button7")
        self.toolBar.addSeparator()
        self.toolBar.addAction(self.button1)
        self.toolBar.addAction(self.button2)
        self.toolBar.addSeparator()
        self.toolBar.addSeparator()
        self.toolBar.addAction(self.button3)
        self.toolBar.addAction(self.button4)
        self.toolBar.addAction(self.button5)
        self.toolBar.addAction(self.button6)
        self.toolBar.addSeparator()
        self.toolBar.addSeparator()
        self.toolBar.addAction(self.button7)
        self.toolBar.addSeparator()
        # 单击工具栏“退出”按钮退出程序
        self.button7.triggered.connect(qApp.quit)
        # 单击工具栏按钮触发自定义槽函数
        self.button1.triggered.connect(self.click1)
        self.button2.triggered.connect(self.click2)
        self.button3.triggered.connect(self.click3)
        self.button4.triggered.connect(self.click4)
        self.button5.triggered.connect(self.click5)
        self.button6.triggered.connect(self.click6)
        #单击"浏览"按钮，选择文件存储路径
        self.viewButton.clicked.connect(self.viewButton_click)
        # 单击QListView列表触发自定义的槽函数
        self.list1.clicked.connect(self.clicked)
        # 设置Dataframe对象显示所有列
        pd.set_option('display.max_columns', None)
        # 设置Dataframe对象列宽为200，默认为50
        pd.set_option('max_colwidth', 200)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Excel数据分析师"))
        self.viewButton.setText(_translate("MainWindow", "浏览"))
        self.rButton1.setText(_translate("MainWindow", "自定义文件夹"))
        self.label.setText(_translate("MainWindow", "输出选项"))
        self.rButton2.setText(_translate("MainWindow", "保存在原文件夹内"))
        self.toolBar.setWindowTitle(_translate("MainWindow", "toolBar"))
        self.button1.setText(_translate("MainWindow", "导入EXCEL"))
        self.button2.setText(_translate("MainWindow", "提取列数据"))
        self.button3.setText(_translate("MainWindow", "定向筛选"))
        self.button3.setToolTip(_translate("MainWindow", "定向筛选"))
        self.button4.setText(_translate("MainWindow", "多表合并"))
        self.button5.setText(_translate("MainWindow", "多表统计排行"))
        self.button5.setToolTip(_translate("MainWindow", "多表统计排行"))
        self.button6.setText(_translate("MainWindow", "生成图表"))
        self.button6.setToolTip(_translate("MainWindow", "生成图表"))
        self.button7.setText(_translate("MainWindow", "退出"))
        self.button7.setToolTip(_translate("MainWindow", "退出"))


    def click1(self):
        # 文件夹路径
        global root
        root = QFileDialog.getExistingDirectory(self, "选择文件夹", "/")
        mylist = []
        # 遍历文件夹文件
        for dirpath, dirnames, filenames in os.walk(root):
            for filepath in filenames:
                # mylist.append(os.path.join(dirpath, filepath))
                mylist.append(os.path.join(filepath))
        # 实例化列表模型，添加数据列表
        self.model = QtCore.QStringListModel()
        # 添加列表数据
        self.model.setStringList(mylist)
        self.list1.setModel(self.model)
        self.list1 = mylist

    # 单击左侧目录右侧表格显示数据
    def clicked(self, qModelIndex):
        global root
        global myrow
        myrow=qModelIndex.row()
        # 获取当前选中行的数据
        a = root + '/' + str(self.list1[qModelIndex.row()])
        df = pd.DataFrame(pd.read_excel(a))
        self.textEdit.setText(str(df))

    #提取列数据
    def click2(self):
        global root
        global myrow
        # 获取当前选中行的数据
        a = root + '/' + str(self.list1[myrow])
        df = pd.DataFrame(pd.read_excel(a))
        #显示指定列数据
        df1 = df[['买家会员名', '收货人姓名', '联系手机','宝贝标题']]
        self.textEdit.setText(str(df1))
        #调用SaveExcel函数，保存数据到Excel
        SaveExcel(df1,self.rButton2.isChecked())

    #定向筛选
    def click3(self):
        global root
        global myrow
        #合并Excel表格
        filearray = []
        filelocation = glob.glob(root + "\\*.xlsx")
        for filename in filelocation:
            filearray.append(filename)
        res = pd.read_excel(filearray[0])
        for i in range(1, len(filearray)):
            A = pd.read_excel(filearray[i])
            res = pd.concat([res, A], ignore_index=False, sort=True)
        # 显示指定列数据
        df1 = res[['买家会员名', '收货人姓名', '联系手机','宝贝标题']]
        df2 = df1.loc[df1['宝贝标题'] == '零基础学Python']
        self.textEdit.setText(str(df2))
        #调用SaveExcel函数，保存定向筛选结果到Excel
        SaveExcel(df2,self.rButton2.isChecked())

    #多表合并
    def click4(self):
        global root
        # 合并指定文件夹下的所有Excel表
        filearray = []
        filelocation = glob.glob(root+"\\*.xlsx")
        for filename in filelocation:
            filearray.append(filename)
        res = pd.read_excel(filearray[0])
        for i in range(1, len(filearray)):
            A = pd.read_excel(filearray[i])
            res = pd.concat([res, A], ignore_index=False, sort=True)

        self.textEdit.setText(str(res.index))
        # 调用SaveExcel函数，将合并后的数据保存到Excel
        SaveExcel(res, self.rButton2.isChecked())

    #多表统计排行
    def click5(self):
        global root
        # 合并Excel表格
        filearray = []
        filelocation = glob.glob(root + "\\*.xlsx")
        for filename in filelocation:
            filearray.append(filename)
        res = pd.read_excel(filearray[0])
        for i in range(1, len(filearray)):
            A = pd.read_excel(filearray[i])
            res = pd.concat([res, A], ignore_index=False, sort=True)
        # 分组统计排序
        # 通过reset_index()函数将groupby()的分组结果转成DataFrame对象
        df = res.groupby(["宝贝标题"])["宝贝总数量"].sum().reset_index()
        df1 = df.sort_values(by='宝贝总数量', ascending=False)
        self.textEdit.setText(str(df1))
        # 调用SaveExcel函数，将统计排行结果保存到Excel
        SaveExcel(df1, self.rButton2.isChecked())

    def click6(self):
        global root
        # 合并Excel表格
        filearray = []
        filelocation = glob.glob(root + "\\*.xlsx")
        for filename in filelocation:
            filearray.append(filename)
        res = pd.read_excel(filearray[0])
        for i in range(1, len(filearray)):
            A = pd.read_excel(filearray[i])
            res = pd.concat([res, A], ignore_index=False, sort=True)
        # 分组统计排序
        # 通过reset_index()函数将groupby()的分组结果转成DataFrame对象
        df=res[(res.类别=='全彩系列')]
        df1 = df.groupby(["图书编号"])["买家实际支付金额"].sum().reset_index()
        df1 = df1.set_index('图书编号')  # 设置索引
        df1 = df1[u'买家实际支付金额'].copy()
        df2=df1.sort_values(ascending=False)  # 排序
        SaveExcel(df2, self.rButton2.isChecked())
        # 图表字体为华文细黑，字号为12
        plt.rc('font', family='SimHei', size=10)
        #plt.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号
        plt.figure("贡献度分析")
        df2.plot(kind='bar')
        plt.ylabel(u'销售收入（元）')
        p = 1.0*df2.cumsum()/df2.sum()
        print(p)
        p.plot(color='r', secondary_y=True, style='-o', linewidth=0.5)
        #plt.title("图书贡献度分析")
        plt.annotate(format(p[9], '.4%'), xy=(9, p[9]), xytext=(9 * 0.9, p[9] * 0.9),
                    arrowprops=dict(arrowstyle="->", connectionstyle="arc3,rad=.1"))  # 添加标记，并指定箭头样式。
        plt.ylabel(u'收入（比例）')
        plt.show()

    #单击“浏览”按钮选择文件存储路径
    def viewButton_click(self):
        global temproot
        temproot = QFileDialog.getExistingDirectory(self, "选择文件夹", "/")
        self.text1.setText(temproot)




# 定义载入主窗体的方法
def show_MainWindow():
        app = QtWidgets.QApplication(sys.argv)
        MainWindow = QtWidgets.QMainWindow()
        ui = Ui_MainWindow()
        ui.setupUi(MainWindow)
        MainWindow.show()
        sys.exit(app.exec_())


if __name__ == "__main__":
        show_MainWindow()
        path=root
        # visitDir(path)


# 冲突标记：feature1 分支修改
print('Feature1 branch change')
