#!/usr/bin/env python
# -*- coding: utf-8 -*- 
# @Time : 2020/4/12 23:29 
# @Author : DZQ
# @File : main.py
import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import *
import xlrd
from threading import Thread
import json
from BaiduIndexSpider import BaiduIndexSpider
from workbook_handle import WorkBook
import math
import time
import random


class MainThread(QThread):
    _my_signal = pyqtSignal(str)
    _keyword_signal = pyqtSignal(list)

    def __init__(self, keywords: list, filePath):
        super(MainThread, self).__init__()
        self.spider = BaiduIndexSpider(keywords)
        self.filePath = filePath
        self.keywords = keywords
        self.workbookNum = 1

    def split_keywords(self, keywords: list) -> [list]:
        return [keywords[i * 10: (i + 1) * 10] for i in range(math.ceil(len(keywords) / 10))]

    def spider_craw(self):
        self._my_signal.emit("正在进行第{}个爬虫".format(self.workbookNum))
        for each in self.spider.get_all_country():
            try:
                self.workbook.write_cell(each)
            except:
                pass
        self._my_signal.emit("已爬取完全国信息")
        self.workbook.init_province_cols()
        self._my_signal.emit("开始爬取各省市信息")
        year = 2011
        for each in self.spider.get_index():
            try:
                self.workbook.write_cell(each)
            except:
                pass
            try:
                date = int(each['date'].split("-")[0])
                if date > year:
                    self._my_signal.emit("爬取到{}年了".format(date))
                    year = date
            except:
                pass
        self._my_signal.emit("爬虫结束，正在保存excel")
        filePath = self.filePath + "/output{}.xls".format(self.workbookNum)
        self.workbookNum += 1
        self.workbook.workbook.save(filePath)
        self._my_signal.emit("保存Excel完成")

    def run(self) -> None:
        if not self.spider.is_login():
            self._my_signal.emit("Cookie过期")
            return
        real_keywords = list()
        self._my_signal.emit("正在判断关键词是否被收录")
        for each in self.keywords:
            if self.spider.is_keyword(each):
                real_keywords.append(each)
        if len(real_keywords) == 0:
            self._my_signal.emit("没有可以爬取的关键词")
            return
        self._keyword_signal.emit(real_keywords)
        self.keywords_list = self.split_keywords(real_keywords)
        self._my_signal.emit("关键词被分解成了{}个组\n".format(len(self.keywords_list)))
        self._my_signal.emit("开始爬虫")
        for each_keyword_list in self.keywords_list:
            self.workbook = WorkBook()
            self.spider.set_keywords(each_keyword_list)
            self.spider_craw()
            time.sleep(random.uniform(30, 35))


class Ui_MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(Ui_MainWindow, self).__init__()
        self.setupUi(self)
        self.retranslateUi(self)

    def open_file(self):
        filePath = QFileDialog.getOpenFileName(self, '选择文件', '', 'Excel files(*.xlsx , *.xls)')
        self.inputFilePath = filePath[0]
        self.filePathText.setPlainText(self.inputFilePath)
        filePathList = filePath[0].split("/")[:-1]
        outputFilePath = "/".join(filePathList)
        self.outputFilePath = outputFilePath
        self.get_keywords()

    def print_keyword(self, keywords):
        for each in keywords:
            self.msgBox.append(each)

    def handle_signal(self, info):
        self.msgBox.append(info)

    def start_spider(self):
        if len(self.keywords) == 0:
            self.msgBox.append("没有可以爬取的关键词")
            return
        self.thread = MainThread(self.keywords, self.outputFilePath)
        self.thread._my_signal.connect(self.handle_signal)
        self.thread._keyword_signal.connect(self.handle_list_signal)
        self.thread.start()

    def save_cookie(self):
        cookie = self.cookieText.toPlainText()
        if len(cookie) < 10:
            self.msgBox.append("Cookie信息太短")
            return
        config = json.loads(open("./config.json", "r", encoding="utf8").read())
        config['cookie'] = cookie
        json.dump(config, open("./config.json", "w", encoding="utf8"), ensure_ascii=False)
        self.msgBox.append("Cookie保存成功")

    def handle_list_signal(self, info):
        self.msgBox.append("获取到以下可以爬取的关键词：")
        thread = Thread(target=self.print_keyword, args=(info,))
        thread.start()
        thread.join()
        self.msgBox.append("共获得{}个被收录的关键词".format(len(info)))

    def get_keywords(self):
        excelFile = xlrd.open_workbook(self.inputFilePath)
        sheet = excelFile.sheet_by_index(0)
        row_num = sheet.nrows
        keywords = list()
        for i in range(row_num):
            value = str(sheet.cell_value(i, 0)).strip()
            if len(value) > 0:
                keywords.append(value)
        self.keywords = keywords
        thread = Thread(target=self.print_keyword, args=(keywords,))
        thread.start()
        thread.join()
        self.msgBox.append("共获取到{}个关键词".format(len(keywords)))

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1050, 744)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.filePathLabel = QtWidgets.QLabel(self.centralwidget)
        self.filePathLabel.setGeometry(QtCore.QRect(30, 50, 101, 51))
        font = QtGui.QFont()
        font.setFamily("楷体")
        font.setPointSize(12)
        self.filePathLabel.setFont(font)
        self.filePathLabel.setObjectName("filePathLabel")
        self.filePathText = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.filePathText.setGeometry(QtCore.QRect(180, 50, 631, 51))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(12)
        self.filePathText.setFont(font)
        self.filePathText.setObjectName("filePathText")
        self.filePathBtn = QtWidgets.QPushButton(self.centralwidget)
        self.filePathBtn.setGeometry(QtCore.QRect(830, 60, 141, 41))
        font = QtGui.QFont()
        font.setFamily("楷体")
        font.setPointSize(12)
        self.filePathBtn.setFont(font)
        self.filePathBtn.setObjectName("filePathBtn")
        self.startSpiderBtn = QtWidgets.QPushButton(self.centralwidget)
        self.startSpiderBtn.setGeometry(QtCore.QRect(390, 150, 201, 61))
        font = QtGui.QFont()
        font.setFamily("楷体")
        font.setPointSize(12)
        self.startSpiderBtn.setFont(font)
        self.startSpiderBtn.setObjectName("startSpiderBtn")
        self.cookieLabel = QtWidgets.QLabel(self.centralwidget)
        self.cookieLabel.setGeometry(QtCore.QRect(40, 290, 81, 41))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(12)
        self.cookieLabel.setFont(font)
        self.cookieLabel.setObjectName("cookieLabel")
        self.cookieText = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.cookieText.setGeometry(QtCore.QRect(180, 270, 631, 81))
        self.cookieText.setObjectName("cookieText")
        self.cookieBtn = QtWidgets.QPushButton(self.centralwidget)
        self.cookieBtn.setGeometry(QtCore.QRect(840, 290, 141, 41))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(12)
        self.cookieBtn.setFont(font)
        self.cookieBtn.setObjectName("cookieBtn")
        self.msgBox = QtWidgets.QTextBrowser(self.centralwidget)
        self.msgBox.setGeometry(QtCore.QRect(160, 380, 681, 301))
        font = QtGui.QFont()
        font.setFamily("楷体")
        font.setPointSize(12)
        self.msgBox.setFont(font)
        self.msgBox.setObjectName("msgBox")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1352, 30))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        self.filePathBtn.clicked.connect(self.open_file)
        self.cookieBtn.clicked.connect(self.save_cookie)
        self.startSpiderBtn.clicked.connect(self.start_spider)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "百度指数爬虫"))
        self.filePathLabel.setText(_translate("MainWindow", "文件目录"))
        self.filePathBtn.setText(_translate("MainWindow", "选择文件"))
        self.startSpiderBtn.setText(_translate("MainWindow", "启动爬虫"))
        self.cookieLabel.setText(_translate("MainWindow", "Cookie"))
        self.cookieBtn.setText(_translate("MainWindow", "更新Cookie"))


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
