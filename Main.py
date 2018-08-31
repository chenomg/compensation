#!/usr/bin/env python3
# -*- coding: utf-8 -*-
'''
# =============================================================================
#      FileName: Main.py
#          Desc:
#        Author: Jase Chen
#         Email: xxmm@live.cn
#      HomePage: https://jase.im/
#       Version: 0.0.1
#       License: GPLv2
#    LastChange: 2018-08-30 21:31:51
#       History:
# =============================================================================
'''

from mainwindow import Ui_MainWindow
from PyQt5.QtWidgets import QMainWindow, QMessageBox, QApplication, QInputDialog, QLineEdit
from PyQt5 import QtGui
from PyQt5.QtGui import QIcon, QPixmap
from time import sleep
import datetime
import sqlite3
import sys
# import logging
import xlrd
import xlwt
import os
import datetime
import re


class Main(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.show()

        # 对象属性
        # 绑定事件
        self.ui.generate_pushButton.clicked.connect(self.test)
        self.ui.gender_comboBox.currentIndexChanged.connect(
            self.set_retire_age)
        self.ui.retire_age_comboBox.currentIndexChanged.connect(
            self.calculate_time_to_retire)
        self.ui.birthday_dateEdit.dateChanged.connect(
            self.calculate_time_to_retire)

        # 设置初始值
        self.ui.gender_comboBox.addItems(["男", "女"])

    def test(self):
        birthday_date = self.ui.birthday_dateEdit.date()

    def set_retire_age(self):
        """
        根据性别更新退休年龄
        """
        G_Combo = self.ui.gender_comboBox
        Retire_age_Combo = self.ui.retire_age_comboBox
        if G_Combo.currentText() == "男":
            Retire_age_Combo.clear()
            Retire_age_Combo.addItems(["60"])
        if G_Combo.currentText() == "女":
            Retire_age_Combo.clear()
            Retire_age_Combo.addItems(["50", "55"])
        self.calculate_time_to_retire()

    def calculate_time_to_retire(self):
        """
        更新程序中距离退休时间一栏
        """
        self.birthday_date = self.ui.birthday_dateEdit.date()
        self.birth = [
            self.birthday_date.year(),
            self.birthday_date.month(),
            self.birthday_date.day()
        ]
        self.birth_date = datetime.date(self.birth[0], self.birth[1],
                                        self.birth[2])
        if self.ui.retire_age_comboBox.currentText():
            self.retire_age = self.birth.copy()
            self.retire_age[0] += int(
                self.ui.retire_age_comboBox.currentText())
            self.retire_date = datetime.date(
                self.retire_age[0], self.retire_age[1], self.retire_age[2])
            today = datetime.date.today()
            retire_date = self.time_to_retire(self.retire_date, today)
            self.ui.time_to_retire_lineEdit.setText((str(
                retire_date[0]) + "年" + str(retire_date[1]) + "月").rjust(10))
            if retire_date[0] < 5:
                self.ui.time_to_retire_lineEdit.setStyleSheet("color:red")
            else:
                self.ui.time_to_retire_lineEdit.setStyleSheet("color:black")

    def time_to_retire(self, retire_date, today_date):
        """
        计算离退休还有多久
        input: 退休日期(retire_date), 今天日期(today_date)
        type:  datetime.date
        return:delta_year, delta_month
        type:  datetime.date
        """
        delta_day = retire_date.day - today_date.day
        delta_month = retire_date.month - today_date.month
        delta_year = retire_date.year - today_date.year
        if delta_day > 0:
            delta_month += 1
        if delta_month < 0:
            delta_month += 12
            delta_year -= 1
        return delta_year, delta_month

    def time_lefted_to_retire(self, gender, birthday, retire_age):
        pass


if __name__ == "__main__":
    app = QApplication(sys.argv)
    MainWindow = Main()
    sys.exit(app.exec_())
