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
        self.birthday_date = self.ui.birthday_dateEdit.date()
        self.birth = [
            self.birthday_date.year(),
            self.birthday_date.month(),
            self.birthday_date.day()
        ]
        # 绑定事件
        self.ui.generate_pushButton.clicked.connect(self.test)
        self.ui.gender_comboBox.currentIndexChanged.connect(
            self.set_retire_age)
        self.ui.retire_age_comboBox.currentIndexChanged.connect(
            self.calculate_time_to_retire)

        # 设置初始值
        self.ui.gender_comboBox.addItems(["男", "女"])

    def test(self):
        birthday_date = self.ui.birthday_dateEdit.date()

    def set_retire_age(self):
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
        if self.ui.retire_age_comboBox.currentText():
            self.retire_age = self.birth.copy()
            self.retire_age[0] += int(
                self.ui.retire_age_comboBox.currentText())
            self.ui.time_to_retire_lineEdit.setText(str(self.retire_age))

    def time_lefted_to_retire(self, gender, birthday, retire_age):
        pass


if __name__ == "__main__":
    app = QApplication(sys.argv)
    MainWindow = Main()
    sys.exit(app.exec_())
