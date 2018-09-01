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
from PyQt5.QtGui import QIcon, QPixmap, QIntValidator, QDoubleValidator
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


class Compensation_Money(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.show()

        # 对象属性, 对象输入内容限制
        validator_year = QIntValidator(1, 60, self)
        validator_month = QIntValidator(1, 11, self)
        validator_personal_ave = QDoubleValidator(1.0, 100000, 2)
        validator_society_ave = QDoubleValidator(1.0, 100000, 2)
        self.ui.working_suspended_year_lineEdit.setValidator(validator_year)
        self.ui.working_suspended_month_lineEdit.setValidator(validator_month)
        self.ui.personal_average_lineEdit.setValidator(validator_personal_ave)
        self.ui.society_average_lineEdit.setValidator(validator_society_ave)

        # 绑定事件
        # 输出按钮
        self.ui.generate_pushButton.clicked.connect(
            self.generate_pushButton_clicked)
        # 关于按钮
        self.ui.about_pushButton.clicked.connect(self.about_pushButton_clicked)
        # 更新退休年龄
        self.ui.gender_comboBox.currentIndexChanged.connect(
            self.set_retire_age)
        # 更新距离退休时间
        self.ui.retire_age_comboBox.currentIndexChanged.connect(
            self.update_time_to_retire)
        self.ui.birthday_dateEdit.dateChanged.connect(
            self.update_time_to_retire)
        # 累计工龄时间绑定
        self.ui.working_start_dateEdit.dateChanged.connect(
            self.update_working_years)
        self.ui.working_suspended_year_lineEdit.textChanged.connect(
            self.update_working_years)
        self.ui.working_suspended_month_lineEdit.textChanged.connect(
            self.update_working_years)
        # 赔偿金额绑定
        self.ui.personal_average_lineEdit.textChanged.connect(
            self.update_compensation_money)
        self.ui.society_average_lineEdit.textChanged.connect(
            self.update_compensation_money)
        self.ui.company_in_dateEdit.dateChanged.connect(
            self.update_compensation_money)
        self.ui.company_out_dateEdit.dateChanged.connect(
            self.update_compensation_money)

        # 设置初始值
        self.ui.gender_comboBox.addItems(["男", "女"])
        self.update_working_years()

    def generate_pushButton_clicked(self):
        pass

    def about_pushButton_clicked(self):
        msgBox = QMessageBox.information(
            self, "About Me...",
            "Thanks for using :D\nComposed by Jase Chen\n2018-09-01",
            QMessageBox.Yes)

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
        self.update_time_to_retire()

    def update_time_to_retire(self):
        """
        更新程序中距离退休时间一栏
        """
        birthday_date = self.ui.birthday_dateEdit.date()
        birth_date = self.dateEdit_to_date(birthday_date)
        birth = [birth_date.year, birth_date.month, birth_date.day]
        if self.ui.retire_age_comboBox.currentText():
            retire_age = birth
            retire_age[0] += int(self.ui.retire_age_comboBox.currentText())
            retire_date = datetime.date(retire_age[0], retire_age[1],
                                        retire_age[2])
            today = datetime.date.today()
            retire_date_Y_M = Compensation_Money.calculate_time_delta(
                today, retire_date)
            self.ui.time_to_retire_lineEdit.setText(
                str(retire_date_Y_M[0]) + "年" + str(retire_date_Y_M[1]) + "月")
            if retire_date_Y_M[0] < 5:
                self.ui.time_to_retire_lineEdit.setStyleSheet("color:red")
            else:
                self.ui.time_to_retire_lineEdit.setStyleSheet("color:black")

    @staticmethod
    def calculate_time_delta(start_date, end_date, lower=False):
        """
        计算时间间隔,返回年,月,区间包含前后两天, 默认不满一月的按一个月计算
        input: 开始日期(start_date), 结束日期(end_date), 去除不满一月的天数(lower)
        type:  datetime.date
        return:delta_year, delta_month
        type:  datetime.date
        """
        delta_day = end_date.day - start_date.day
        delta_month = end_date.month - start_date.month
        delta_year = end_date.year - start_date.year
        if lower:
            if delta_day < 0:
                delta_month -= 1
        else:
            if delta_day >= 0:
                delta_month += 1
        if delta_month < 0:
            delta_month += 12
            delta_year -= 1
        if delta_month == 12:
            delta_month = 0
            delta_year += 1
        return delta_year, delta_month

    def update_working_years(self):
        date = self.ui.working_start_dateEdit.date()
        working_start_date = self.dateEdit_to_date(date)
        # 未上班时间
        year_sus_str = self.ui.working_suspended_year_lineEdit.text()
        month_sus_str = self.ui.working_suspended_month_lineEdit.text()
        year_sus = 0
        month_sus = 0
        if year_sus_str:
            year_sus = int(year_sus_str)
        if month_sus_str:
            month_sus = int(month_sus_str)
        delta_year, delta_month = Compensation_Money.calculate_working_years(
            working_start_date, year_sus, month_sus)
        self.ui.working_years_lineEdit.setText("{}年{}月".format(
            delta_year, delta_month))

    def dateEdit_to_date(self, dateEdit):
        """Description

        @param dateEdit: QT5中dateEdit控件
        @type  dateEdit:  dateEdit.date()

        @return:  date
        @rtype :  datetime.date()

        @raise e:  Description
        """
        date = datetime.date(dateEdit.year(), dateEdit.month(), dateEdit.day())
        return date

    @staticmethod
    def calculate_working_years(working_start_date, year_sus, month_sus):
        """Description

        @param working_start_date, year_sus, month_sus: 工作开始日期, 间隔年月
        @type  working_start_date: datetime.date, int, int

        @return: working years and months
        @rtype : tuple

        @raise e:  Description
        """
        today = datetime.date.today()
        time_delta = Compensation_Money.calculate_time_delta(
            working_start_date, today, lower=True)
        delta_year = time_delta[0] - year_sus
        delta_month = time_delta[1] - month_sus
        if delta_month < 0:
            delta_month += 12
            delta_year -= 1
        return delta_year, delta_month

    def update_compensation_money(self):
        company_in_date = self.dateEdit_to_date(
            self.ui.company_in_dateEdit.date())
        company_out_date = self.dateEdit_to_date(
            self.ui.company_out_dateEdit.date())
        personal_average_str = self.ui.personal_average_lineEdit.text()
        society_average_str = self.ui.society_average_lineEdit.text()
        if not personal_average_str:
            personal_average_str = '0'
        if not society_average_str:
            society_average_str = '0'
        personal_average = float(personal_average_str)
        society_average = float(society_average_str)
        compensation_money = Compensation_Money.calculate_compensation_money(
            company_in_date, company_out_date, personal_average,
            society_average)
        self.ui.compensation_money_lineEdit.setText(str(compensation_money) + "元")

    @staticmethod
    def calculate_compensation_money(company_in_date, company_out_date,
                                     personal_average, society_average):
        def money_before(company_in_date, company_out_date, personal_average):
            if company_in_date.year < 2008:
                if company_out_date.year > 2007:
                    company_out_date = datetime.date(2007, 12, 31)
                date_delta = Compensation_Money.calculate_time_delta(
                    company_in_date, company_out_date)
                if date_delta[0] > 11 or (date_delta[0] == 11
                                          and not date_delta[1]):
                    return personal_average * 12
                else:
                    if date_delta[1]:
                        return personal_average * (date_delta[0] + 1)
                    else:
                        return personal_average * date_delta[0]
            else:
                return 0

        def money_after(company_in_date, company_out_date, personal_average,
                        society_average):
            if company_out_date.year > 2007:
                company_in_date = datetime.date(2008, 1, 1)
                date_delta = Compensation_Money.calculate_time_delta(
                    company_in_date, company_out_date)
                if date_delta[1] > 6:
                    month_avail = date_delta[0] + 1
                elif 0 < date_delta[1] < 7:
                    month_avail = date_delta[0] + 0.5
                else:
                    month_avail = date_delta[0]
                if personal_average > society_average:
                    if month_avail >= 12:
                        return society_average * 12
                    else:
                        return society_average * month_avail
                else:
                    return personal_average * month_avail
            else:
                return 0

        compensation_money = money_before(
            company_in_date, company_out_date, personal_average) + money_after(
                company_in_date, company_out_date, personal_average,
                society_average)
        return compensation_money


if __name__ == "__main__":
    app = QApplication(sys.argv)
    MainWindow = Compensation_Money()
    sys.exit(app.exec_())
