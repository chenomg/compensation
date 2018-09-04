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
import datetime
import calendar
import sys
# import logging
import xlrd
import xlwt
import os


class Compensation(QMainWindow):
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
        # 批量输出按钮
        self.ui.xls_calculate_pushButton.clicked.connect(
            self.xls_calculate_pushButton_clicked)
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
        self.update_compensation_money()

    def generate_pushButton_clicked(self):
        name = self.ui.name_lineEdit.text()
        gender = self.ui.gender_comboBox.currentText()
        birth = Compensation.dateEdit_to_dateStr(self.ui.birthday_dateEdit)
        retire_age = self.ui.retire_age_comboBox.currentText()
        working_start = Compensation.dateEdit_to_dateStr(
            self.ui.birthday_dateEdit)
        year_sus = self.ui.working_suspended_year_lineEdit.text().zfill(1)
        month_sus = self.ui.working_suspended_month_lineEdit.text().zfill(1)
        company_in = Compensation.dateEdit_to_dateStr(
            self.ui.company_in_dateEdit)
        company_out = Compensation.dateEdit_to_dateStr(
            self.ui.company_out_dateEdit)
        personal_average = self.ui.personal_average_lineEdit.text()
        society_average = self.ui.society_average_lineEdit.text()
        time_to_retire = self.ui.time_to_retire_lineEdit.text()
        working_years_Y_M = self.ui.working_years_lineEdit.text()
        compensation_mon_bef = self.ui.compensation_mon_bef_lineEdit.text()
        compensation_mon_aft = self.ui.compensation_mon_aft_lineEdit.text()
        compensation = self.ui.compensation_money_lineEdit.text()
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet("赔偿金登记表")
        row0 = [
            '序号',
            '姓名',
            '性别',
            '生日',
            '退休年龄',
            '个人工作开始时间',
            '未上班累计时间(年)',
            '未上班累计时间(月)',
            '进本单位时间',
            '从本单位离职时间',
            '上年度个人平均工资(元)',
            '上年度社会平均工资三倍(元)',
            '离退休时间',
            '累计工龄',
            '赔偿月数(08年前)',
            '赔偿月数(08年及以后)',
            '赔偿金额(元)',
        ]
        for i in range(len(row0)):
            sheet.write(0, i, row0[i])
        row1 = [
            1,
            name,
            gender,
            birth,
            retire_age,
            working_start,
            year_sus,
            month_sus,
            company_in,
            company_out,
            personal_average,
            society_average,
            time_to_retire,
            working_years_Y_M,
            compensation_mon_bef,
            compensation_mon_aft,
            compensation[:len(compensation) - 1],
        ]
        for i in range(len(row1)):
            sheet.write(1, i, row1[i])
        workbook.save(os.path.join(os.getcwd(), "赔偿金登记表-{}.xls".format(name)))
        info = QMessageBox.information(self, "信息", "个人数据导出成功", QMessageBox.Ok)

    def xls_calculate_pushButton_clicked(self):
        """
        批量计算data.xls文件中的数据
        """
        try:
            workbook = xlrd.open_workbook('data.xls')
            sheet = workbook.sheet_by_index(0)
            sheet_rows = sheet.nrows
            row = []
            for i in range(sheet_rows):
                row.append(sheet.row_values(i))
            for i in range(1, len(row)):
                for j in [6, 7]:
                    if not row[i][j]:
                        row[i][j] = 0
                birth_date = Compensation.dateStr_to_date(row[i][3])
                retire_age = int(row[i][4])
                retire_date = Compensation.calculate_time_to_retire(
                    birth_date, retire_age)
                # 离退休时间
                row[i][12] = "{}年{}月".format(retire_date[0], retire_date[1])
                working_start_date = Compensation.dateStr_to_date(row[i][5])
                year_sus = int(row[i][6])
                month_sus = int(row[i][7])
                working_years_Y_M = Compensation.calculate_working_years(
                    working_start_date, year_sus, month_sus)
                # 累计工龄
                row[i][13] = "{}年{}月".format(working_years_Y_M[0],
                                             working_years_Y_M[1])
                company_in_date = Compensation.dateStr_to_date(row[i][8])
                company_out_date = Compensation.dateStr_to_date(row[i][9])
                personal_average = float(row[i][10])
                society_average = float(row[i][11])
                compensation = Compensation.calculate_compensation_money(
                    company_in_date, company_out_date, personal_average,
                    society_average)
                # 08年前赔偿金额
                row[i][14] = "{}".format(compensation[1])
                # 08年后赔偿金额
                row[i][15] = "{}".format(compensation[2])
                # 赔偿金额
                row[i][16] = "{}".format(compensation[0])
            data = xlwt.Workbook()
            table = data.add_sheet("赔偿金登记表")
            for i in range(len(row)):
                for j in range(len(row[i])):
                    table.write(i, j, row[i][j])
            data.save('data_new.xls')
            info = QMessageBox.information(self, "信息", "数据导出成功",
                                           QMessageBox.Ok)
        except Exception as e:
            QMessageBox.information(self, "信息", str(e), QMessageBox.Ok)

    @staticmethod
    def dateStr_to_date(dateStr):
        """
        将日期(例:20010101)转换为datetime.date类
        """
        year = int(dateStr[:4])
        month = int(dateStr[4:6])
        day = int(dateStr[6:])
        return datetime.date(year, month, day)

    @staticmethod
    def dateEdit_to_dateStr(dateEdit):
        """
        将QdateEdit转换为日期(例:20010101)
        """
        date = dateEdit.date()
        year = str(date.year())
        month = str(date.month()).zfill(2)
        day = str(date.day()).zfill(2)
        return year + month + day

    def about_pushButton_clicked(self):
        msgBox = QMessageBox.information(
            self, "About Me...",
            "Thanks for using :D\n\nComposed by Jase Chen\n\n2018-09-03",
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
        if self.ui.retire_age_comboBox.currentText():
            birthday_date = self.ui.birthday_dateEdit.date()
            birth_date = self.dateEdit_to_date(birthday_date)
            retire_age = int(self.ui.retire_age_comboBox.currentText())
            retire_date_Y_M = Compensation.calculate_time_to_retire(
                birth_date, retire_age)
            self.ui.time_to_retire_lineEdit.setText(
                str(retire_date_Y_M[0]) + "年" + str(retire_date_Y_M[1]) + "月")
            if retire_date_Y_M[0] < 5:
                self.ui.time_to_retire_lineEdit.setStyleSheet("color:red")
            else:
                self.ui.time_to_retire_lineEdit.setStyleSheet("color:black")

    @staticmethod
    def calculate_time_to_retire(birth_date, retire_age):
        """
        计算距离退休的时间,月份向上取整
        """
        birth = [birth_date.year, birth_date.month, birth_date.day]
        retire_date_ = birth.copy()
        retire_date_[0] += retire_age
        if retire_date_[2] != 1:
            retire_date_[2] -= 1
            retire_date = datetime.date(retire_date_[0], retire_date_[1],
                                        retire_date_[2])
        else:
            retire_date = datetime.date(
                retire_date_[0], retire_date_[1],
                retire_date_[2]) - datetime.timedelta(days=1)
        today = datetime.date.today()
        retire_date_Y_M = Compensation.calculate_time_delta(today, retire_date)
        return retire_date_Y_M

    @staticmethod
    def calculate_time_delta(start_date, end_date, lower=False, check_=False):
        """
        计算时间间隔,返回年,月,区间包含前后两天, 默认不满一月的按一个月计算
        input: 开始日期(start_date), 结束日期(end_date), \
                去除不满一月的天数(lower), 检查最后是否刚满一个月,两种情况(check_)
        type:  datetime.date
        return:delta_year, delta_month
        type:  datetime.date
        """

        def is_Full_one_month():
            nonlocal start_date
            nonlocal end_date
            is_Full_one = False
            end_month_end_date = datetime.date(
                end_date.year, end_date.month,
                calendar.monthrange(end_date.year, end_date.month)[1])
            if start_date.day == 1:
                if end_date.day == end_month_end_date.day:
                    is_Full_one = True
            return is_Full_one

        def is_Full_cross_month():
            nonlocal start_date
            nonlocal end_date
            is_Full_cross_month = False
            if start_date.day - 1 == end_date.day:
                is_Full_cross_month = True
            return is_Full_cross_month

        delta_day = end_date.day - start_date.day
        delta_month = end_date.month - start_date.month
        delta_year = end_date.year - start_date.year
        if lower:
            if delta_day < -1:
                delta_month -= 1
            if is_Full_one_month():
                delta_month += 1
        else:
            if delta_day >= 0:
                delta_month += 1
            # if is_Full_one_month():
            # delta_month += 1
        if delta_month < 0:
            delta_month += 12
            delta_year -= 1
        if delta_month == 12:
            delta_month = 0
            delta_year += 1
        if not check_:
            return delta_year, delta_month
        if check_:
            return delta_year, delta_month, is_Full_one_month(
            ) or is_Full_cross_month()

    def update_working_years(self):
        """
        更新工龄数值,月份向下取整
        """
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
        delta_year, delta_month = Compensation.calculate_working_years(
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
        time_delta = Compensation.calculate_time_delta(
            working_start_date, today, lower=True)
        delta_year = time_delta[0] - year_sus
        delta_month = time_delta[1] - month_sus
        if delta_month < 0:
            delta_month += 12
            delta_year -= 1
        return delta_year, delta_month

    def update_compensation_money(self):
        """
        更新赔偿金数据
        """
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
        try:
            personal_average = float(personal_average_str)
        except Exception as e:
            personal_average = 0
        try:
            society_average = float(society_average_str)
        except Exception as e:
            society_average = 0
        compensation_money = Compensation.calculate_compensation_money(
            company_in_date, company_out_date, personal_average,
            society_average)
        self.ui.compensation_money_lineEdit.setText(
            str(compensation_money[0]) + "元")
        self.ui.compensation_mon_bef_lineEdit.setText(
            str(compensation_money[1]) + "月")
        self.ui.compensation_mon_aft_lineEdit.setText(
            str(compensation_money[2]) + "月")

    @staticmethod
    def calculate_compensation_money(company_in_date, company_out_date,
                                     personal_average, society_average):
        """
        计算赔偿金数据,包括08年以前及之后的赔偿金数据,目前计算是两者最大各12个月
        """
        month_before = 0
        month_after = 0

        def money_before(company_in_date, company_out_date, personal_average):
            nonlocal month_before
            if company_in_date.year < 2008:
                if company_out_date.year > 2007:
                    company_out_date = datetime.date(2007, 12, 31)
                date_delta = Compensation.calculate_time_delta(
                    company_in_date, company_out_date)
                if date_delta[0] > 11 or (date_delta[0] == 11
                                          and not date_delta[1]):
                    month_before = 12
                    return personal_average * 12
                else:
                    if date_delta[1]:
                        month_before = date_delta[0] + 1
                        return personal_average * (date_delta[0] + 1)
                    else:
                        month_before = date_delta[0]
                        return personal_average * date_delta[0]
            else:
                return 0

        def money_after(company_in_date, company_out_date, personal_average,
                        society_average):
            nonlocal month_after
            if company_out_date.year > 2007:
                if company_in_date.year < 2008:
                    company_in_date = datetime.date(2008, 1, 1)
                date_delta = Compensation.calculate_time_delta(
                    company_in_date, company_out_date, check_=True)
                if date_delta[1] > 6:
                    month_avail = date_delta[0] + 1
                elif date_delta[1] == 6:
                    if date_delta[2]:
                        month_avail = date_delta[0] + 1
                    else:
                        month_avail = date_delta[0] + 0.5
                elif 0 < date_delta[1] < 6:
                    month_avail = date_delta[0] + 0.5
                else:
                    month_avail = date_delta[0]
                if personal_average > society_average:
                    if month_avail >= 12:
                        month_after = 12
                        return society_average * 12
                    else:
                        month_after = month_avail
                        return society_average * month_avail
                else:
                    month_after = month_avail
                    return personal_average * month_avail
            else:
                return 0

        compensation_money = money_before(
            company_in_date, company_out_date, personal_average) + money_after(
                company_in_date, company_out_date, personal_average,
                society_average)
        return compensation_money, month_before, month_after


if __name__ == "__main__":
    app = QApplication(sys.argv)
    MainWindow = Compensation()
    sys.exit(app.exec_())
