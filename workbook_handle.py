#!/usr/bin/env python
# -*- coding: utf-8 -*- 
# @Time : 2020/4/14 21:57 
# @Author : DZQ
# @File : workbook_handle.py

import xlwt
import json
import datetime


class WorkBook:
    def __init__(self):
        self.workbook = xlwt.Workbook()
        self.area_to_sheet = dict()
        file = open("./config.json", "r", encoding="utf8")
        self.province = dict(json.loads(file.read())['province'])
        self.keywords = dict()
        file.close()
        self.init_workbook()
        self.totalKeyworkNum = 0
        self.init_province_cols()

    def init_workbook(self):
        country_sheet = self.workbook.add_sheet("全国", True)
        self.area_to_sheet['0'] = country_sheet
        for each in self.province.keys():
            sheet = self.workbook.add_sheet(self.province[each], True)
            self.area_to_sheet[each] = sheet
        self.creat_first_col()

    def set_style(self, name, height, bold=False):
        style = xlwt.XFStyle()
        font = xlwt.Font()
        font.name = name
        font.bold = bold
        font.color_index = 4
        font.height = height
        style.font = font
        return style

    def creat_first_col(self):
        sheet = self.area_to_sheet['0']
        start_date = datetime.datetime.strptime("2006-06-01", "%Y-%m-%d")
        end_date = datetime.datetime.strptime("2010-12-31", "%Y-%m-%d")
        num = 2
        style = self.set_style('Times New Roman', 220, True)
        self.titleStyle = style
        while start_date <= end_date:
            sheet.write(num, 0, start_date.strftime("%Y-%m-%d"), style)
            start_date = start_date + datetime.timedelta(days=1)
            num += 1
        for each in self.area_to_sheet.keys():
            if each == '0':
                continue
            sheet = self.area_to_sheet[each]
            start_date = datetime.datetime.strptime("2011-01-01", "%Y-%m-%d")
            end_date = datetime.datetime.now()
            num = 2
            style = self.set_style('Times New Roman', 220, True)
            self.titleStyle = style
            while start_date <= end_date:
                sheet.write(num, 0, start_date.strftime("%Y-%m-%d"), style)
                start_date = start_date + datetime.timedelta(days=1)
                num += 1

    def init_province_cols(self):
        for each_sheet in self.area_to_sheet.keys():
            if each_sheet == '0':
                continue
            sheet = self.area_to_sheet[each_sheet]
            for each in self.keywords.keys():
                num = self.keywords[each]
                sheet.write_merge(0, 0, num * 3 + 1, num * 3 + 3, each)
                sheet.write(1, num * 3 + 1, "pc")
                sheet.write(1, num * 3 + 2, "wise")
                sheet.write(1, num * 3 + 3, "all")

    def write_cell(self, data: dict):
        date = datetime.datetime.strptime(data['date'], "%Y-%m-%d")
        year = int(date.year)
        if year < 2011:
            keyword = data['keyword']
            sheet = self.area_to_sheet['0']
            if keyword in self.keywords.keys():
                num = self.keywords[keyword]
            else:
                num = self.totalKeyworkNum
                self.keywords[keyword] = self.totalKeyworkNum
                self.totalKeyworkNum += 1
                sheet.write_merge(0, 0, num * 3 + 1, num * 3 + 3, keyword)
                sheet.write(1, num * 3 + 1, "pc")
                sheet.write(1, num * 3 + 2, "wise")
                sheet.write(1, num * 3 + 3, "all")
            date = data['date']
            rowNum = (datetime.datetime.strptime(date, "%Y-%m-%d") - datetime.datetime.strptime("2006-06-01",
                                                                                                "%Y-%m-%d")).days + 2
            index = int(data['index'])
            sheet.write(rowNum, num * 3 + 1, index)
            sheet.write(rowNum, num * 3 + 2, 0)
            sheet.write(rowNum, num * 3 + 3, index)
        else:
            area = data['area']
            sheet = self.area_to_sheet[area]
            keyword = data['keyword']
            type = data['type']
            date = data['date']
            index = int(data['index'])
            rowNum = (datetime.datetime.strptime(date, "%Y-%m-%d") - datetime.datetime.strptime("2011-01-01",
                                                                                                "%Y-%m-%d")).days + 2
            num = self.keywords[keyword]
            if type == 'pc':
                colNum = num * 3 + 1
            elif type == 'wise':
                colNum = num * 3 + 2
            else:
                colNum = num * 3 + 3
            sheet.write(rowNum, colNum, index)
