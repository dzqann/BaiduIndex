#!/usr/bin/env python
# -*- coding: utf-8 -*- 
# @Time : 2020/4/13 16:44 
# @Author : DZQ
# @File : BaiduIndexSpider.py

from urllib.parse import urlencode
import math
from queue import Queue
import datetime
import random
import time
import json
import requests
from lxml import etree

headers = {
    'Host': 'index.baidu.com',
    'Connection': 'keep-alive',
    'X-Requested-With': 'XMLHttpRequest',
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.77 Safari/537.36',
}


class BaiduIndexSpider():
    def __init__(self, keywords: list):
        self.keywords = keywords
        self.params_queue = Queue()
        config_file = open("./config.json", "r", encoding="utf8").read()
        self.cookie = json.loads(config_file)['cookie']
        self.area = 0
        self.province = dict(json.loads(config_file)['province'])

    def set_keywords(self, keywords):
        self.keywords = keywords

    def is_login(self):
        req = self.get("https://www.baidu.com")
        html = etree.HTML(req)
        login = html.xpath("//a[@name='tj_login']")
        return len(login) == 0

    def is_keyword(self, keyword):
        result = self.get("http://index.baidu.com/api/AddWordApi/checkWordsExists?word={}".format(keyword))
        try:
            result = json.loads(result)
            data = result['data']['result']
            return len(data) == 0
        except:
            return False

    def get_index(self):
        self.init_queue("2011-01-01", datetime.datetime.now().strftime("%Y-%m-%d"), self.keywords)
        while True:
            try:
                if self.params_queue.empty():
                    break
                params_data = self.params_queue.get(timeout=1)
                for each_area in self.province.keys():
                    self.area = each_area
                    encrpt_datas, uniqid = self.get_encrypt_datas(
                        start_date=params_data['start_date'],
                        end_date=params_data['end_date'],
                        keywords=params_data['keywords']
                    )
                    key = self.get_key(uniqid)
                    for encrypt_data in encrpt_datas:
                        encrypt_data['pc']['data'] = self.decrypt_func(key, encrypt_data['pc']['data'])
                        encrypt_data['all']['data'] = self.decrypt_func(key, encrypt_data['all']['data'])
                        encrypt_data['wise']['data'] = self.decrypt_func(key, encrypt_data['wise']['data'])
                        for formated_data in self.format_data(encrypt_data):
                            yield formated_data
                    time.sleep(random.uniform(2, 3))
                time.sleep(random.uniform(9, 15))
            except requests.Timeout:
                self.params_queue.put(params_data)

    def init_queue(self, start_date: str, end_date: str, keywords: list):
        keywords_list = self.split_keywords(keywords)
        time_range_list = self.get_time_range_list(start_date, end_date)
        for start_date, end_date in time_range_list:
            for keywords in keywords_list:
                params = {
                    'keywords': keywords,
                    'start_date': start_date,
                    'end_date': end_date
                }
                self.params_queue.put(params)

    def get_all_country(self):
        self.init_queue("2006-06-01", "2010-12-31", self.keywords)
        while True:
            if self.params_queue.empty():
                break
            try:
                params_data = self.params_queue.get(timeout=1)
                self.area = 0
                encrpt_datas, uniqid = self.get_encrypt_datas(
                    start_date=params_data['start_date'],
                    end_date=params_data['end_date'],
                    keywords=params_data['keywords']
                )
                key = self.get_key(uniqid)
                for encrypt_data in encrpt_datas:
                    encrypt_data['pc']['data'] = self.decrypt_func(key, encrypt_data['pc']['data'])
                    for formated_data in self.format_data(encrypt_data):
                        yield formated_data
                time.sleep(random.uniform(5, 9))
            except requests.Timeout:
                self.params_queue.put(params_data)

    def get_encrypt_datas(self, start_date, end_date, keywords):
        params = {
            'word': ','.join(keywords),
            'startDate': start_date.strftime('%Y-%m-%d'),
            'endDate': end_date.strftime('%Y-%m-%d'),
            'area': self.area
        }
        url = 'http://index.baidu.com/api/SearchApi/index?' + urlencode(params)
        html = self.get(url)
        datas = json.loads(html)
        uniqid = datas['data']['uniqid']
        encrypt_datas = []
        for single_data in datas['data']['userIndexes']:
            encrypt_datas.append(single_data)
        return (encrypt_datas, uniqid)

    def get(self, url):
        headers['Cookie'] = self.cookie
        res = requests.get(url, headers=headers, timeout=30)
        if not res.status_code == 200:
            raise requests.Timeout
        return res.text

    def format_data(self, data):
        keyword = str(data['word'])
        start_date = datetime.datetime.strptime(data['pc']['startDate'], '%Y-%m-%d')
        end_date = datetime.datetime.strptime(data['pc']['endDate'], '%Y-%m-%d')
        date_list = list()
        while start_date <= end_date:
            date_list.append(start_date)
            start_date += datetime.timedelta(days=1)
        endYear = int(str(end_date.year))
        if endYear < 2011:
            all_kind = ['pc']
        else:
            all_kind = ['all', 'pc', 'wise']
        for kind in all_kind:
            index_datas = data[kind]['data']
            for i, cur_date in enumerate(date_list):
                try:
                    index_data = index_datas[i]
                except IndexError:
                    index_data = ''
                formated_data = {
                    'keyword': keyword,
                    'type': kind,
                    'date': cur_date.strftime('%Y-%m-%d'),
                    'index': index_data if index_data else '0',
                    'area': self.area
                }
                yield formated_data

    def get_key(self, uniqid):
        url = 'http://index.baidu.com/Interface/api/ptbk?uniqid=%s' % uniqid
        html = self.get(url)
        datas = json.loads(html)
        key = datas['data']
        return key

    def decrypt_func(self, key, data):
        a = key
        i = data
        n = {}
        s = []
        for o in range(len(a) // 2):
            n[a[o]] = a[len(a) // 2 + o]
        for r in range(len(data)):
            s.append(n[i[r]])
        return ''.join(s).split(',')

    def get_time_range_list(self, start_date: str, end_date: str) -> list:
        date_range_list = list()
        startdate = datetime.datetime.strptime(start_date, '%Y-%m-%d')
        enddate = datetime.datetime.strptime(end_date, '%Y-%m-%d')
        while True:
            tempdate = startdate + datetime.timedelta(days=300)
            if tempdate > enddate:
                date_range_list.append((startdate, enddate))
                break
            date_range_list.append((startdate, tempdate))
            startdate = tempdate + datetime.timedelta(days=1)
        return date_range_list

    def split_keywords(self, keywords: list) -> [list]:
        return [keywords[i * 5: (i + 1) * 5] for i in range(math.ceil(len(keywords) / 5))]
