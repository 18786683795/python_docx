# -*- coding: utf-8 -*-
"""
Created on Fri May 22 15:21:58 2020

@author: Lenovo
"""

from django.shortcuts import render
from docxtpl import DocxTemplate,InlineImage
from docx.shared import Mm, Inches, Pt
 

base_url = 'E:/GZ/Django/Django_API-1/Django_baogao/'
asset_url = base_url + 'test模板.docx'
tpl = DocxTemplate(asset_url)
context = {'text': '哈哈哈，来啦',
           't1':'燕子',
            't2':'杨柳',
            't3':'桃花',
            't4':'针尖',
            't5':'头涔涔',
            't6':'泪潸潸',
            't7':'茫茫然',
            't8':'伶伶俐俐',
            'picture1': InlineImage(tpl, '1.jpg', width=Mm(80), height=Mm(60)),}

user_labels = ['姓名', '年龄', '性别', '入学日期']
context['user_labels'] = user_labels
user_dict1 = {'number': 1, 'cols': ['林小熊', '27', '男', '2019-03-28']}
user_dict2 = {'number': 2, 'cols': ['林小花', '27', '女', '2019-03-28']}
user_list = []
user_list.append(user_dict1)
user_list.append(user_dict2)

context['user_list'] = user_list
tpl.render(context)
tpl.save(base_url + 'test.docx')



