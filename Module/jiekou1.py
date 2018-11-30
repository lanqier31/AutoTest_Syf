#!/usr/bin/env python
# -*-coding:utf-8 -*-

import requests,json
from Commons import globals
from Commons import  operateExcel
from openpyxl.reader.excel import load_workbook
import Config

url = 'http://192.168.10.243/Syf1.2.0/OriginalReportB/GetPatReport'
params ={}
params['HospitalNumber'] ='4146732'
headers ={}
headers['Accept']='application/json, text/javascript, */*; q=0.01'
headers['Origin']='http://192.168.10.243'
headers['X-Requested-With']='XMLHttpRequest'
headers['User-Agent']='Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.77 Safari/537.36'
headers['Content-Type']='application/x-www-form-urlencoded; charset=UTF-8'
headers['Accept-Encoding']='gzip, deflate'
headers['Accept-Language']='zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7'
headers['Cookie']='__RequestVerificationToken_L1N5ZjEuMi4w0=7EQoupSreI4QFeLw3WSjfAAdiEcsouJVDUVxd_AVOgORpGlwFfC7V6DBXXbVaLL5qbVjWuXM1LVomz7lxtxFdHZNty8KGc90sX-707YBdYY1; ASP.NET_SessionId=3bk4dpuxmegbxne1bvomzb03'

autocase = Config.autocase_path
book = load_workbook(autocase)
n=2
# try:
sheet=book['111']
r = requests.post(url,data=params,headers=headers)
#转换为python类型的字典格式,json包的响应结果，调用json(),转换成python类型
json_r = r.json()
for key in json_r:
    print key,json_r[key]
    sheet['A'+str(n)]=key
    sheet['B'+str(n)]=str(json_r[key])
    n = n+1
book.save(autocase)
