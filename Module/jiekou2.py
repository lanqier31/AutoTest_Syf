#!/usr/bin/env python
# -*-coding:utf-8 -*-

import requests,json
from Commons import globals
from Commons import  operateExcel
from openpyxl.reader.excel import load_workbook
from Commons.Confighttp import getPatReport,headers
import Config
jiekou_path = Config.jiekou_path
url = 'http://192.168.10.243/Syf1.2.0/OriginalReportB/SaveReport'
with open(jiekou_path, 'rb') as fp:
    a = fp.read()
    a = json.loads(a)

jsonData = {}
jsonData['checkHospitalName']=''
jsonData['mrID']=a['hospitalNumber']
jsonData['ReportFilePath'] = getPatReport(a['hospitalNumber'])['ReportFilePath']
jsonData['reportName'] = getPatReport(a['hospitalNumber'])['reportName']
jsonData['patientName'] = a['patientName']
jsonData['checkDate'] = a['checkDate']
jsonData['checkDoctorName'] = '管理员'
jsonData['checkItem'] = getPatReport(a['hospitalNumber'])['reportName']
jsonData['checkResult'] =a['checkResult']
jsonData['checkConclusion'] =a['checkConclusion']
jsonData['ReportNo'] =getPatReport(a['hospitalNumber'])['ReportNo']

'''
jsonData={"reportName": "NGS", "mrID": "4146732", "checkDate": "2018-07-24", "checkHospitalName": "", "checkDoctorName": "\u7ba1\u7406\u5458", "checkConclusion": "%E6%9C%AC%E6%AC%A1%E6%A3%80%E6%B5%8B%E6%9C%AA%E5%8F%91%E7%8E%B0%E5%B7%B2%E7%9F%A5%E7%AA%81%E5%8F%98%EF%BC%8C%E8%AF%A6%E8%A7%81%E6%8A%A5%E5%91%8A%E5%8D%95%E3%80%82", "checkItem": "NGS", "checkResult": "%E7%9A%84%E5%8F%91%E9%80%81%E5%88%B0%E5%8F%91%E9%80%81%E5%88%B0", "ReportFilePath": "4146732_NG201800147.webp", "ReportNo": "4146732_NG201800147", "patientName": "\u738b\u654f\u73e0"}
'''


params ={}
params['HospitalNumber'] =a['hospitalNumber']
params['jsonData']=json.dumps(jsonData)
headers =headers

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


