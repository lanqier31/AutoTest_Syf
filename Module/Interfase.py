#!/usr/bin/env python
# -*-coding:utf-8 -*-

import requests,json
from Commons import globals

url = "http://192.168.10.243/Syf1.2.0/SyfHospitalClinicalDataCenter/QueryPatientReportListByKey"
params = {"HospitalNumber":'3444471',"QueryKey":"A","StartDateTime":"2015-05-21 23:59:59","EndDateTime":"1994-04-30 00:00:00","CheckReportType":"超声声像","NianQiShiDuanDates":'[]'}
headers = {}
headers['Accept']='application/json, text/javascript, */*; q=0.01'
headers['Origin']='http://192.168.10.243'
headers['X-Requested-With']='XMLHttpRequest'
headers['User-Agent']='Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.77 Safari/537.36'
headers['Content-Type']='application/x-www-form-urlencoded; charset=UTF-8'
headers['Accept-Encoding']='gzip, deflate'
headers['Accept-Language']='zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7'
headers['Cookie']='ASP.NET_SessionId=nmov53zbgqhcfxff3vwh5rtd'
try:
    r = requests.post(url,params=params,headers=headers)
    #转换为python类型的字典格式,json包的响应结果，调用json(),转换成python类型
    json_r = r.json()
    for key in json_r:
        print key
    print json_r['objData']
except BaseException as e:
    print("请求不能完成:",str(e))


