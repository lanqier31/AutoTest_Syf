#!/usr/bin/env python
# -*-coding:utf-8 -*-

import requests,json,Config
from Commons import globals


headers ={}
headers['Accept']='application/json, text/javascript, */*; q=0.01'
headers['Origin']='http://192.168.10.243'
headers['X-Requested-With']='XMLHttpRequest'
headers['User-Agent']='Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.77 Safari/537.36'
headers['Content-Type']='application/x-www-form-urlencoded; charset=UTF-8'
headers['Accept-Encoding']='gzip, deflate'
headers['Accept-Language']='zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7'
headers['Cookie']='__RequestVerificationToken_L1N5ZjEuMi4w0=7EQoupSreI4QFeLw3WSjfAAdiEcsouJVDUVxd_AVOgORpGlwFfC7V6DBXXbVaLL5qbVjWuXM1LVomz7lxtxFdHZNty8KGc90sX-707YBdYY1; ASP.NET_SessionId=3bk4dpuxmegbxne1bvomzb03'

class ConfigHttp:
    def __init__(self):
        global host, port, timeout
        host = Config.baseUrl
        self.log = globals.log()
        self.headers = {}
        self.params = {}
        self.data = {}
        self.url = None
        self.files = {}

    def set_url(self, url):
        self.url = host + url

    def set_headers(self, header):
        self.headers = header

    def set_params(self, param):
        self.params = param

    def set_data(self, data):
        self.data = data

    def set_files(self, file):
        self.files = file

    # defined http get method
    def get(self):
        try:
            response = requests.get(self.url, params=self.params, headers=self.headers)
            # response.raise_for_status()
            return response
        except Exception as e:
            self.log(e.message)
            return None

    # defined http post method
    def post(self):
        try:
            response = requests.post(self.url, headers=self.headers, data=self.data)
            # response.raise_for_status()
            return response
        except Exception as e:
            self.log(e.message)
            return None


def getPatReport(hid):
    url = 'http://192.168.10.243/Syf1.2.0/OriginalReportB/GetPatReport'
    params = {}
    params['HospitalNumber'] = hid
    r = requests.post(url, data=params, headers=headers)
    ObjData = r.json()['ObjData']
    ReportFilePath = ObjData[0]
    ReportNo = ReportFilePath.replace('.webp', '')
    reportName = 'NGS'
    if ObjData:
        for i in range(0, len(ObjData)):
            if "NG" in ObjData[i]:
                ReportFilePath = ReportFilePath
                ReportNo = ReportFilePath.replace('.webp', '')
                reportName = 'NGS'
                continue
            elif "BR" in ObjData[i]:
                ReportFilePath = ReportFilePath
                ReportNo = ReportFilePath.replace('.webp', '')
                reportName = 'BRAF'
                continue
    return {"ReportFilePath": ReportFilePath, "reportName": reportName, "ReportNo": ReportNo}