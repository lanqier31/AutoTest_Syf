import requests
from requests.cookies import RequestsCookieJar
from urllib import unquote
import xlwt
import chardet
import json,re
from Config import  jiekou_path
import sys
reload(sys)
sys.setdefaultencoding("utf8")

#Wangjian 2018/11/30 17:15:43
url1 = "http://192.168.10.110/syf1.2.0/Login/Index"
url2 = "http://192.168.10.110/syf1.2.0/OriginalReportB/ReportProcessing"
url3 = "http://192.168.10.110/syf1.2.0/Patient/CheckPatientExists"
url4 = "http://192.168.10.110/syf1.2.0/OriginalReportB/GetPatReport"
url5 = "http://192.168.10.110/syf1.2.0/OriginalReportB/ImageToText?url=http://192.168.10.110//syf1.2.0/ReportImages/4146732/4146732-20181009091820555827.jpg"
url6 = "http://192.168.10.110/syf1.2.0/OriginalReportB/SaveReport"
url7 = "http://192.168.10.110/syf1.2.0/login/validloginstatus"
par = {"type":"add",
       "editmark":"true"}
headers = {"Accept": "application/json, text/javascript, */*; q=0.01",
           "Accept-Encoding": "gzip, deflate",
           "Accept-Language": "zh-CN,zh;q=0.9",
           "User-Agent": "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.36",
           "X-Requested-With": "XMLHttpRequest",
           "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
           "__RequestVerificationToken": "_n5OtcsKbR86qCwUX6T-xrlNcswqrWycpvDn-Ra8eswUlopt5tDVvPBctxYMmyEIFH64KZZnEI0NIfC-eRZN_y8omr72T2KmYMlriei9CL41",
           "Referer": "http://192.168.10.110/syf1.2.0/login/index",
           "Origin": "http://192.168.10.110",
           "Connection":"keep-alive",
           "hsot":"192.168.10.110"}
body = {"userName":"30048",
        "password":"8613",
        "remember":True}
json1 = {"hospitalNumber":"4146732"}
json2 = {"reportFilePath":"4146732-20181009091820555827.jpg",
         "hospitalNumber":"4146732"}
with open(jiekou_path, 'rb') as fp:
    a = fp.read()
    json3 = eval(a)
print json3

s = requests.session()
s.headers = headers
c = requests.cookies.RequestsCookieJar()
c.set("ASP.NET_SessionId","2x0qswwp0tp0ddmpoiqybx3x")
c.set("__RequestVerificationToken_L3N5ZjEuMi4w0","Vz3OhrJALyuN-jcCxsc71jORKDYlKH5mw0_U3eAZKTdqLYwciFCX4yDd_SgPMJEj15-flnGBS_9_WyIcuCfhWd-MOSPNKwrG_OUhu1qy3Pw1")
s.cookies.update(c)
s.post(url=url1,data=body,verify=False)
# s.headers["Referer"] = "http://192.168.10.110/syf1.2.0/SyfHospitalClinicalDataCenter/index"
s.get(url=url2,params=par,verify=False)
s.headers["Referer"]= "http://192.168.10.110/syf1.2.0/OriginalReportB/ReportProcessing?type=add&editmark=true"
s.post(url=url3,data=json1,verify=False)
s.post(url=url4,data=json1,verify=False)
s.post(url=url5,data=json2,verify=False)
s.headers["Accept"]="application/json, text/javascript, */*; q=0.01"
s.headers["User-Agent"]="Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.36"
s.headers["X-Requested-With"]="XMLHttpRequest"
s.headers["Content-Type"]="application/x-www-form-urlencoded; charset=UTF-8"
s.headers["Accept-Language"]="zh-CN,zh;q=0.9"
r = s.post(url=url6,data=json3,verify=False)
json_r = r.json()
print json_r
