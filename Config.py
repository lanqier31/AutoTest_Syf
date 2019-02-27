# -*- coding: utf-8 -*-
# Author：WeirGao

import os

from selenium import webdriver
import sys
reload(sys)   #必须要reload
sys.setdefaultencoding('utf-8')

driverOptions = webdriver.ChromeOptions()
# driverOptions.add_argument('--headless')   #浏览器不提供可视化页面
driverOptions.add_argument("--start-maximized")  #默认屏幕最大化
# driverOptions.add_argument("--no-sandbox")
# driverOptions.add_argument("--disable-dev-shm-usage")
driverOptions.add_argument('--disable-gpu') #谷歌文档提到需要加上这个属性来规避bug
driverOptions.add_argument('disable-infobars') # 规避浏览器显示 Chrome正在受到自动软件的控制
driverOptions.add_argument(r"user-data-dir=C:\Users\Administrator\AppData\Local\Google\Chrome\User Data")  # 读取用户设置内容

ChromeDriver = webdriver.Chrome("chromedriver",0,driverOptions)
IP = '192.168.10.243/'
Version = 'syf1.2.0'
LoginUrl= 'http://'+IP+Version+'/login/index'
baseUrl = 'http://192.168.10.243/'

basedir = os.path.abspath(os.path.dirname(__file__))


log_file_path = os.path.join(basedir, 'Log')

screens_file_path=os.path.join(basedir, 'PageScreen')
autocase_path = os.path.join(basedir, 'AutoCase/AutoTestCases.xlsx')
interfaceTest_path = os.path.join(basedir, 'AutoCase/interfaceTest.xlsx')
jiekou_path = os.path.join(basedir, 'AutoCase/body.txt')
redkey = os.path.join(basedir, 'AutoCase/ddd.xlsx')

#sqlserver
# conn = pymssql.connect('192.168.10.164', 'sa', 'sa', '20180806')

reportType={
    "Bchao_pre":"超声声像",
    "Bchao_fellowup":"超声声像",
    "pathology" : "常规病理",
    "pathology_bingdong":"冰冻标本",
    "ImgA":"影像A",
    "ImgB":"影像B",
    "Img131I":"核素影像",
    "AsssyA":"检验A",
    "BFna_cell":"细胞病理",
    "B_FNA":"B-FNA操作",
}