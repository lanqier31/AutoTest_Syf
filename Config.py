# -*- coding: utf-8 -*-
# Author：WeirGao

import os
import pymssql
from selenium import webdriver
import sys
reload(sys)   #必须要reload
sys.setdefaultencoding('utf-8')

ChromeDriver=webdriver.Chrome()
IP = '192.168.10.243/'
Version = 'syf1.2.0'
LoginUrl='http://'+IP+Version+'/login/index'


basedir = os.path.abspath(os.path.dirname(__file__))


log_file_path = os.path.join(basedir, 'Log')

screens_file_path=os.path.join(basedir, 'PageScreen')

autocase_path = os.path.join(basedir,'AutoCase/AutoTestCases.xlsx')
redkey = os.path.join(basedir,'AutoCase/ddd.xlsx')

#sqlserver
# conn = pymssql.connect('192.168.10.164', 'sa', 'sa', '20180806')