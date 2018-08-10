# -*- coding: GBK -*-
# Author：WeirGao

import os
from selenium import webdriver
import sys

ChromeDriver=webdriver.Chrome()
IP = '192.168.10.110/'
Version = 'syf1.2.0'
LoginUrl='http://'+IP+Version+'/login/index'


basedir = os.path.abspath(os.path.dirname(__file__))


loginfo_file_path = os.path.join(basedir, 'Log')

screens_file_path=os.path.join(basedir, 'PageScreen')

autocase_path = os.path.join(basedir,'AutoCase/AutoTestCases.xlsx')