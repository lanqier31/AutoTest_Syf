#!/usr/bin/env python
# -*-coding:utf-8 -*-
# Author:  Weir Gao --<>
# Purpose:影像B的校验代码化测试
# Created: 2018/7/24


import sys
import os
import Config
from Commons import Login, SyfClinicalReport,ReportList,operateExcel


#we use global varialbe driver
driver=Config.ChromeDriver
#Url for login is a global varible
url=Config.LoginUrl

Login.login_Syf(url,'30048','5913')
Login.maxmize_window()
HospitalNums = operateExcel.All_content('Hid')
for Hid in HospitalNums:   #遍历要测试的病历号
    SyfClinicalReport.goto_Report()
    SyfClinicalReport.input_Hid(Hid)
    SyfClinicalReport.selectType('ImgB')
    SyfClinicalReport.yearSelect('a','2')
    SyfClinicalReport.jiaoyan_ImgB(Hid)
    SyfClinicalReport.yearSelect('a','5')
    SyfClinicalReport.jiaoyan_ImgB(Hid)
    SyfClinicalReport.yearSelect('a','10')
    SyfClinicalReport.jiaoyan_ImgB(Hid)