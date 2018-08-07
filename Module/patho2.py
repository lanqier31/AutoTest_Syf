#!/usr/bin/env python
# -*-coding:utf-8 -*-
# Author:  Weir Gao --<>
# Purpose:
# Created: 2018/7/24

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
import time
from time import sleep
import unittest
import sys
import os
import openpyxl
from openpyxl.reader.excel import load_workbook
import Config
from Commons import Login, SyfClinicalReport,ReportList,operateExcel,conn_sqlserver

#we use global varialbe driver
driver=Config.ChromeDriver
#Url for login is a global varible
url=Config.LoginUrl

Login.login_Syf(url,'30048','5913')
Login.maxmize_window()
SyfClinicalReport.goto_Report()

reports = operateExcel.get_column('ReportNo')
n = 2  # excel row
for re in reports:   #遍历要测试的病历号
    #不同病历号，直接替换reportNo 不适用
    SyfClinicalReport.input_Hid('2796182')
    WebDriverWait(driver, 10).until_not(
        lambda the_driver: the_driver.find_element_by_class_name('divBlockHid').is_displayed())
    driver.find_element_by_id('txtCheckReportType').click()
    driver.find_element_by_css_selector("div[data-text='病理形态学']").click()  # 选择病理形态学的报告类别
    WebDriverWait(driver, 10).until(
        lambda the_driver: the_driver.find_element_by_id('tbodyReportList').is_displayed())
    tbodyReportList = driver.find_element_by_id('tbodyReportList')
    tbodyreport = tbodyReportList.find_elements_by_tag_name('tr')[0]
    driver.execute_script('$("#divReportList tr")[0].setAttribute("report-no", "' + re + '" )')
    sleep(3)
    tbodyreport = tbodyReportList.find_elements_by_tag_name('tr')[0]
    tbodyreport.click()
    sleep(1)
    SyfClinicalReport.alert_close()

