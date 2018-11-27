#!/usr/bin/env python
# -*-coding:utf-8 -*-
# Author:  Weir Gao --<>
# Purpose:
# Created: 2018/7/24

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
import time
import unittest
import sys
import os
import openpyxl
from openpyxl.reader.excel import load_workbook
import HTMLTestRunner
import Config
from Commons import Login, SyfClinicalReport


#we use global varialbe driver
driver=Config.ChromeDriver
#Url for login is a global varible
url=Config.LoginUrl


class reportShow(unittest.TestCase):
    def test_01reportshow(self):
        Login.login_Syf(url,'30048','5913')
        Login.maxmize_window()
        SyfClinicalReport.goto_Report()
        SyfClinicalReport.input_Hid('3444471')

    def test_02surgeryNum(self):
        num_surgery = SyfClinicalReport.num_surgery()
        print num_surgery

    def test_03surgety_pathology(self):
        SyfClinicalReport.surgery_pathology()


if __name__=='__main__':
     unittest.main()