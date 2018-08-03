# -*- coding: utf-8 -*-
# Author：WeirGao

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
import Config
import time
import operateExcel
from time import sleep

#报告列表页面
url ='http://'+Config.IP+Config.Version+u'/SyfHospitalClinicalDataCenter/index'

driver=Config.ChromeDriver

def goto_reportList():
    driver.get(url)

def del_checkCode(Hid):
    """"delete  the checkCode of patient"""
    try:
        WebDriverWait(driver, 30, 0.5).until(
        lambda the_driver: the_driver.find_element_by_id('txtAll').is_displayed())

        driver.find_element_by_id('txtQuery').send_keys(Hid)
        driver.find_element_by_id('btnSearch').click()
        WebDriverWait(driver, 20).until(lambda the_driver: the_driver.find_element_by_id('dataTalbeContainer').is_displayed())
        sleep(3)
        is_disappeared = WebDriverWait(driver, 20, 1).until_not(
            lambda x: x.find_element_by_xpath('//div[@class="md-LoadData"]').is_displayed())   #/html/body/div[3]
        if not is_disappeared:
            print Hid+"手术列表信息加载超时"
        driver.find_element_by_name('checkAll').click()    #choose checkAll
        driver.find_element_by_id('btnDelete').click()
        sleep(2)
        driver.switch_to_alert().accept() #switch_to_alert
    except Exception as e:
        print('发生了异常：', e)