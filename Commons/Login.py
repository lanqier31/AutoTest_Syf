# -*- coding: GBK -*-
# Author：WeirGao

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
import Config
import sqlite3
import pyodbc
import time
from openpyxl.reader.excel import load_workbook
import sys
import os
#we use global varialbe driver
driver=Config.ChromeDriver

user_name_id='txtUserName'
user_password_id='txtPassword'
loginbtn_id='btnConfirm'
login_alert_id='spanMessage'
all_link_id='txtAll'

def maxmize_window():
    driver.maximize_window()

#点击登陆图标
def click_loginbtn():
    loginbtn=driver.find_element_by_id(loginbtn_id)
    ActionChains(driver).move_to_element(loginbtn).perform()
    loginbtn.click()

#重置用户名
def ResetLoginName(name):
    WebDriverWait(driver, 10).until(lambda the_driver: the_driver.find_element_by_id(loginbtn_id).is_displayed())
    driver.find_element_by_id(user_name_id).clear()
    driver.find_element_by_id(user_name_id).send_keys(name)

#重置密码
def ResetPassword(pwd):
    WebDriverWait(driver, 10).until(lambda the_driver: the_driver.find_element_by_id(loginbtn_id).is_displayed())
    driver.find_element_by_id(user_password_id).clear()
    driver.find_element_by_id(user_password_id).send_keys(pwd)

# 登陆Syf系统
def login_Syf(url, name, passwd):
    driver.get(url)
    WebDriverWait(driver, 10).until(lambda the_driver: the_driver.find_element_by_id(loginbtn_id).is_displayed())
    ResetLoginName(name)
    ResetPassword(passwd)
    click_loginbtn()
    time.sleep(5)
    # WebDriverWait(driver, 10).until(lambda the_driver: the_driver.find_element_by_id(all_link_id).is_displayed())
    if 'Login/SelectSystem' in driver.current_url:
        print 'Login Success, test login is pass'
    else:
        print 'expected string is not in the url after login'