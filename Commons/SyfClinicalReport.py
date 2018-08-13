# -*- coding: utf-8 -*-
# Author：WeirGao

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.support import expected_conditions as EC
from openpyxl.reader.excel import load_workbook
import Config
import time
import operateExcel
from time import sleep
#进入初诊页面

driver=Config.ChromeDriver
autocase = Config.autocase_path
book = load_workbook(autocase)

#病历号输入框
txtHid ="$('#txtHospitalNumber')"

#截图地址
screen = Config.screens_file_path
#Bchao 截图地址
bscreen = Config.screens_file_path+'/Bchao/'
#影像B截图地址
imgBscreen = Config.screens_file_path+'/ImgB/'
#细胞病理学截图地址
cytoscreen = Config.screens_file_path+'/Cytology/'
#常规病理截图地址
pathoscreen = Config.screens_file_path+'/pathology/'

def goto_Report():

    url = 'http://'+Config.IP+Config.Version+'/SyfHospitalClinicalDataCenter/QueryReport?type=add&editmark=true'
    driver.get(url)

def input_Hid(hid):
    '''输入病历号'''
    # exec_js = "$('#txtHospitalNumber').val(hid)"
    driver.execute_script(txtHid+'.val("' + hid + '")')
    driver.execute_script(txtHid+".blur()")
    sleep(5)
    # WebDriverWait(driver, 10).until(
    #     lambda the_driver: the_driver.find_element_by_id('selShouShuList').is_displayed())
    alert_close()

def num_surgery():
    '''检查手术次数'''
    surgeryList =Select(driver.find_element_by_id('selShouShuList'))
    options_list = surgeryList.options
    return len(options_list)


def alert_close():
    """判断是否存在提醒框并关闭"""
    alert = EC.alert_is_present()(driver)
    if alert:
        if (u'外院手术' in alert.text):
            alert.dismiss()
        elif (u'没有此病人信息'in alert.text):
            alert.accept()
            print "不存在该病历号"
        else:
            alert.accept()




#判断元素是否可见
def is_element_visible(element):
    try:
        the_element = EC.visibility_of_element_located(element)
        assert the_element(driver)
        flag = True
    except:
        flag = False
    return flag


#截图
def screenpatho():
    nowTime = time.strftime("%Y%m%d.%H.%M.%S")
    driver.get_screenshot_as_file(r'../PageScreen/pathology/%s.jpg' % nowTime)


def screenImgB():
    nowTime = time.strftime("%Y%m%d.%H.%M.%S")
    driver.get_screenshot_as_file(r'../PageScreen/ImgB/%s.jpg' % nowTime)

def screenBchao():
    nowTime = time.strftime("%Y%m%d.%H.%M.%S")
    driver.get_screenshot_as_file(r'../PageScreen/Bchao/%s.jpg' % nowTime)

def screenPatho(Hid):
    driver.get_screenshot_as_file(pathoscreen + Hid + '.png')

def undo():
    driver.find_element_by_id('btnRes').click()
    sleep(1)
    alert_close()
    sleep(1)


def selectImgB():
    wait_loading()
    driver.find_element_by_id('txtCheckReportType').click()
    sleep(1)
    driver.find_element_by_css_selector("div[data-text='影像B']").click()  # 选择病理形态学的报告类别
    WebDriverWait(driver, 10).until(
        lambda the_driver: the_driver.find_element_by_id('tbodyReportList').is_displayed())
    sleep(1)


def yearSelect(value):
    sleep(2)
    WebDriverWait(driver, 10).until(
        lambda the_driver: the_driver.find_element_by_id('divOperationList').is_displayed())
    # 读取首次手术下的随访
    table = driver.find_element_by_xpath('//div[@id="divOperationList"]/table[last()]')
    sleep(2)
    table.find_element_by_xpath('tbody/tr[4]/td[1]').click()
    sleep(1)
    table.find_element_by_name('selNianQi').click()
    year = Select(table.find_element_by_name('selNianQi'))
    year.select_by_value(value)
    sleep(1)
    alert_close()


def jiaoyan_Bchao(Hid):
    sheet = book['Bchao']
    try:
        WebDriverWait(driver, 10).until(
            lambda the_driver: the_driver.find_element_by_id('tbodyReportList').is_displayed())
    except Exception as e:
        print e
    alert_close()
    tbodyReportList = driver.find_element_by_id('tbodyReportList')
    reportList = tbodyReportList.find_elements_by_tag_name('tr')
    if(len(reportList)== 1):
        if (reportList[0].text=='没有报告信息.'):
            return Hid + u'无B超报告'
    n = operateExcel.max_row('Bchao') + 1
    for j in range(len(reportList)):  # 遍历B超报告份数
        reportList = tbodyReportList.find_elements_by_tag_name('tr')
        reportList[j].click()
        sleep(1)
        # 判断是否有alert
        alert = EC.alert_is_present()(driver)
        if alert:
            if (u'外院手术' in alert.text):
                alert.dismiss()
            else:
                alert.accept()
        className = reportList[j].find_element_by_xpath('td[3]/div').get_attribute("class")
        while ('StateNoCom' != className and 'StateCFCom' != className):
            WebDriverWait(driver, 20).until_not(
                lambda the_driver: the_driver.find_element_by_class_name('divBlockHid').is_displayed())
            undo()
            alert_close()
            reportList = tbodyReportList.find_elements_by_tag_name('tr')
            reportList[j].click()  # 焦点重新回到该报告上
            className = reportList[j].find_element_by_xpath('td[3]/div').get_attribute("class")
        WebDriverWait(driver, 20).until_not(
            lambda the_driver: the_driver.find_element_by_xpath('//div[@class="divBlockHid"]').is_displayed())
        checkResult = driver.find_element_by_xpath('//div[@id="divUltrasonography"]/div[1]/div[5]/div[2]').text
        checkConclusion = driver.find_element_by_xpath(
            '//div[@id="divUltrasonography"]/div[1]/div[6]/div[2]').text
        WebDriverWait(driver, 10).until(
            lambda the_driver: the_driver.find_element_by_id('btnTest').is_displayed())
        # driver.find_element_by_id('btnTest').click()  # 点击校验按钮
        suoj_XY = driver.find_element_by_xpath('//div[@id="divUltrasonography"]/div[3]/div[5]/div[2]').text  # 腺叶所见（随访）
        zhend_XY = driver.find_element_by_xpath('//div[@id="divUltrasonography"]/div[3]/div[6]/div[2]').text  # 腺叶诊断

        suoj_XC = driver.find_element_by_xpath('//div[@id="divUltrasonography"]/div[3]/div[7]/div[2]').text  # 腺床所见
        zhend_XC = driver.find_element_by_xpath('//*[@id="divUltrasonography"]/div[3]/div[8]/div[2]').text  # 腺床诊断

        suoj_QC = driver.find_element_by_xpath('//*[@id="divUltrasonography"]/div[3]/div[9]/div[2]').text  # 清扫床所见
        zhend_QC = driver.find_element_by_xpath('//*[@id="divUltrasonography"]/div[3]/div[10]/div[2]').text  # 清扫床诊断

        suoj_PL = driver.find_element_by_xpath('//*[@id="divUltrasonography"]/div[3]/div[11]/div[2]').text  # 毗邻区所见
        zhend_PL= driver.find_element_by_xpath('//*[@id="divUltrasonography"]/div[3]/div[12]/div[2]').text  # 毗邻区诊断

        sheet['A'+str(n)] = Hid
        sheet['B'+str(n)] = checkResult
        sheet['C'+str(n)] = checkConclusion
        sheet['D'+str(n)] = suoj_XY
        sheet['E'+str(n)] = zhend_XY
        sheet['F'+str(n)] = suoj_XC
        sheet['G'+str(n)] = zhend_XC
        sheet['H'+str(n)] = suoj_QC
        sheet['I'+str(n)] = zhend_QC
        sheet['J'+str(n)] = suoj_PL
        sheet['K'+str(n)] = zhend_PL

        book.save(autocase)
        n = n+1
        driver.get_screenshot_as_file(bscreen+Hid+'.png')
        # driver.find_element_by_id('btnQuery').click()
        WebDriverWait(driver, 10).until(
            lambda the_driver: the_driver.find_element_by_id('tbodyReportList').is_displayed())



def jiaoyan_ImgB(Hid):
    """影像B校验代码化"""
    sheet = book['ImgB']
    try:
        WebDriverWait(driver, 10).until(
            lambda the_driver: the_driver.find_element_by_id('tbodyReportList').is_displayed())
    except Exception as e:
        print e
    alert_close()

    tbodyReportList = driver.find_element_by_id('tbodyReportList')
    reportList = tbodyReportList.find_elements_by_tag_name('tr')
    if(len(reportList)== 1):
        if (reportList[0].text=='没有报告信息.'):
            return Hid + u'无影像学B报告'
    n = operateExcel.max_row('ImgB') + 1
    for j in range(len(reportList)):  # 遍历影像B报告份数
        reportList = tbodyReportList.find_elements_by_tag_name('tr')
        reportList[j].click()
        sleep(1)
        # 判断是否有alert
        alert = EC.alert_is_present()(driver)
        if alert:
            if (u'外院手术' in alert.text):
                alert.dismiss()
            else:
                alert.accept()
                continue
        className = reportList[j].find_element_by_xpath('td[3]/div').get_attribute("class")
        while('StateNoCom'!= className and 'StateCFCom'!= className):
            WebDriverWait(driver, 20).until_not(
                lambda the_driver: the_driver.find_element_by_class_name('divBlockHid').is_displayed())
            undo()
            alert_close()
            reportList = tbodyReportList.find_elements_by_tag_name('tr')
            reportList[j].click()  #焦点重新回到该报告上
            className = reportList[j].find_element_by_xpath('td[3]/div').get_attribute("class")
        checkResult = driver.find_element_by_xpath('//div[@id="divImagingExamination"]/div[1]/div[5]/div[2]').text
        checkConclusion = driver.find_element_by_xpath(
            '//div[@id="divImagingExamination"]/div[1]/div[6]/div[2]').text
        WebDriverWait(driver, 20).until_not(
            lambda the_driver: the_driver.find_element_by_class_name('divBlockHid').is_displayed())
        driver.find_element_by_id('btnCode').click()  # 点击代码化
        wait_loading()
        FindingA = driver.find_element_by_xpath('//div[@name = "MajorAnomalies"]').text
        AbnormalClassify = driver.find_element_by_xpath('//input[@name="AbnormalClassify"]').text
        sheet['A' + str(n)] = Hid
        sheet['B' + str(n)] = checkResult
        sheet['C' + str(n)] = checkConclusion
        sheet['D' + str(n)] = FindingA
        sheet['E' + str(n)] = AbnormalClassify
        book.save(autocase)
        n = n+1
        screenImgB()
        # nowTime = time.strftime("%Y%m%d.%H.%M.%S")
        # driver.get_screenshot_as_file(r'../PageScreen/ImgB/%s.jpg' % nowTime)


def wait_loading():
    """waiting for divBlockHid disappear"""
    result = WebDriverWait(driver, 20).until_not(
        lambda the_driver: the_driver.find_element_by_class_name('divBlockHid').is_displayed())
    if result:
        return True
    else:
        return False
