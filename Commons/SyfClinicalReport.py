# -*- coding: utf-8 -*-
# Author：WeirGao

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.support import expected_conditions as EC
import Config
import time
import operateExcel
from time import sleep
#进入初诊页面

driver=Config.ChromeDriver

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
def screenAsTime():
    nowTime = time.strftime("%Y%m%d.%H.%M.%S")
    driver.get_screenshot_as_file(r'../PageScreen/%s.jpg' % nowTime)


def screenPatho(Hid):
    driver.get_screenshot_as_file(pathoscreen + Hid + '.png')

def undo():
    driver.find_element_by_id('btnRes').click()
    sleep(1)
    alert_close()
    sleep(1)


def selectImgB():
    driver.find_element_by_id('txtCheckReportType').click()
    driver.find_element_by_css_selector("div[data-text='影像B']").click()  # 选择病理形态学的报告类别
    WebDriverWait(driver, 10).until(
        lambda the_driver: the_driver.find_element_by_id('tbodyReportList').is_displayed())
    sleep(1)

def surgery_pathology(Hid):
    """遍历某病历号下的常规病理，并显示所见和诊断的内容，校验后的所见诊断的内容"""
    WebDriverWait(driver, 10).until(
            lambda the_driver: the_driver.find_element_by_id('selShouShuList').is_displayed())
    surgeryList = driver.find_element_by_id('selShouShuList')
    options = surgeryList.find_elements_by_tag_name('option')
    if len(options)==0:
        print "手术次数未获取到"
        return "error"
    for i in range(len(options)):   #遍历手术次数
        WebDriverWait(driver, 10).until(
            lambda the_driver: the_driver.find_element_by_id('selShouShuList').is_displayed())
        surgeryList.click()
        surgeryList.find_elements_by_tag_name('option')[i].click()
        sleep(2)

        driver.find_element_by_id('txtCheckReportType').click()
        driver.find_element_by_css_selector("div[data-text='病理形态学']").click()   #选择病理形态学的报告类别
        WebDriverWait(driver, 10).until(
            lambda the_driver: the_driver.find_element_by_id('tbodyReportList').is_displayed())
        tbodyReportList = driver.find_element_by_id('tbodyReportList')
        reportList = tbodyReportList.find_elements_by_tag_name('tr')
        for j in range(len(reportList)):
            reportTitle = reportList[j].find_elements_by_tag_name('td')[0].get_attribute('title')
            if (u'常规组织病理' in reportTitle or u'手术标本' in reportTitle):
                reportList[j].click()
                sleep(1)
                #判断是否有alert
                alert = EC.alert_is_present()(driver)
                if alert:
                    alert.accept()
                checkResult = driver.find_element_by_xpath('//div[ @ id = "divPathology"]/div[1]/div[5]/div[2]').text
                checkConclusion = driver.find_element_by_xpath('//div[ @ id = "divPathology"]/div[1]/div[6]/div[2]').text
                WebDriverWait(driver, 10).until(
                    lambda the_driver: the_driver.find_element_by_id('btnTest').is_displayed())

                driver.find_element_by_id('btnTest').click()    #点击校验按钮
                suoj_Fo = driver.find_element_by_xpath('//div[@id="divPathology"]/div[3]/div[5]/div[2]').text   #腺内灶所见
                zhend_Fo = driver.find_element_by_xpath('//div[@id="divPathology"]/div[3]/div[6]/div[2]').text  #腺内灶诊断
                suoj_Ln = driver.find_element_by_xpath('//div[@id="divPathology"]/div[3]/div[7]/div[2]').text   #淋巴结所见
                zhend_Ln = driver.find_element_by_xpath('//*[@id="divPathology"]/div[3]/div[8]/div[2]').text    #淋巴结诊断
                operateExcel.WriteExcel(Hid,'A'+str(n),'pathology')
                operateExcel.WriteExcel(checkResult,'B'+str(n),'pathology')
                operateExcel.WriteExcel(checkConclusion,'C'+str(n),'pathology')
                operateExcel.WriteExcel(suoj_Fo,'D'+str(n),'pathology')
                operateExcel.WriteExcel(zhend_Fo,'E'+str(n),'pathology')
                operateExcel.WriteExcel(suoj_Ln,'F'+str(n),'pathology')
                operateExcel.WriteExcel(zhend_Ln,'G'+str(n),'pathology')
                n = n+1
                nowTime = time.strftime("%Y%m%d.%H.%M.%S")
                driver.get_screenshot_as_file(r'../PageScreen/%s.jpg' % nowTime)
                driver.find_element_by_xpath('//span[@data-cmd="InsertOrderedList-xnz"]').click()
                driver.find_element_by_xpath('//span[@data-cmd="InsertOrderedList-ln"]').click()
                sleep(1)
                nowTime = time.strftime("%Y%m%d.%H.%M.%S")
                driver.get_screenshot_as_file(r'../PageScreen/%s.jpg' % nowTime)
                driver.find_element_by_id('btnQuery').click()
                WebDriverWait(driver, 10).until(
                    lambda the_driver: the_driver.find_element_by_id('tbodyReportList').is_displayed())


def followup_click():
    driver.find_element_by_xpath('//div[@id="divOperationList"]/table/tbody/tr[4]').click()

def yearSelect(value):
    year = Select(driver.find_element_by_name('selNianQi'))
    year.select_by_value(value)
    WebDriverWait(driver, 10).until(
        lambda the_driver: the_driver.find_element_by_id('tbodyReportList').is_displayed())
    alert = EC.alert_is_present()(driver)
    if alert:
        alert.accept()

def jiaoyan_Bchao(Hid):
    try:
        WebDriverWait(driver, 10).until(
            lambda the_driver: the_driver.find_element_by_id('tbodyReportList').is_displayed())
    except Exception as e:
        print e
    alert_close()
    tbodyReportList = driver.find_element_by_id('tbodyReportList')
    reportList = tbodyReportList.find_elements_by_tag_name('tr')
    for j in range(len(reportList)):  # 遍历术前B超报告份数
        reportList[j].click()
        sleep(1)
        # 判断是否有alert
        alert = EC.alert_is_present()(driver)
        if alert:
            if (u'外院手术' in alert.text):
                alert.dismiss()
            else:
                alert.accept()
        undo()
        alert_close()
        WebDriverWait(driver, 20).until_not(
            lambda the_driver: the_driver.find_element_by_xpath('//div[@class="divBlockHid"]').is_displayed())
        undo()
        alert_close()
        checkResult = driver.find_element_by_xpath('//div[@id="divUltrasonography"]/div[1]/div[5]/div[2]').text
        checkConclusion = driver.find_element_by_xpath(
            '//div[@id="divUltrasonography"]/div[1]/div[6]/div[2]').text
        WebDriverWait(driver, 10).until(
            lambda the_driver: the_driver.find_element_by_id('btnTest').is_displayed())
        # driver.find_element_by_id('btnTest').click()  # 点击校验按钮
        suoj_Fo = driver.find_element_by_xpath('//div[@id="divUltrasonography"]/div[3]/div[5]/div[2]').text  # 腺内灶所见
        zhend_Fo = driver.find_element_by_xpath('//div[@id="divUltrasonography"]/div[3]/div[6]/div[2]').text  # 腺内灶诊断
        suoj_Ln = driver.find_element_by_xpath('//div[@id="divUltrasonography"]/div[3]/div[7]/div[2]').text  # 淋巴结所见
        zhend_Ln = driver.find_element_by_xpath('//*[@id="divUltrasonography"]/div[3]/div[8]/div[2]').text  # 淋巴结诊断
        n = operateExcel.max_row('Bchao') + 1
        operateExcel.WriteExcel(Hid, 'A' + str(n), 'Bchao')
        operateExcel.WriteExcel(checkResult, 'B' + str(n), 'Bchao')
        operateExcel.WriteExcel(checkConclusion, 'C' + str(n), 'Bchao')
        operateExcel.WriteExcel(suoj_Fo, 'D' + str(n), 'Bchao')
        operateExcel.WriteExcel(zhend_Fo, 'E' + str(n), 'Bchao')
        operateExcel.WriteExcel(suoj_Ln, 'F' + str(n), 'Bchao')
        operateExcel.WriteExcel(zhend_Ln, 'G' + str(n), 'Bchao')
        n = n + 1



def jiaoyan_ImgB(Hid):
    """影像B校验代码化"""
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
        className = reportList[j].find_element_by_xpath('td[3]/div').get_attribute("class")
        while('StateNoCom'!= className and 'StateCFCom'!= className):
            undo()
            alert_close()
            reportList = tbodyReportList.find_elements_by_tag_name('tr')
            reportList[j].click()  #焦点重新回到该报告上
        checkResult = driver.find_element_by_xpath('//div[@id="divImagingExamination"]/div[1]/div[5]/div[2]').text
        checkConclusion = driver.find_element_by_xpath(
            '//div[@id="divImagingExamination"]/div[1]/div[6]/div[2]').text
        WebDriverWait(driver, 20).until(
            lambda the_driver: the_driver.find_element_by_id('btnCode').is_displayed())
        driver.find_element_by_id('btnCode').click()  # 点击代码化
        AbnormalClassify = driver.find_element_by_xpath('//input[@name="AbnormalClassify"]').text
        n = operateExcel.max_row('ImgB') + 1
        operateExcel.WriteExcel(Hid, ('A' + str(n)), 'ImgB')
        operateExcel.WriteExcel(checkResult, 'B' + str(n), 'ImgB')
        operateExcel.WriteExcel(checkConclusion, 'C' + str(n), 'ImgB')
        operateExcel.WriteExcel(AbnormalClassify, 'D' + str(n), 'ImgB')
        # operateExcel.WriteExcel(zhend_Fo, 'E' + str(n), 'ImgB')
        # operateExcel.WriteExcel(suoj_Ln, 'F' + str(n), 'ImgB')
        # operateExcel.WriteExcel(zhend_Ln, 'G' + str(n), 'ImgB')

        driver.get_screenshot_as_file(imgBscreen+Hid+'.png')
        # driver.find_element_by_id('btnQuery').click()
        WebDriverWait(driver, 10).until(
            lambda the_driver: the_driver.find_element_by_id('tbodyReportList').is_displayed())