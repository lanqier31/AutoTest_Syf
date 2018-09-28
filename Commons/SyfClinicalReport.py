# -*- coding: utf-8 -*-
# Author：WeirGao

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
from selenium.webdriver.support import expected_conditions as EC
from openpyxl.reader.excel import load_workbook
from datetime import datetime
# from openpyxl.styles import Font, colors, Alignment
import Config
import time
import operateExcel,globals
from time import sleep
#进入初诊页面

driver=Config.ChromeDriver
autocase = Config.autocase_path
book = load_workbook(autocase)

#病历号输入框
txtHid ="$('#txtHospitalNumber')"

def goto_Report():

    url = 'http://'+Config.IP+Config.Version+'/SyfHospitalClinicalDataCenter/QueryReport?type=add&editmark=true'
    driver.get(url)

def input_Hid(hid):
    '''输入病历号'''
    # exec_js = "$('#txtHospitalNumber').val(hid)"
    driver.execute_script(txtHid+'.val("' + hid + '")')
    alert=alert_close()
    if alert =='不存在该病历号':
        globals.log(str(hid)+u'没有此病人信息')
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
            return "不存在该病历号"
        else:
            alert.accept()

#截图
def screenpatho():
    nowTime = time.strftime("%Y%m%d.%H.%M.%S")
    driver.get_screenshot_as_file(r'../PageScreen/pathology/%s.png' % nowTime)


def screenImgB():
    nowTime = time.strftime("%Y%m%d.%H.%M.%S")
    driver.get_screenshot_as_file(r'../PageScreen/ImgB/%s.png' % nowTime)

def screenAssayA():
    nowTime = time.strftime("%Y%m%d.%H.%M.%S")
    driver.get_screenshot_as_file(r'../PageScreen/AssayA/%s.png' % nowTime)

def screenBchao():
    nowTime = time.strftime("%Y%m%d.%H.%M.%S")
    driver.get_screenshot_as_file(r'../PageScreen/Bchao/%s.png' % nowTime)

def screenPatho(Hid):
    driver.get_screenshot_as_file(pathoscreen + Hid + '.png')

def undo():
    driver.find_element_by_id('btnRes').click()
    sleep(1)
    alert_close()
    sleep(1)


def selectImgB():
    """ 报告类型列表中选择影像B"""
    wait_loading()
    driver.find_element_by_id('txtCheckReportType').click()
    sleep(1)
    clickNum = 0
    ImgB = driver.find_element_by_css_selector("div[data-text='影像B']")
    while not (ImgB.is_displayed()):
        if clickNum ==2:
            return "影像B的报告类别没有找到"
        else:
            driver.find_element_by_id('txtCheckReportType').click()
            clickNum= clickNum+1
    ImgB.click()  # 选择病理形态学的报告类别
    WebDriverWait(driver, 10).until(
        lambda the_driver: the_driver.find_element_by_id('tbodyReportList').is_displayed())
    sleep(1)


def selectAssayA():
    wait_loading()
    driver.find_element_by_id('txtCheckReportType').click()
    sleep(1)
    clickNum = 0
    AssayA = driver.find_element_by_css_selector("div[data-text='化验A']")
    while not (AssayA.is_displayed()):
        if clickNum == 2:
            return "化验A的报告类别没有找到"
        else:
            driver.find_element_by_id('txtCheckReportType').click()
            clickNum = clickNum + 1
    AssayA.click()  # 选择病理形态学的报告类别
    WebDriverWait(driver, 10).until(
        lambda the_driver: the_driver.find_element_by_id('tbodyReportList').is_displayed())
    sleep(1)


def yearSelect(value):
    sleep(2)
    alert_close()
    WebDriverWait(driver, 10).until(
        lambda the_driver: the_driver.find_element_by_id('divOperationList').is_displayed())
    # 读取首次手术下的随访
    table = driver.find_element_by_xpath('//div[@id="divOperationList"]/table[last()]')
    sleep(1)
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
        wait_loading()
        reportList = tbodyReportList.find_elements_by_tag_name('tr')
        if (len(reportList) <= j):
            return (Hid + u"报告日期超出范围")
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
            wait_loading()
            undo()
            alert_close()
            wait_loading()
            reportList = tbodyReportList.find_elements_by_tag_name('tr')
            if(len(reportList)<=j):
                return (Hid+u"报告日期超出范围")
            reportList[j].click()  # 焦点重新回到该报告上
            alert_close()
            className = reportList[j].find_element_by_xpath('td[3]/div').get_attribute("class")
        wait_loading()
        checkResult = driver.find_element_by_xpath('//div[@id="divUltrasonography"]/div[1]/div[5]/div[2]').text
        checkConclusion = driver.find_element_by_xpath(
            '//div[@id="divUltrasonography"]/div[1]/div[6]/div[2]').text
        WebDriverWait(driver, 10).until(
            lambda the_driver: the_driver.find_element_by_id('btnTest').is_displayed())
        # driver.find_element_by_id('btnTest').click()  # 点击校验按钮
        jiaoyan_format_Bchao()  #B超校验时对淋巴结进行格式化处理
        result = readBCtext()
        sheet['A'+str(n)] = Hid
        sheet['B'+str(n)] = checkResult
        sheet['C'+str(n)] = checkConclusion
        sheet['D'+str(n)] = result['suoj_xy']
        sheet['E'+str(n)] = result['zhend_xy']
        sheet['F'+str(n)] = result['suoj_xc']
        sheet['G'+str(n)] = result['zhend_xc']
        sheet['H'+str(n)] = result['suoj_qc']
        sheet['I'+str(n)] = result['zhend_qc']
        sheet['J'+str(n)] = result['suoj_pl']
        sheet['K'+str(n)] = result['zhend_pl']
        book.save(autocase)
        n = n+1
        screenBchao()
        driver.find_element_by_id('btnCode').click()  # 点击代码化
        sleep(3)
        Codepai_Bchao() # 代码化中所有的派生
        screenBchao()
        driver.find_element_by_id('btnQuery').click() #返回
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
        wait_loading()
        # driver.find_element_by_id('btnTest').click()  # 点击代码化
        jiaoyanText = ImgBjiaoyantext()
        screenImgB()
        btnCode = driver.find_element_by_id('btnCode')  # 点击代码化
        while not btnCode.is_displayed():
            driver.find_element_by_id('btnTest').click()
        btnCode.click()
        sleep(2)
        wait_loading()
        FindingA = driver.find_element_by_xpath('//div[@name = "MajorAnomalies"]').text
        # AbnormalClassify = driver.find_element_by_xpath('//input[@name="AbnormalClassify"]').get_attribute('value')
        inputliebie = 'return $(\'input[name = "AbnormalClassify"]\').val()'
        Leibie = driver.execute_script(inputliebie)
        sheet['A' + str(n)] = Hid
        sheet['B' + str(n)] = checkResult
        sheet['C' + str(n)] = checkConclusion
        sheet['D' + str(n)] = jiaoyanText['suoj_fsz']
        sheet['E' + str(n)] = jiaoyanText['suoj_xm']
        sheet['F' + str(n)] = jiaoyanText['suoj_fm']
        sheet['G' + str(n)] = FindingA
        sheet['H' + str(n)] = Leibie
        book.save(autocase)
        n = n+1
        screenImgB()
        # nowTime = time.strftime("%Y%m%d.%H.%M.%S")
        # driver.get_screenshot_as_file(r'../PageScreen/ImgB/%s.jpg' % nowTime)


def jiaoyan_AssayA(Hid):
    """影像B校验代码化"""
    # sheet = book['AssayA']
    try:
        WebDriverWait(driver, 10).until(
            lambda the_driver: the_driver.find_element_by_id('tbodyReportList').is_displayed())
    except Exception as e:
        print e
    alert_close()
    tbodyReportList = driver.find_element_by_id('tbodyReportList')
    reportList = tbodyReportList.find_elements_by_tag_name('tr')
    if (len(reportList) == 1):
        if (reportList[0].text == '没有报告信息.'):
            return Hid + u'无化验A报告'
    for j in range(len(reportList)):  # 遍历报告份数
        sleep(2)
        reportList = tbodyReportList.find_elements_by_tag_name('tr')
        wait_loading()
        reportList[j].click()
        wait_loading()
        if j == 0 and len(reportList)>1:
            reportdate = datetime.strptime(reportList[j].find_element_by_xpath('td[2]').text, '%Y-%m-%d')
            nextdate = datetime.strptime(reportList[j + 1].find_element_by_xpath('td[2]').text, '%Y-%m-%d')
            next = (nextdate - reportdate).days
            if abs(next)==1:
                wait_loading()
                reportList[j].click()
                wait_loading()
                element= reportList[j+1]
                target = driver.find_element_by_id('divBiochemical')
                sleep(1)
                ActionChains(driver).drag_and_drop(element, target).perform()
                sleep(1)
        elif 0< j < (len(reportList)-1):
            predate = datetime.strptime(reportList[j-1].find_element_by_xpath('td[2]').text,'%Y-%m-%d')
            reportdate = datetime.strptime(reportList[j].find_element_by_xpath('td[2]').text,'%Y-%m-%d')
            nextdate = datetime.strptime(reportList[j+1].find_element_by_xpath('td[2]').text,'%Y-%m-%d')
            pre = (reportdate-predate).days
            next = (nextdate-reportdate).days
            if abs(pre)<=1:
                continue
            if abs(next)==1:
                wait_loading()
                reportList[j].click()
                wait_loading()
                element= reportList[j+1]
                target = driver.find_element_by_id('divBiochemical')
                sleep(1)
                ActionChains(driver).drag_and_drop(element, target).perform()
                sleep(1)
        elif j!=0 and j ==len(reportList)-1:
            predate = datetime.strptime(reportList[j - 1].find_element_by_xpath('td[2]').text, '%Y-%m-%d')
            reportdate = datetime.strptime(reportList[j].find_element_by_xpath('td[2]').text, '%Y-%m-%d')
            pre = (reportdate - predate).days
            if abs(pre)<=1:
                continue
        wait_loading()
        driver.find_element_by_id('btnCode').click()
        sleep(2)
        screenAssayA()
        # driver.find_element_by_id('btnSave').click()
        # sleep(1)

def jiaoyan_BFna_cell(Hid):
    """细胞病理学校验代码化"""
    sheet = book['Bfna_cell']
    try:
        WebDriverWait(driver, 10).until(
            lambda the_driver: the_driver.find_element_by_id('tbodyReportList').is_displayed())
    except Exception as e:
        print e
    alert_close()

    tbodyReportList = driver.find_element_by_id('tbodyReportList')
    reportList = tbodyReportList.find_elements_by_tag_name('tr')
    if (len(reportList) == 1):
        if (reportList[0].text == '没有报告信息.'):
            return Hid + u'无报告'
    n = operateExcel.max_row('Bfna_cell') + 1
    for j in range(len(reportList)):  # 遍历报告份数
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
        while ('StateNoCom' != className and 'StateCFCom' != className):
            WebDriverWait(driver, 20).until_not(
                lambda the_driver: the_driver.find_element_by_class_name('divBlockHid').is_displayed())
            undo()
            alert_close()
            reportList = tbodyReportList.find_elements_by_tag_name('tr')
            reportList[j].click()  # 焦点重新回到该报告上
            className = reportList[j].find_element_by_xpath('td[3]/div').get_attribute("class")
        checkResult = driver.find_element_by_xpath('//div[@id="divImagingExamination"]/div[1]/div[5]/div[2]').text
        checkConclusion = driver.find_element_by_xpath(
            '//div[@id="divImagingExamination"]/div[1]/div[6]/div[2]').text
        WebDriverWait(driver, 20).until_not(
            lambda the_driver: the_driver.find_element_by_class_name('divBlockHid').is_displayed())
        driver.find_element_by_id('btnCode').click()  # 点击代码化
        sleep(2)
        wait_loading()
        FindingA = driver.find_element_by_xpath('//div[@name = "MajorAnomalies"]').text
        AbnormalClassify = driver.find_element_by_xpath('//input[@name="AbnormalClassify"]').get_attribute('value')
        inputliebie = '$(\'input[name = "AbnormalClassify"]\').val()'
        Leibie = driver.execute_script(inputliebie)
        sheet['A' + str(n)] = Hid
        sheet['B' + str(n)] = checkResult
        sheet['C' + str(n)] = checkConclusion
        sheet['D' + str(n)] = FindingA
        sheet['E' + str(n)] = AbnormalClassify
        book.save(autocase)
        n = n + 1
        screenImgB()


def readBCtext():
    """"读取超声校验报告中各个框的值"""
    suoj_xy = "return $('#divUltrasonography > div.divRightBlc03 > div:nth-child(5) > div.divSuojInput').html();"
    zhend_xy = "return $('#divUltrasonography > div.divRightBlc03 > div:nth-child(6) > div.divZhedInput').html();"
    suoj_xc = "return $('#divUltrasonography > div.divRightBlc03 > div:nth-child(7) > div.divZhedInput').html();"
    zhend_xc = "return $('#divUltrasonography > div.divRightBlc03 > div:nth-child(8) > div.divZhedInput').html();"
    suoj_qc = "return $('#divUltrasonography > div.divRightBlc03 > div:nth-child(9) > div.divSuojInput').html();"
    zhend_qc = "return $('#divUltrasonography > div.divRightBlc03 > div:nth-child(10) > div.divZhedInput').html();"
    suoj_pl = "return $('#divUltrasonography > div.divRightBlc03 > div:nth-child(11) > div.divZhedInput').html();"
    zhend_pl = "return $('#divUltrasonography > div.divRightBlc03 > div:nth-child(12) > div.divZhedInput').html();"
    return {
        "suoj_xy":driver.execute_script(suoj_xy),
        "zhend_xy":driver.execute_script(zhend_xy),
        "suoj_xc":driver.execute_script(suoj_xc),
        "zhend_xc":driver.execute_script(zhend_xc),
        "suoj_qc":driver.execute_script(suoj_qc),
        "zhend_qc":driver.execute_script(zhend_qc),
        "suoj_pl":driver.execute_script(suoj_pl),
        "zhend_pl":driver.execute_script(zhend_pl),
    }


def ImgBjiaoyantext():
    """读取影像B校验报告中的各个所见框的值"""
    suoj_fsz ="return $('#divImagingExamination > div.divRightBlc03 > div:nth-child(5) > div.divSuojInput').html();"
    suoj_xm = "return $('#divImagingExamination > div.divRightBlc03 > div:nth-child(6) > div.divZhedInput').html();"
    suoj_fm = "return $('#divImagingExamination > div.divRightBlc03 > div:nth-child(7) > div.divSuojInput').html();"

    return {
        "suoj_fsz":driver.execute_script(suoj_fsz),
        "suoj_xm":driver.execute_script(suoj_xm),
        "suoj_fm":driver.execute_script(suoj_fm),
    }


def Codepai_Bchao():
    """"B超代码化中的派生"""
    codePai = {
    "xnz" : driver.find_element_by_xpath('//div[@id="divUltrasonography"]/div[5]/div[3]/div[1]/div'),
    "ln" : driver.find_element_by_xpath('//div[@id="divUltrasonography"]/div[5]/div[5]/div[1]/div'),
    "ln_qs" : driver.find_element_by_xpath('//div[@id="divUltrasonography"]/div[5]/div[6]/div[1]/div'),
    "ln_pl" : driver.find_element_by_xpath('//div[@id="divUltrasonography"]/div[5]/div[7]/div[1]/div'),
    "total" : driver.find_element_by_xpath('//div[@id="divUltrasonography"]/div[5]/div[8]/div[1]/div')}
    for key in codePai:
        if codePai[key].is_displayed():
            codePai[key].click()
            sleep(1)


def jiaoyan_format_Bchao():
    """B超校验时对淋巴结进行格式化处理"""
    order_ln1 = driver.find_element_by_xpath('//span[@data-cmd="InsertOrderedList-Bchao_01"]')
    order_ln2 = driver.find_element_by_xpath('//span[@data-cmd="InsertOrderedList-Bchao_02"]')
    if order_ln1.is_displayed():
        order_ln1.click()
        sleep(1)
    if order_ln2.is_displayed():
        order_ln2.click()
        sleep(1)

def is_element_visible(element):
    """判断是元素是否可见"""
    try:
        the_element = EC.visibility_of_element_located(element)
        assert the_element(driver)
        flag = True
    except:
        flag = False
    return flag


def wait_loading():
    """waiting for divBlockHid disappear"""
    WebDriverWait(driver, 20).until_not(
        lambda the_driver: the_driver.find_element_by_class_name('divBlockHid').is_displayed())
