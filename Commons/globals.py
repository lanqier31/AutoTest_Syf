# -*-coding:utf-8-*-
import Config
from openpyxl import Workbook

fileName = Config.log_file_path+'/logfile.txt'
txtPath = Config.log_file_path+'/shuqian2.txt'

def log(content):
    f = file(fileName, "a+")
    f.write(content+'\\n')
    f.close()


def readtxtToExcel(excelName):
    f = open(txtPath,'r')
    lines = f.readlines()
    # 新建一个excel文件
    wb = Workbook()
    # 新建一个sheet
    sheet = wb.create_sheet('Data', index=1)
    i = 1
    for line in lines:
        for r in line.split(','):
            sheet['A'+str(i)] = r
            i = i + 1
    wb.save(excelName+'.xlsx')
    f.close()

