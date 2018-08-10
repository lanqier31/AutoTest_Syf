# -*-coding:utf-8-*-
import Config
from Commons import operateExcel
from openpyxl.reader.excel import load_workbook

fileName = Config.log_file_path+'logfile.txt'


def log(content):
    f = file(fileName, "a+")
    f.write(content+'\\n')
    f.close()

