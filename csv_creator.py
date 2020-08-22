import os
import openpyxl
import PyPDF2
import csv
import textwrap
import logging
import time
import datetime

from datetime import date
from openpyxl.workbook import Workbook

curr_dir = os.getcwd()

print('You are in "' +curr_dir+ '" Do you want change directory? y/n')

dir_choise = input()

if dir_choise =='y':
    print('Please provide new directory with necessary files \n>>')
    os.chdir(input())
if dir_choise =='n':
    curr_dir = curr_dir

not_tested = 0
date = date.today()
print(date)

class Test:
    def __init__(self, id, source, flag, auto):
        self.ID = id
        self.source = source
        self.flag = flag
        self.auto = auto

dict_test_all = {}
dict_test_basic = {}
dict_test_extended = {}