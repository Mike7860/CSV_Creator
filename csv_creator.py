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

#main class
class Test:
    def __init__(self, id, source, flag, auto):
        self.ID = id
        self.source = source
        self.flag = flag
        self.auto = auto

#dictionaries declaration
dict_test_all = {}
dict_test_basic = {}
dict_test_extended = {}

with open('all.csv', newline='') as csvfile:
    all = csv.reader(csvfile)
    for row in all:
        #
        if row[0] == "Test ID" and row[1] == "Source" and row[2] == "Flag" and row[3] == "Automate":
            print("ok")
            continue
        dict_test_all[row[0]] = Test(row[0],row[1],row[2],row[3])
    print(dict_test_all)
workbook_1 = openpyxl.load_workbook('not_test_1.xlsx', 'rb', data_only=True)
not_test_1 =  workbook_1['not_tested']
all_rows = not_test_1.max_row - 1
nt1 = "D" + str(all_rows+1)
nt1_range = not_test_1['A2':nt1]

print('Please select scope of testing: \n1) basic \n2) extended\n>>')
choise = input ()

if choise == '1':
    for A2, B2, C2, D2 in nt1_range:
        if str(B2.value) != "" and dict_test_all[B2.value].ID == str(B2.value):
            dict_test_all[B2.value].flag = "2"
            print("Not_testable: %s" %(dict_test_all[B2.value].ID))
            not_tested = not_tested + 1
    print("Not tested (based on first file)", not_tested)

if choise == '2':
    for A2, B2, C2, D2 in nt1_range:
        if str(C2.value) != "" and dict_test_all[C2.value].ID == str(C2.value):
            dict_test_all[C2.value].flag = "2"
            print("Not_testable: %s" %(dict_test_all[C2.value].ID))
            not_tested = not_tested + 1
    print("Not tested (based on first file)", not_tested)