#coding=utf-8
'''
Created on 2015年1月16日
'''
import xlrd
import sys
from selenium import webdriver
sys.path.append('..')

path = r'..\data\data.xls'
data = xlrd.open_workbook(path)

def getCaseName():
    '''获取测试用例名称'''
    table = data.sheet_by_name('Test Cases')
    li = table.col_values(0)[1:]
    return li

def getCaseStep(li_casenames):
    '''获取测试用例的步骤数'''
    table = data.sheet_by_name('Test Steps')
    steps = table.col_values(0)[1:]
    for li in li_casenames:
        i = 0
        for stepname in steps:
            if li == stepname:
                i+=1
        if i!=0:
            print '%s have %d ' %(li,i)
            getKeyWord_Object_value(li, i)

def getKeyWord_Object_value(casename,steps):
    table = data.sheet_by_name('Test Steps')
    keyword = table.col_values(3)[1:steps+1]
    obj = table.col_values(4)[1:steps+1]
    print zip(keyword,obj)

def action(steps):
    pass
    
    

if __name__ == '__main__':
    li_casenames = getCaseName()
    print type(li_casenames)
    getCaseStep(li_casenames)

