#coding:utf-8

import tkMessageBox
from omc_modify import *
from selenium import webdriver
from datetime import  *
import time
from selenium.webdriver.common.action_chains import ActionChains 
from selenium.webdriver.common.keys import Keys 
from selenium.webdriver.support.ui import WebDriverWait
import xlsxwriter,xlrd
import ConfigParser
import os
from inquire import *
from modify2 import *
from modify import *
from compare import *


'''
Created on 2016年9月30日

@author: MengLei
'''
if __name__ == '__main__':
    b = webdriver.Firefox()
     
    '修改最大值保存下发结果，并再次查询基站上报结果，最后进行对比'
    logoin(b)
    modify(b)
    logoin(b)    
    inquire(b)


    excel1='MOD_RESULT_2-'+str(datetime.now().date().isoformat())+'-1.xlsx'
    excel2='MOD_RESULT_2-'+str(datetime.now().date().isoformat())+'.xlsx'
    print '1excel1',excel1
    print '1excel2',excel2

    excel2 = 'LST_RESULT-'+str(datetime.now().date().isoformat())+'.xlsx'
    excel1 = 'MOD_RESULT-'+str(datetime.now().date().isoformat())+'.xlsx'
    compare(excel1,excel2)
     
     
    '修改正常值保存下发结果，并再次查询基站上报结果，最后进行对比'
    logoin(b)        
    modify2(b)
    logoin(b)        
    inquire(b)
     
    excel1='MOD_RESULT_2-'+str(datetime.now().date().isoformat())+'.xlsx'
    excel2='LST_RESULT-'+str(datetime.now().date().isoformat())+'-1.xlsx'
    print '2 excel1', excel1
    print '2excel2',excel2
    compare(excel1,excel2)    
     
    
    
    
    
    
    
    
    
    