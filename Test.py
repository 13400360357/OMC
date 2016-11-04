#coding:utf-8
import os
from  datetime import*
'''
@author: MengLei
'''
from selenium import webdriver
import time

b=webdriver.Firefox()
b.get('http://192.168.9.70:9999/WebRoot/')
time.sleep(0.5)
# print b.find_element_by_xpath('/html/body').text
print 'haha'
print '哈哈'
print '生成Excel文件为： '+os.getcwd()+os.sep+'report'+os.sep+'LST_RESULT-'+datetime.now().date().isoformat()+'.xlsx'
print 'lala'
# print 'hello'
# b=webdriver.Firefox()
# print '1'
# b.get('http://192.168.9.32:8080/smallcell/')
# print '2'
# b.maximize_window()
# print '3'