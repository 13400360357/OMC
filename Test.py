
#coding:utf-8
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

# print 'hello'
# b=webdriver.Firefox()
# print '1'
# b.get('http://192.168.9.32:8080/smallcell/')
# print '2'
# b.maximize_window()
# print '3'