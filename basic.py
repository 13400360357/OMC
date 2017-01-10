#coding:utf-8
'''# Created on 2016年9月29日
@author: MengLei'''            

from selenium import webdriver
from datetime import  *
import time,os,xlsxwriter,xlrd,ConfigParser
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains 
from selenium.webdriver.common.keys import Keys 
from selenium.webdriver.support.ui import Select

class insert_sort:
    def __init__(self,list):
        self.list=list
    def run(self):
        ' 插入排序'
        count = len(self.lists)
        for i in range(1, count):
            key = self.lists[i]
            j = i - 1
            while j >= 0:
                if self.lists[j] > key:
                    self.lists[j + 1] = self.lists[j]
                    self.lists[j] = key
                j -= 1
        return self.lists

class importinfo:
    def run(self,ini,sections):
        '''导入ini文件信息'''
        cp = ConfigParser.ConfigParser()
        cp.read(ini)
        items={}
        item=cp.items(sections)
        items.update(item)
        return items
    
class opencatalogue:
    def importinfo(self,text,sections,catalogue_option,son_catalogue_option):    
        '1. 导入目录xpath'
        ini=importinfo()
        zuocemulu_xpath=ini.run(text,sections)
        time.sleep(1)
        catalogue_xpath=zuocemulu_xpath.pop(catalogue_option)
        son_catalogue_xpath=zuocemulu_xpath.pop(son_catalogue_option)
        return ini,catalogue_xpath,son_catalogue_xpath
    
    def run(self,b,catalogue_xpath,son_catalogue_xpath):
        '2. 判断并打开网页子目录'
        try:
            time.sleep(1)
            logoutbtton = WebDriverWait(b, 3).until(lambda c: b.find_element_by_xpath(son_catalogue_xpath))
            logoutbtton.click()
        except:
            yemianyuansu=b.find_element_by_xpath(catalogue_xpath)
            time.sleep(1)
            ActionChains(b).move_to_element(yemianyuansu).click().perform()
            time.sleep(1)
            print  catalogue_xpath
            print  son_catalogue_xpath
            logoutbtton = WebDriverWait(b, 3).until(lambda c: b.find_element_by_xpath(son_catalogue_xpath)) 
            logoutbtton.click()

class CreateExcel:
        
    '*******************************************************windows系统*******************************************************'
    #===========================================================================
    # def run(self,filena):
    #     '按当天日期创建excel文件'
    #     nrows = 0
    #     if  os.path.exists(os.getcwd()+os.sep+'report'+os.sep+filena+datetime.now().date().isoformat()+'.xlsx'):
    #         excelnum=0
    #         while 1:
    #             excelnum +=1
    #             if not os.path.exists(os.getcwd()+os.sep+'report'+os.sep+filena+datetime.now().date().isoformat()+'-'+str(excelnum)+'.xlsx'):
    #                 workbook = xlsxwriter.Workbook(os.getcwd()+os.sep+'report'+os.sep+filena+datetime.now().date().isoformat()+'-'+str(excelnum)+'.xlsx')
    #                 worksheet1 = workbook.add_worksheet('1')
    #                 worksheet2 = workbook.add_worksheet('2')
    #                 print 'Create Excel： '+os.getcwd()+os.sep+'report'+os.sep+filena+datetime.now().date().isoformat()+'-'+str(excelnum)+'.xlsx'
    #                 break
    #     else:
    #         workbook = xlsxwriter.Workbook(os.getcwd()+os.sep+'report'+os.sep+filena+datetime.now().date().isoformat()+'.xlsx')
    #         worksheet1 = workbook.add_worksheet('1')
    #         worksheet2 = workbook.add_worksheet('2')
    #         print 'Create Excel： '+os.getcwd()+os.sep+'report'+os.sep+filena+datetime.now().date().isoformat()+'.xlsx'    
    #             
    #     return workbook,worksheet1,worksheet2
    #     '*******************************************************windows系统*******************************************************'
    #===========================================================================
          
    '********************************************************Linux系统********************************************************'
    def run(self,filena):
        '按当天日期创建excel文件'
        nrows = 0
        if  os.path.exists('/home/meng/workspace/OMC/report/'+filena+datetime.now().date().isoformat()+'.xlsx'):
            excelnum=0
            while 1:
                excelnum +=1
                if not os.path.exists('/home/meng/workspace/OMC/report/'+filena+datetime.now().date().isoformat()+'-'+str(excelnum)+'.xlsx'):
                    workbook = xlsxwriter.Workbook('/home/meng/workspace/OMC/report/'+filena+datetime.now().date().isoformat()+'-'+str(excelnum)+'.xlsx')
                    worksheet1 = workbook.add_worksheet()
                    worksheet2 = workbook.add_worksheet('2')
                    print 'Create Excel： /home/meng/workspace/OMC/report/'+filena+datetime.now().date().isoformat()+'-'+str(excelnum)+'.xlsx'
                    break
        else:
            workbook = xlsxwriter.Workbook('/home/meng/workspace/OMC/report/'+filena+datetime.now().date().isoformat()+'.xlsx')
            worksheet1 = workbook.add_worksheet()
            worksheet2 = workbook.add_worksheet('2')
            print 'Create Excel：/home/meng/workspace/OMC/report/'+filena+datetime.now().date().isoformat()+'.xlsx'    
                    
        return workbook,worksheet1,worksheet2
        '********************************************************Linux系统********************************************************'

class CreateTxt:
    
    #===========================================================================
    # '*******************************************************windows系统*******************************************************'    
    # def run(self,filena):
    #     '4.按当天日期创建txt文件'
    #     if  os.path.exists(os.getcwd()+os.sep+'report'+os.sep+'err-'+filena+datetime.now().date().isoformat()+'.txt'):
    #         excelnum=0
    #         while 1:
    #             excelnum +=1
    #             if not os.path.exists(os.getcwd()+os.sep+'report'+os.sep+'err-'+filena+datetime.now().date().isoformat()+'-'+str(excelnum)+'.txt'):
    #                 txt = open(os.getcwd()+os.sep+'report'+os.sep+'err-'+filena+datetime.now().date().isoformat()+'-'+str(excelnum)+'.txt','a')
    #                 print 'Create txt： '+os.getcwd()+os.sep+'report'+os.sep+'err-'+filena+datetime.now().date().isoformat()+'-'+str(excelnum)+'.txt'
    #                 break
    #     else:
    #         txt = open(os.getcwd()+os.sep+'report'+os.sep+'err-'+filena+datetime.now().date().isoformat()+'.txt','a')
    #         print 'Create txt： '+os.getcwd()+os.sep+'report'+os.sep+'err-'+filena+datetime.now().date().isoformat()+'.txt'
    #     return txt
    #     '*******************************************************windows系统*******************************************************'
    #===========================================================================
    
    '*******************************************5*************Linux系统********************************************************'
    def run(self,filena):
        '4.按当天日期创建txt文件'
        if  os.path.exists('/home/meng/workspace/OMC/report/err-'+filena+datetime.now().date().isoformat()+'.txt'):
            excelnum=0
            while 1:
                excelnum +=1
                if not os.path.exists('/home/meng/workspace/OMC/report/err-'+filena+datetime.now().date().isoformat()+'-'+str(excelnum)+'.txt'):
                    txt = open('/home/meng/workspace/OMC/report/err-'+filena+datetime.now().date().isoformat()+'-'+str(excelnum)+'.txt','a')
                    print 'Create txt:/home/meng/workspace/OMC/report/err-'+filena+datetime.now().date().isoformat()+'-'+str(excelnum)+'.txt'
                    break
        else:
            txt = open('/home/meng/workspace/OMC/report/err-'+filena+datetime.now().date().isoformat()+'.txt','a')
            print 'Create txt:/home/meng/workspace/OMC/report/err-'+filena+datetime.now().date().isoformat()+'.txt'
        return txt
        '********************************************************Linux系统********************************************************'
class browser:
    def run(self):
#         self.b = webdriver.Chrome()
        self.b=webdriver.PhantomJS(executable_path='/home/meng/Downloads/env/phantomjs-2.1.1-linux-x86_64/bin/phantomjs')
#         self.b=webdriver.PhantomJS(executable_path=r'D:\phantomjs\bin\phantomjs.exe')
        
        self.b.set_window_size(1920, 1080)
        time.sleep(2)
        self.b.maximize_window()
        return self.b
    


class login:
    def __init__(self,b):
        self.b = b

    def run(self):
        im=importinfo()
        inquire=im.run('web_ele.ini','login')
        self.b.get(inquire['url'])
        d=WebDriverWait(self.b, 10).until(lambda c:self.b.find_element_by_xpath(inquire['smcurl_userid']))
        d.clear()
        d.send_keys(inquire['username'])
        time.sleep(0.5)
        e=self.b.find_element_by_xpath(inquire['smcurl_pwdid'])
        e.clear()
        e.send_keys(inquire['userpassword'])
        time.sleep(0.5)
        self.b.find_element_by_xpath(inquire['smcurl_loginid']).click()
        return self.b

class logout:
    def run(self,b):
        im=importinfo()
        logout=im.run('web_ele.ini','logout')
        b.find_element_by_xpath(logout['admin']).click()
        time.sleep(0.5)
        b.find_element_by_xpath(logout['logout']).click()
        
class compare:
    def run(self,excel1,excel2):
        '1. 创建Excel'
        excel=CreateExcel()
        workbook,worksheet=excel.run('compare')
        return workbook,worksheet
        
        '2. 读取文件，并对比'
        readexcel1=xlrd.open_workbook(excel1)
        readsheet1=readexcel1.sheets()[0]
        rows1=readsheet1.nrows
        
        readexcel2=xlrd.open_workbook(excel2)
        readsheel2=readexcel2.sheets()[0]
        rows2=readsheel2.nrows
        
        if rows1 != rows2:
            print 'diff rows,break!'
        else:
            '对每一行进行判断'
            n=0
            for row in range(0,rows1):
                context1=readsheet1.row_values(row)
                context2=readsheel2.row_values(row)
                
                '保存标题行数据'
                if 'RESILT' in context1:
                    con=[]
                    con.append(context1)
                
                '判断本行是否相等 ，不相等，将标题一同写入EXCEL'
                if context1==context2:
                    pass
                else:
                        worksheet.write(n,0,con)
                        n+=1
                        worksheet.write(n,0,'line %d not equal,fault'%row)
                        n+=1
        workbook.close()
        print 'compare finished...'

    