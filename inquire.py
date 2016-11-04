#coding:utf-8

import tkMessageBox
from selenium import webdriver
from datetime import  *
import time
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains 
from selenium.webdriver.common.keys import Keys 
import xlsxwriter,xlrd
import ConfigParser
import os

'''
Created on 2016年9月29日
@author: MengLei
'''


def importinfo(arg):
    '''导入ini文件信息'''
    cp = ConfigParser.ConfigParser()
    cp.read('web_ele.ini')
    args={}
    item=cp.items(arg.decode('utf-8').encode('gbk'))
    args.update(item)
    return args

def logoin(b):
    '''登录网页'''
    arg=importinfo('login')
    time.sleep(1)
    b.get(arg['url'])
    b.maximize_window()
    d=WebDriverWait(b, 10).until(lambda c:b.find_element_by_xpath(arg['smcurl_userid']))
    d.clear()
    d.send_keys(arg['username'])
    time.sleep(0.5)
    e=WebDriverWait(b, 10).until(lambda c:b.find_element_by_xpath(arg['smcurl_pwdid']))
    e.clear()
    e.send_keys(arg['userpassword'])
    b.find_element_by_xpath(arg['smcurl_loginid']).click()
    return b

def opencatalogue(b,mydirectory,myrootdirectory):
    '''判断并打开网页子目录'''
    try:
        time.sleep(1)
        logoutbtton = WebDriverWait(b, 15).until(lambda c: b.find_element_by_xpath(mydirectory))
        logoutbtton.click()
    except:
        a=b.find_element_by_xpath(myrootdirectory)
        time.sleep(1)
        ActionChains(b).move_to_element(a).click().perform()
        time.sleep(1)
        logoutbtton = WebDriverWait(b, 15).until(lambda c: b.find_element_by_xpath(mydirectory)) 
        logoutbtton.click()
        
def opentreepath(b):
    '打开左侧目录'
    arg=importinfo('基本配置')
    time.sleep(1)
    myrootdirectory=arg.pop(('基站').decode('utf-8').encode('gbk'))
    mydirectory=arg.pop(('基站_配置').decode('utf-8').encode('gbk'))
    opencatalogue(b,mydirectory,myrootdirectory)
    time.sleep(0.5)
    
def inquire(b):
    '1.按当天日期创建Excel文件准备写入，创建图形库准备弹窗提示'
    if  os.path.exists(os.getcwd()+os.sep+'report'+os.sep+'LST_RESULT-'+datetime.now().date().isoformat()+'.xlsx'):
        excelnum=0
        while 1:
            excelnum +=1
            #===============================================================
            # '目前不限制每日新建EXCEL文件数量，暂时注释掉'
            # if excelnum >20:
            #     break
            #===============================================================
            if not os.path.exists(os.getcwd()+os.sep+'report'+os.sep+'LST_RESULT-'+datetime.now().date().isoformat()+'-'+str(excelnum)+'.xlsx'):
                workbook = xlsxwriter.Workbook(os.getcwd()+os.sep+'report'+os.sep+'LST_RESULT-'+datetime.now().date().isoformat()+'-'+str(excelnum)+'.xlsx')
                worksheet1 = workbook.add_worksheet()
                print '生成Excel文件为： '+os.getcwd()+os.sep+'report'+os.sep+'LST_RESULT-'+datetime.now().date().isoformat()+'-'+str(excelnum)+'.xlsx'
                #===========================================================
                # '目前不需要读取模块，暂时注释掉'
                # readexcel = xlrd.open_workbook(u'基站信息查询'+datetime.now().date().isoformat()+u'号'+str(excelnum)+'.xlsx')
                 #===========================================================
                break
    else: 
        workbook = xlsxwriter.Workbook(os.getcwd()+os.sep+'report'+os.sep+'LST_RESULT-'+datetime.now().date().isoformat()+'.xlsx')
        worksheet1 = workbook.add_worksheet()
        print '生成Excel文件为： '+os.getcwd()+os.sep+'report'+os.sep+'LST_RESULT-'+datetime.now().date().isoformat()+'.xlsx'
        #===================================================================
        # '目前不需要读取模块，暂时注释掉'
        # readexcel = xlrd.open_workbook(u'基站信息查询'+datetime.now().date().isoformat()+'.xlsx')
        #===================================================================

    #=======================================================================
    # '目前不限制每日新建EXCEL文件 数量，暂时注释掉'
    # msgbox='今日创建文件过多,已生成超过20个EXCLE文件，请删除D:\OMC\文件夹下，'+datetime.now().date().isoformat()+'号，相关<基站信息查询>excel文件后，重新运行脚本！'
    # print  msgbox
    # msg =tkMessageBox.showwarning(u'你好', u'尽快了解了')
    #=======================================================================
    
    
    '2.打开目录，导入元素'
    opentreepath(b)
    a=importinfo('页面元素')
    arg=importinfo('查询')
    list=[]
    list=arg.values()

    '3填入基站ID，并选定，执行'
    ddd =b.find_element_by_xpath(a[('查询输入').decode('utf-8').encode('gbk')])
    ddd.send_keys(a[('基站_移动').decode('utf-8').encode('gbk')])
    time.sleep(1)
    b.find_element_by_xpath(a[('查询按钮').decode('utf-8').encode('gbk')]).click()
    time.sleep(1)
    b.find_element_by_xpath(a[('选定按钮').decode('utf-8').encode('gbk')]).click()
    
    '4循环打开下拉列表，选择各个元素，并保存查询结果'
    nrows = 0
    for i in range(1,len(list)+1):
#     for round in range(1,2):
        '1).打开下拉列表，选择查找元素，并点击执行'
        print '**************************************开始%d项测试**************共%d项*************'%(i,len(list))
        aa=b.find_element_by_xpath(a['select'])
        time.sleep(1)
        ActionChains(b).move_to_element(aa).double_click().perform()
        time.sleep(1)
        print '本次查询元素为',list[(i-1)]
        pp = b.find_element_by_xpath(list[(i-1)])
        ActionChains(b).move_to_element(pp).double_click().perform()
        b.find_element_by_xpath(a[('执行').decode('utf-8').encode('gbk')]).click()
        time.sleep(1)
        nline=0
        
        '2).获得1、2行的查询结果，并保存到Excel表格'
        '/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[元素]/div[索引]/div[列]/span[行]'
        for m in range(1,2):
            '1)).获取第1和2行的结果（1查询项、2基站ID），并保存Excel'
            f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/span[%d]'%(i,m)
            try:
                g=WebDriverWait(b,300).until(lambda x: b.find_element_by_xpath(f)).text
                worksheet1.write(nrows,0,g)
                nrows+=1
            except:
                worksheet1.write(nrows, 0,'等待结果超时，本元素终止查询')
                nrows+=1
                print '等待结果超时，本项查询终止，进行下一项查询！'
                break
            finally:
                print '查询第%d行的结果完成，数据已写入Excel'%m
             
             
        '3).获得所有索引（3、4、5、6、7、8...行）的结果'
        '(1).判断有多少个索引'   
        try:
            for nlin in range(1,2):
                f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/span[%d]'%(i,nlin)
                b.find_element_by_xpath(f)
                nline=1
    #                     print '结果返回错误值，将索引直接定义为1行'
        except:
            try:
                for nlin in range(1,22):
                    f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/div[%d]'%(i,nlin)
                    b.find_element_by_xpath(f)
                    nline=nlin
                '一共有行索引'
            except:
                pass
            print '查询结果一共包含%d行'%(nline*2+2)
            
            
        '(2).判断索引有几列*******************'
        all_col=[]
        for n in range(1,22):
            f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/div/div[%d]'%(i,n)
            try:
                findele=b.find_element_by_xpath(f)
                all_col.append(findele)
                'all_col为列数'
            except:
                if len(all_col)>=1:
                    '证明是有结果，可能只有1列，则不在进行查找'
                    pass
                else:
    #                             print '没有返回正确结果，xpath查找错误的返回结果'
                    if m==3 and n ==1:
                        f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/span[3]'%(i)
                        try:
                            findele=b.find_element_by_xpath(f)
                            all_col.append(findele)
    #                                     print '错误的返回结果,强制定为只有1列（实际情况也只有1列）'
                        except:
                            pass
    #                 print 'all_col有%d列,具体元素为：'%(len(all_col)),all_co     
       
                    
        '(3).获取内容导出到Excel*******************'
        '''i项目，ii,nline索引，iii行数（共2行）,n列数  '''
        for ii in range(1,nline+1):
            for iii in range(1,3):
                for n in range(1,len(all_col)+1):
    
                    '1.获取第3和第4行的正常结果，具体数值'
                    fff='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/div[%d]/div[%d]/span[%d]'%(i,ii,n,iii)
    #                             print 'm为：%d，查询元素的具体xpath为： %s'%(iii,f)
                    try:
                        g = b.find_element_by_xpath(fff).text
                        worksheet1.write(nrows,n-1,g)
                    except:
    
                        
                        '2.获取第3和第4行的异常结果获取,，写入excel'
                        if ii==1 and iii==1 and n ==1:
                            'iii==1(第三行) and n ==1（第一列）'
                            f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/span[3]'%(i)
                        elif  ii==1 and iii==2 and n==1:
                            'iii==2（第四行） and n==1（第一列）'
                            f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/span[4]'%(i)
                        else:
                            f=''
                        try:
                            g = b.find_element_by_xpath(f).text
                            worksheet1.write(nrows,n-1,g)
                        except:
                            break
    
                nrows+=1
                time.sleep(0.2)
                '一行内容填写完成'
                print '第%d行的结果已写入Excel'%(2*ii+iii)
                        
        time.sleep(1)
        print '第%d项测试已全部保存'%(i)
    
    time.sleep(1)
    workbook.close()

    '退出登录'
    arg=importinfo('logout') 
    time.sleep(3)
    logoutbtton = WebDriverWait(b, 10).until(lambda c: b.find_element_by_xpath(arg['smcurl_logouid'])) 
    ActionChains(b).click(logoutbtton).perform()
    time.sleep(1)
    logoutbtton2 =b.find_element_by_xpath(arg['smcurl_logouid2'])
    ActionChains(b).move_to_element(logoutbtton2).click().perform()
    print '退出系统登录'
    
    print '**************************测试完成，关闭测试！！************************************'
    
if __name__ == '__main__':
    b = webdriver.Firefox()
    b = logoin(b)
    inquire(b)