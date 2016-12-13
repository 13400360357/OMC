#coding:utf-8

from selenium import webdriver
from datetime import  *
import time
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains 
from selenium.webdriver.common.keys import Keys 
import xlsxwriter
import xlrd
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
    item=cp.items(arg)
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
    e=b.find_element_by_xpath(arg['smcurl_pwdid'])
    e.clear()
    e.send_keys(arg['userpassword'])
    b.find_element_by_xpath(arg['smcurl_loginid']).click()
    return b


def opencatalogue(b,mydirectory,myrootdirectory):
    '''判断并打开网页子目录'''
    try:
        time.sleep(1)
        logoutbtton = WebDriverWait(b, 3).until(lambda c: b.find_element_by_xpath(mydirectory))
        logoutbtton.click()
    except:
        a=b.find_element_by_xpath(myrootdirectory)
        time.sleep(1)
        ActionChains(b).move_to_element(a).click().perform()
        time.sleep(1)
        logoutbtton = WebDriverWait(b, 3).until(lambda c: b.find_element_by_xpath(mydirectory)) 
        logoutbtton.click()
        
def opentreepath(b):
    '打开左侧目录'
    arg=importinfo('jibenpeizhi')
    time.sleep(1)
    myrootdirectory=arg.pop('jizhan')
    mydirectory=arg.pop('jizhanpeizhi')
    opencatalogue(b,mydirectory,myrootdirectory)
    time.sleep(0.5)
    
def inquire(b):
    
    #===========================================================================
    # '********************************************************Linux系统********************************************************'
    # '按当天日期创建excel文件'
    # nrows = 0
    # if  os.path.exists(/home/meng/workspace/OMC/report/MOD_RESULT-'+datetime.now().date().isoformat()+'.xlsx'):
    #     excelnum=0
    #     while 1:
    #         excelnum +=1
    #         if not os.path.exists('/home/meng/workspace/OMC/report/MOD_RESULT-'+datetime.now().date().isoformat()+'-'+str(excelnum)+'.xlsx'):
    #             workbook = xlsxwriter.Workbook('/home/meng/workspace/OMC/report/MOD_RESULT-'+datetime.now().date().isoformat()+'-'+str(excelnum)+'.xlsx')
    #             worksheet1 = workbook.add_worksheet()
    #             print 'Create Excel： /home/meng/workspace/OMC/report/MOD_RESULT-'+datetime.now().date().isoformat()+'-'+str(excelnum)+'.xlsx'
    #             break
    # else:
    #     workbook = xlsxwriter.Workbook('/home/meng/workspace/OMC/report/MOD_RESULT-'+datetime.now().date().isoformat()+'.xlsx')
    #     worksheet1 = workbook.add_worksheet()
    #     print 'Create Excel：/home/meng/workspace/OMC/report/MOD_RESULT-'+datetime.now().date().isoformat()+'.xlsx'    
    #      
    # '********************************************************Linux系统********************************************************'
    #===========================================================================
    
    
    '*******************************************************windows系统*******************************************************'
    '按当天日期创建excel文件'
    nrows = 0
    if  os.path.exists(os.getcwd()+os.sep+'report'+os.sep+'MOD_RESULT-'+datetime.now().date().isoformat()+'.xlsx'):
        excelnum=0
        while 1:
            excelnum +=1
            if not os.path.exists(os.getcwd()+os.sep+'report'+os.sep+'MOD_RESULT-'+datetime.now().date().isoformat()+'-'+str(excelnum)+'.xlsx'):
                workbook = xlsxwriter.Workbook(os.getcwd()+os.sep+'report'+os.sep+'MOD_RESULT-'+datetime.now().date().isoformat()+'-'+str(excelnum)+'.xlsx')
                worksheet1 = workbook.add_worksheet()
                print 'Create Excel： '+os.getcwd()+os.sep+'report'+os.sep+ 'MOD_RESULT-'+datetime.now().date().isoformat()+'-'+str(excelnum)+'.xlsx'
                break
    else:
        workbook = xlsxwriter.Workbook(os.getcwd()+os.sep+'report'+os.sep+'MOD_RESULT-'+datetime.now().date().isoformat()+'.xlsx')
        worksheet1 = workbook.add_worksheet()
        print 'Create Excel： '+os.getcwd()+os.sep+'report'+os.sep+'MOD_RESULT-'+datetime.now().date().isoformat()+'.xlsx'    
      
    '*******************************************************windows系统*******************************************************'
    
    
    '2.打开目录，导入元素'
    opentreepath(b)
    a=importinfo('yemianyuansu')
    arg=importinfo('chaxun')
    list=[]
    list=arg.values()

    '3填入基站ID，并选定，执行'
    time.sleep(1)
    ddd =WebDriverWait(b,20).until(lambda x: b.find_element_by_xpath(a['chaxunshuru']))
    ddd.send_keys(a['jizhan_yidong'])
    time.sleep(1)
    b.find_element_by_xpath(a['chaxunanniu']).click()
    time.sleep(1)
    b.find_element_by_xpath(a['xuandinganniu']).click()
    
    '4循环打开下拉列表，选择各个元素，并保存查询结果'
    nrows = 0
    for i in range(1,len(list)+1):
#     for i in range(1,2):
        '1).打开下拉列表，选择查找元素，并点击执行'
        print '***************************start testing...      total %d      this is %d*******************************'%(i,len(list))
        aa=b.find_element_by_xpath(a['select'])
        time.sleep(1)
#         ActionChains(b).move_to_element(aa).double_click().perform()
        aa.click()
        time.sleep(1)
        print 'ele is:',list[(i-1)]
        pp = b.find_element_by_xpath(list[(i-1)])
        ActionChains(b).double_click(pp).perform()
        b.find_element_by_xpath(a['zhixing']).click()
        time.sleep(1)
        nline=0
        
        '2).获得1、2行的查询结果，并保存到Excel表格'
        '/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[元素]/div[索引]/div[列]/span[行]'
        for m in range(1,3):
            '1)).获取第1和2行的结果（1查询项、2基站ID），并保存Excel'
            f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/span[%d]'%(i,m)
            try:
                g=WebDriverWait(b,300).until(lambda x: b.find_element_by_xpath(f)).text
                worksheet1.write(nrows,0,g)
                nrows+=1
            except:
                worksheet1.write(nrows, 0,'等待结果超时，本元素终止查询')
                nrows+=1
                '等待结果超时，本项查询终止，进行下一项查询！'
                print 'time out... end this ele . starting the next ele'
                break
            finally:
                '查询第%d行的结果完成，数据已写入Excel'%m
                print '%d line is writting into Excel'%m
             
             
        '3).获得所有索引（3、4、5、6、7、8...行）的结果'
        '(1).判断有多少个索引'   
        try:
            
            '结果返回错误值，将索引直接定义为1行'
            for nlin in range(3,5):
                f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/span[%d]'%(i,nlin)
                b.find_element_by_xpath(f)
            nline=1
                   
        except:
            '正常返回值，判断多少个索引'
            try:
                for nlin in range(1,22):
                    f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/div[%d]'%(i,nlin)
                    b.find_element_by_xpath(f)
                    nline=nlin
                '一共有行索引'
            except:
                pass
            '查询结果一共包含%d行'%(nline*2+2)

            
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
                    break
                else:
                    '没有返回正确结果，强制定为只有1列'
                    f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/span[%d]'%(i)
                    try:
                        findele=b.find_element_by_xpath(f)
                        all_col.append(findele)
                        '错误的返回结果,强制定为只有1列（实际情况也只有1列）'
                    except:
                            pass
            'all_col有%d列,具体元素为：'%(len(all_col)),all_col   
       
                    
        '(3).获取内容导出到Excel*******************'
        '''i项目，ii,nline索引，iii行数（共2行）,n列数  '''
        for ii in range(1,nline+1):
            for iii in range(1,3):
                for n in range(1,len(all_col)+1):
    
                    '1.获取第3和第4行的正常结果，具体数值'
                    fff='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/div[%d]/div[%d]/span[%d]'%(i,ii,n,iii)
                    'm为：%d，查询元素的具体xpath为： %s'%(iii,f)
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
                print '%d line is writting into Excel'%(2*ii+iii)
                        
        time.sleep(1)
        '第%d项测试已全部保存'%(i)
        print 'the testing of %d ele is finished,starting the next ele'%(i)        
    
    time.sleep(1)
    workbook.close()
#     '退出登录'
#     arg=importinfo('logout') 
#     time.sleep(3)
#     logoutbtton = WebDriverWait(b, 10).until(lambda c: b.find_element_by_xpath(arg['smcurl_logouid'])) 
#     ActionChains(b).click(logoutbtton).perform()
#     time.sleep(1)
#     logoutbtton2 =b.find_element_by_xpath(arg['smcurl_logouid2'])
#     ActionChains(b).move_to_element(logoutbtton2).click().perform()
#     print '退出系统登录'
    
    print '**************************all the ele is finished，stop the test!!!************************************'
    
    
if __name__ == '__main__':
#     b = webdriver.Chrome()
#     b=webdriver.Firefox()
    b=webdriver.PhantomJS(executable_path='/home/meng/Downloads/env/phantomjs-2.1.1-linux-x86_64/bin/phantomjs')
#     b=webdriver.PhantomJS(executable_path='/root/Downloads/phantomjs-2.1.1-linux-x86_64/bin/phantomjs')
    b = logoin(b)
    inquire(b)