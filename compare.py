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
from omc_inquire import *
import os


'''
Created on 2016年9月29日
@author: MengLei
'''
def compare(excel1,excel2):
    '创建文件'
    if  os.path.exists(os.getcwd()+os.sep+'report'+os.sep+u'基站信息修改结果与查询结果对比'+datetime.now().date().isoformat()+'.xlsx'):
        excelnum=0
        while 1:
            excelnum +=1
            #===============================================================
            # '目前不限制每日新建EXCEL文件数量，暂时注释掉'
            # if excelnum >20:
            #     break
            #===============================================================
            if not os.path.exists(os.getcwd()+os.sep+'report'+os.sep+u'基站信息修改结果与查询结果对比'+datetime.now().date().isoformat()+u'号'+str(excelnum)+'.xlsx'):
                workbook = xlsxwriter.Workbook(os.getcwd()+os.sep+'report'+os.sep+u'基站信息修改结果与查询结果对比'+datetime.now().date().isoformat()+u'号'+str(excelnum)+'.xlsx')
                worksheet1 = workbook.add_worksheet()
                print '222'
                print u'生成Excel文件为： '+os.getcwd()+os.sep+'report'+os.sep+ u'基站信息修改结果与查询结果对比'+datetime.now().date().isoformat()+u'号'+str(excelnum)+'.xlsx'
                #===========================================================
                # '目前不需要读取模块，暂时注释掉'
                # readexcel = xlrd.open_workbook(u'基站信息查询'+datetime.now().date().isoformat()+u'号'+str(excelnum)+'.xlsx')
                #===========================================================
                break
    else:
        workbook = xlsxwriter.Workbook(os.getcwd()+os.sep+'report'+os.sep+u'基站信息修改结果与查询结果对比'+datetime.now().date().isoformat()+'.xlsx')
        worksheet1 = workbook.add_worksheet()
        print'333'
        print u'生成Excel文件为： '+os.getcwd()+os.sep+'report'+os.sep+u'基站信息修改结果与查询结果对比'+datetime.now().date().isoformat()+'.xlsx'    
      
    
    
    
    '读取文件，并对比'

    
    excelflag=0
    readexcel1=xlrd.open_workbook(excel1)
    readsheet1=readexcel1.sheets()[0]
    shortrows=readsheet1.nrows
    
    readexcel2=xlrd.open_workbook(excel2)
    readsheel2=readexcel2.sheets()[0]
    longrows=readsheel2.nrows
#     print '第一个表格有%d行'%rows1
#     print '第二个表格有%d行'%rows2
    nrows=0
    if shortrows>longrows:
        readexcel1=xlrd.open_workbook(excel2)
        readsheet1=readexcel1.sheets()[0]
        shortrows=readsheet1.nrows
        
        readexcel2=xlrd.open_workbook(excel1)
        readsheel2=readexcel2.sheets()[0]
        longrows=readsheel2.nrows
        excelflag=1
        
    print  'shortrows',shortrows
    for n in range(0,shortrows):
        print  'shortrows的N',n
        a_equal_b=0
        con1=[]
        con2=[]
        context1=readsheet1.row_values(n)
        context2=readsheel2.row_values(n)
        print 'context11',context1
        print 'context22',context2



        for i in context1:
            if i =='':
                pass
            else:
                con1.append(i)
        context1=con1
  
        for ii in context2:
            if ii =='':
                pass
            else:
                con2.append(ii)
        context2=con2                
        print 'context1',context1
        print 'context2',context2
        
        if context1==context2:
            print'第%d行，相等，pass'%n
            pass
        else:
            '修改第一行和 n%4==0的context'
            if n==0 or n%4==0:
                print 'n==0 or n4==0,当前是第%d行'%n
                context1='LST' + str(con1)[6:-1]
                context2=str(con2)[3:-1]
                print context1
                print context2
                
                '修改完毕后再次 判断是否相等'
                if  context1 == context2:
                    pass
                    print  '修改完毕后 pass'
                else:
                    print '第一个表格第%d对比了第二个表格%d行之后的所有数据'%(n,n)
                    nn=0
                    
                    for nn in range(longrows):
                        con2=[]
                        context2=readsheel2.row_values(nn)
                        for ii in context2:
                            if ii =='':
                                pass
                            else:
                                con2.append(ii)
                        context2=str(con2)[3:-1]
                        print 'context2 is %s,第%d行'%(context2,nn)

                        if context1 == context2:
                            a_equal_b=1
                            #===================================================
                            # if excelflag==0:
                            #     worksheet1.write(nrows,0,excel1 + u'第%d行与'%(n+1)+ excel2 +u'第%d行相同'%(nn+1))
                            # else:
                            #     worksheet1.write(nrows,0,excel1 + u'第%d行与'%(n+1)+ excel2 +u'第%d行相同'%(nn+1))
                            #===================================================
                            break
                        
                    if nn == longrows-1 and a_equal_b==0:
                        if excelflag==0:
                            worksheet1.write(nrows,0,excel1 + u'第%d行'%(n+1)+'not found')
                        else:
                            worksheet1.write(nrows,0,excel1 + u'第%d行'%(n+1)+'not found')                            
            
            else:
                nn=0
                print 'not 第0行，也not 4的倍数（新的一个查询项，需要替换LST为MOD）,当前是第%d行'%n
                print 'context1 is ',context1
                for nn in range(longrows):
                    print 'not 4==0的context，第%d个元素的%d次'%(n,nn)
                    con2=[]
                    context2=readsheel2.row_values(nn)
                    for ii in context2:
                        if ii =='':
                            pass
                        else:
                            con2.append(ii)
                    context2=con2
                    print 'context2 is ',context2

                    if context1 == context2:
                        a_equal_b=1
                        #=======================================================
                        # if excelflag==0:
                        #     worksheet1.write(nrows,0,excel1 + u'第%d行与'%(n+1)+ excel2 +u'第%d行相同'%(nn+1))
                        # else:
                        #     worksheet1.write(nrows,0,excel1 + u'第%d行与'%(n+1)+ excel2 +u'第%d行相同'%(nn+1))
                        #=======================================================
                        break
                    
                if nn == longrows-1 and a_equal_b==0:
                    if excelflag==0:
                        worksheet1.write(nrows,0,excel1 + u'第%d行'%(n+1)+'not found')
                    else:
                        worksheet1.write(nrows,0,excel1 + u'第%d行'%(n+1)+'not found')                            

            
        print  'nrows+1'
        nrows+=1
                
    workbook.close()
    print '测试结束！！！！！'
    



    
if __name__ == '__main__':
#     xx=u'基站信息修改结果'+datetime.now().date().isoformat()+u'号'+'.xlsx'
#     xxx=u'基站信息修改结果'+datetime.now().date().isoformat()+u'号'+'1'+'.xlsx'
#     excel1='MOD_RESULT_2-'+str(datetime.now().date().isoformat())+'-'+'1.xlsx'
#     excel2='MOD_RESULT_2-'+str(datetime.now().date().isoformat())+'.xlsx'
#     print excel1
#     print excel2
    excel1='LST_RESULT-2016-09-30-1.xlsx'
    excel2='MOD_RESULT_2-2016-09-30.xlsx'
    
    compare(excel1,excel2)