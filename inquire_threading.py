#coding:utf-8
'''Created on 2016年12月6日
@author: MengLei'''

from basic import *
from mail_smtp import mail
from excel_column_perspective import excel_perspective
import threading
from test.test_threading_local import target

class inquire():
    def __init__(self,b,yemianyuansu,workbook,worksheet1,worksheet2):
        self.b=b
        self.yemianyuansu=yemianyuansu
        self.workbook=workbook
        self.worksheet1=worksheet1
        self.worksheet2=worksheet2

    def run(self,ini,text,sections):
        '1. 导入要查询的元素'
        list=ini.run(text,sections)

        '2. 循环打开下拉列表，选择各个元素，并保存查询结果'
        nrows = 0
        for i in range(1,len(list)+1):
#         for i in range(1,2):
            sepacial_flag =0

            '1).打开下拉列表，选择查找元素，并点击执行'
            print '***************************start testing...      total %d      this is %d*******************************'%(i,len(list)),list['%d'%i]
            self.b.find_element_by_xpath(self.yemianyuansu['select']).click()
            time.sleep(1)
            
            #===================================================================
            # '截图看一下效果'
            # self.b.save_screenshot('screenshot.png')
            # print 'screenshot.png over...'            
            #===================================================================
            
            pp =self.b.find_element_by_xpath('//div[contains(text(),"%s")]'%list['%d'%i])
            time.sleep(1)
            ActionChains(self.b).move_to_element(pp).double_click().perform()
            time.sleep(1)
            #===================================================================
            # '截图看一下效果'
            # self.b.save_screenshot('screenshot1.png')
            # print 'screenshot1.png over...'            
            #===================================================================

            self.b.find_element_by_xpath(self.yemianyuansu['zhixing']).click()
            
            '写入开始时间'
            self.worksheet1.write(i,3,time.strftime('%Y-%m-%d %H:%M', time.localtime(time.time())))
            time.sleep(1)
            nline=0
             
             
             
            '2).获得1、2行的查询结果，并保存到Excel表格'
            '//*[@id="showParamValues"]/div[元素]/span[行]'
            for m in range(1,3):
                f='//*[@id="showParamValues"]/div[%d]/span[%d]'%(i,m)
                try:
                    #===================================================================
                    # '截图看一下效果'
                    # self.b.save_screenshot('screenshot1.png')
                    # print 'screenshot1.png over...'            
                    #===================================================================
                    g=WebDriverWait(self.b,20).until(lambda x: self.b.find_element_by_xpath(f)).text
                    self.worksheet2.write(nrows,0,g)
                    nrows+=1
                    print '%d line is writting into Excel'%m
                    time.sleep(0.2)
                except:
                    self.worksheet2.write(nrows, 0,'等待结果超时，本元素终止查询')
                    nrows+=1
                    '等待结果超时，本项查询终止，进行下一项查询！'
                    print 'time out... end this ele . starting the next ele'
                    
            '写入具体完成时间'
            self.worksheet1.write(i,4,time.strftime('%Y-%m-%d %H:%M', time.localtime(time.time())))
            
            '3).获得所有索引（3、4、5、6、7、8...行）的结果'
            '(1).判断有多少个索引'   
            for nindex in range(1,22):
                f='//*[@id="showParamValues"]/div[%d]/div[%d]'%(i,nindex)
                try:
                    self.b.find_element_by_xpath(f)
                    nline=nindex
                except:
                    if nline>0:
                        break
                    else:
                        nline=1
            '查询结果一共包含%d个索引'%(nline)
                        
            '(2).判断索引有几列*******************'
            all_col=[]
            for n in range(1,22):
                f='//*[@id="showParamValues"]/div[%d]/div/div[%d]'%(i,n)
                try:
                    findele=self.b.find_element_by_xpath(f)
                    all_col.append(findele)
                except:
                    if len(all_col)>=1:
                        '证明是有结果，可能只有1列，则不在进行查找'
                        break
                    else:
                        '没有返回正确结果，强制定为只有1列'
#                         print 'failed result. all_col must=1 all_col.append('')'
                        all_col.append('')
            'all_col有%d列,具体元素为：'%(len(all_col)),all_col
            
            '(3).获取内容导出到Excel*******************'
            '''i项目，ii,nline索引，iii行数（共2行）,n列数  '''
            for ii in range(1,nline+1):
                for iii in range(1,3):
                    for n in range(1,len(all_col)+1):
         
                        '1.获取第3和第4行的正常结果，具体数值'
                        fff='//*[@id="showParamValues"]/div[%d]/div[%d]/div[%d]/span[%d]'%(i,ii,n,iii)
                        'm为：%d，查询元素的具体xpath为： %s'%(iii,f)
                        try:
                            g = self.b.find_element_by_xpath(fff).text
                            self.worksheet2.write(nrows,n-1,g)
                        except:
                            '2.获取第3和第4行的异常结果获取,，写入excel'
                            if ii==1 and iii==1 and n ==1:
                                'iii==1(第三行) and n ==1（第一列）'
                                f='//*[@id="showParamValues"]/div[%d]/span[3]'%(i)
                                try:
                                    g = self.b.find_element_by_xpath(f).text
                                    self.worksheet2.write(nrows,n-1,g)
                                    sepacial_flag+=1
                                except:
                                    pass                                
                            elif  ii==1 and iii==2 and n==1:
                                'iii==2（第四行） and n==1（第一列）'
                                f='//*[@id="showParamValues"]/div[%d]/span[4]'%(i)
                                try:
                                    g = self.b.find_element_by_xpath(f).text
                                    self.worksheet2.write(nrows,n-1,g)
                                    sepacial_flag+=1
                                except:
                                    pass                                
                            else:
                                pass
                            
                    nrows+=1
                    time.sleep(0.2)
                    '一行内容填写完成'
                    print '%d line is writting into Excel'%(2*ii+iii)
                             
            time.sleep(1)
            '第%d项测试已全部保存'%(i)
            
            if sepacial_flag==0:
                
                self.worksheet1.write(i,0,'pass.')
            else:
                self.worksheet1.write(i,0,'failed')
            
            '写入基站名称'
            self.worksheet1.write(i,1,self.yemianyuansu['jizhan_kuandai'])
            '写入具体查询元素'
            self.worksheet1.write(i,2,list['%d'%i])

            


            print 'start the next...'
        time.sleep(1)
        self.workbook.close()
        print '***************************test finished!   close workbook...'

if __name__ == '__main__':
    def inquire():
        '各种实例化'
        browser=browser()
        excel=CreateExcel()
        ini=importinfo()
        opencatalogue=opencatalogue()
        
        '登录'
        b=browser.run()
        login=login(b)
        login.run()
        workbook,worksheet1,worksheet2,excel_path=excel.run('inquire')
        worksheet1.write(0,0,'Upgrade_result')
        worksheet1.write(0,1,'jijizhan')
        worksheet1.write(0,2,'LST_parameter')
        worksheet1.write(0,3,'start time')
        worksheet1.write(0,4,'finish time')
        
        try:
            b.find_element_by_xpath('/html/body/div[39]/div[3]/a/span/span').click()
        except:
            pass
        
        '导入左侧目录xpath'
        zuocemulu_xpath=ini.run('web_ele.ini','zuocemulu')
        time.sleep(0.5)
        catalogue_xpath=zuocemulu_xpath.pop('jizhan')
        son_catalogue_xpath=zuocemulu_xpath.pop('jizhan_peizhi')   
        
        '打开左侧目录'
        time.sleep(0.5)
        opencatalogue.run(b,catalogue_xpath,son_catalogue_xpath)
        time.sleep(0.5)
        yemianyuansu=ini.run('web_ele.ini', 'yemianyuansu')
        
        '选定基站执行'
        time.sleep(1)
        ddd =WebDriverWait(b,20).until(lambda x: b.find_element_by_xpath(yemianyuansu['chaxunshuru']))
        ddd.send_keys(yemianyuansu['jizhan_kuandai'])
        time.sleep(1)
        b.find_element_by_xpath(yemianyuansu['chaxunanniu']).click()
        time.sleep(1)
        b.find_element_by_xpath(yemianyuansu['xuandinganniu']).click()
        
        '开始查询，并保存结果'
        session2=inquire(b, yemianyuansu,workbook, worksheet1,worksheet2)
        session2.run(ini,'web_ele.ini','chaxun')
    
        '透视'
        excel_perspective=excel_perspective()
        excel_perspective.run(excel_path)
        
        '邮件发送'
        mail=mail()
        '打印一下Excel文件的具体路径+名称'
    #     print ('report'+os.sep+'perspective_'+str(excel_path.split(os.sep)[-1:]).split('\'')[1])
        mail.fasong('report'+os.sep+'perspective_'+str(excel_path.split(os.sep)[-1:]).split('\'')[1])
    th=[threading.Thread(target=inquire) for i in range(3)]
    for t in th:
        t.start()
    
    
    
    
    
    