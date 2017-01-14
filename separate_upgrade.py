#coding:utf-8
'''Created on 2017年1月4日
@author: MengLei'''

from basic import *
from mail_smtp import mail
from excel_column_perspective import excel_perspective

class upgrade():
    def __init__(self,b,shengjicelue_yemianyuansu,shengji_input):
        self.b=b
        self.yemianyuansu=shengjicelue_yemianyuansu
        self.shengji_input=shengji_input
        
    def addtask(self):
        time.sleep(1)
        self.b.find_element_by_xpath(self.yemianyuansu['ruanjianxitongshengji']).click()
        time.sleep(1)
        
        '开始添加升级'
        self.b.find_element_by_xpath(self.yemianyuansu['tianjia']).click()
        time.sleep(2)
        self.b.find_element_by_xpath(self.yemianyuansu['xinjianrenwumingcheng_input']).send_keys(self.shengji_input['renwumingcheng'])
        time.sleep(0.5)
        self.b.find_element_by_xpath(self.yemianyuansu['yunyingshang_input']).send_keys(self.shengji_input['yunyingshang'])
        time.sleep(0.5)
        self.b.find_element_by_xpath(self.yemianyuansu['yunyingshang_chaxun_input']).click()
        time.sleep(0.5)
#         self.b.find_element_by_xpath(self.yemianyuansu['yunyingshang_click']).click()
        time.sleep(0.5)
        self.b.find_element_by_xpath(self.yemianyuansu['jizhan_input']).send_keys(self.shengji_input['jizhanliebiao'])
        time.sleep(0.5)
        self.b.find_element_by_xpath(self.yemianyuansu['jizhan_chaxun']).click()
        time.sleep(0.5)
        self.b.find_element_by_xpath(self.yemianyuansu['jizhan_click']).click()
        time.sleep(0.5)
        self.b.find_element_by_xpath(self.yemianyuansu['jizhan_xuanding']).click()
        time.sleep(0.5)
        self.b.find_element_by_xpath(self.yemianyuansu['yixuanjizhan_click']).click()
        time.sleep(0.5)
        self.b.find_element_by_xpath(self.yemianyuansu['xiayibu']).click()
        time.sleep(1)
        
    def selectversion_new(self):        
        '选择要升级版本'
        try:
            print 'target new version:',shengji_input['mubiaobanben_new']
            self.b.find_element_by_xpath('//span[@title="%s"]'%self.shengji_input['mubiaobanben_new']).click()
        except:
            try:
                self.b.find_element_by_xpath(self.yemianyuansu['xiayiye_click']).click()
                time.sleep(1)
                self.b.find_element_by_xpath('//span[@title="%s"]'%self.shengji_input['mubiaobanben']).click()
            except:
                pass
        time.sleep(2)
        self.b.find_element_by_xpath(self.yemianyuansu['wancheng']).click()
         
        '取消升级'
#         self.b.find_element_by_xpath('//*[@id="step_2"]/div/div[2]/div/a[1]/span/span').click()            
            
    def selectversion_old(self):        
        '选择要升级版本'
        try:
            print 'target old version:',shengji_input['mubiaobanben_old']
            self.b.find_element_by_xpath('//span[@title="%s"]'%self.shengji_input['mubiaobanben_old']).click()
        except:
            try:
                self.b.find_element_by_xpath(self.yemianyuansu['xiayiye_click']).click()
                time.sleep(1)
                self.b.find_element_by_xpath('//span[@title="%s"]'%self.shengji_input['mubiaobanben']).click()
            except:
                pass            
        self.b.find_element_by_xpath(self.yemianyuansu['wancheng']).click()
         
        '取消升级'
#         self.b.find_element_by_xpath('//*[@id="step_2"]/div/div[2]/div/a[1]/span/span').click()
            
    def checkresult(self):
        '确认升级'
        
        '等待升级结果'
        time.sleep(1)
        WebDriverWait(b,100).until(lambda x: self.b.find_element_by_xpath('//div[contains(text(),"meng01")]'))
        
        '查找结果'
        for i in range(3):
            aa=self.b.find_element_by_xpath('//*[@id="datagrid-row-r4-2-%d"]/td[2]/div'%i)
            if aa.text == 'meng01':
                break
         
        '判断是否升级完成'
        while True:
            bb=self.b.find_element_by_xpath('//*[@id="datagrid-row-r4-2-%d"]/td[6]/div'%i).text
            
            try:
                '升级前 执行结果栏 _进度状态的xpath：row-r13'
                Task_Results =  self.b.find_element_by_xpath('//*[@id="datagrid-row-r13-2-%d"]/td[3]/div'%i).text
            except:
                try:
                    '升级后 执行结果栏 _进度状态的xpath：row-r14 (加了1)'
                    Task_Results =  self.b.find_element_by_xpath('//*[@id="datagrid-row-r14-2-%d"]/td[3]/div'%i).text
                except:
                    pass
      
            if bb == u'已结束' or bb == 'End':
                print Task_Results.decode('utf-8'),'upgrade finished...'
                break
            else:
                print Task_Results.decode('utf-8')
                time.sleep(20)
                
                 
        '判断升级是否成功'
        time.sleep(1)
        Upgrade_result=self.b.find_element_by_xpath('//*[@id="datagrid-row-r4-2-%d"]/td[7]/div'%i).text
        time.sleep(0.2)
        Task_Results =  self.b.find_element_by_xpath('//*[@id="datagrid-row-r14-2-%d"]/td[3]/div'%i).text
        
        print Upgrade_result
        print Task_Results
        if Upgrade_result==u'成功' or  Upgrade_result == 'upgrade complete':
            print 'upgrade success...'
            pass
        else:
            print 'upgrade failed,resean:',Upgrade_result
        time.sleep(1)
        
        return Upgrade_result,Task_Results


    def delete(self):
        time.sleep(1)
        self.b.find_element_by_xpath(self.yemianyuansu['ruanjianxitongshengji']).click()
        time.sleep(1)
        '查找结果'
        for ii in range(10):
            time.sleep(0.5)
            print 'starting check result'
            try:
                '查找结果id-row-r15？？？实际自动化连续操作时，确实是这样的'
                ee=self.b.find_element_by_xpath('//*[@id="datagrid-row-r15-2-%d"]/td[2]/div'%ii)
            except:
                pass
                                             
            if ee.text == 'meng01':
                break
        time.sleep(1)
        
        '删除结果'
        try:
            '删除结果id-row-r15？？？实际自动化连续操作时，确实是这样的'
            self.b.find_element_by_xpath('//*[@id="datagrid-row-r15-2-%d"]/td[10]/div/div[4]'%ii).click()
        except:
            pass
                   
                                        
        '确定删除'
        time.sleep(1)
        try:
            '确定删除div[58]，实际自动化连续操作时，确实是这样的'
            lll=b.find_element_by_xpath('/html/body/div[60]/div[3]/a[1]/span/span')
        except:
            pass                
        time.sleep(0.5)
        lll.click()
        
        
        '取消删除'
        #=======================================================================
        # time.sleep(1)
        # try:
        #     print 'div[58]'
        #     b.find_element_by_xpath('/html/body/div[58]/div[3]/a[2]/span/span').click()
        # except:
        #     pass                
        #=======================================================================
        
        
        


if __name__ == '__main__':
    '打开浏览器'
    browser=browser()
    b=browser.run()
    
    '各种实例化'
    login=login(b)
    excel=CreateExcel()
    ini=importinfo()
    workbook,worksheet1,worksheet2,excel_path=excel.run('upgrade')
    opencatalogue=opencatalogue()
    logout=logout()
    worksheet1.write(0,0,'Upgrade_result')
    worksheet1.write(0,1,'shengjijizhan')
    worksheet1.write(0,2,'target version')
    worksheet1.write(0,3,'start time')
    worksheet1.write(0,4,'finish time')
    worksheet1.write(0,5,'Task_Progress(failed reason)')
    
    '登录循环'
#     for iii in range(1,101):
    for iii in range(1,2):
        '登录'
        login.run()
         
        '导入左侧目录xpath'
        zuocemulu_xpath=ini.run('web_ele.ini','zuocemulu')
        time.sleep(0.5)
        catalogue_xpath=zuocemulu_xpath.pop('jizhan')
        son_catalogue_xpath=zuocemulu_xpath.pop('jizhan_celueguanli')
        son_son_catalogue_xpath=zuocemulu_xpath.pop('jizhan_celueguanli_shengjicelue')
        
        
        '打开左侧目录'
        opencatalogue.run(b,catalogue_xpath,son_catalogue_xpath)
        time.sleep(0.5)
        opencatalogue.run(b,son_catalogue_xpath,son_son_catalogue_xpath)
        
        
        shengjicelue_yemianyuansu=ini.run('web_ele.ini','shengjicelue_yemianyuansu')
        shengji_input = ini.run('web_ele.ini','shengji')
        '升级'
        up=upgrade(b,shengjicelue_yemianyuansu,shengji_input)
        up.addtask()
        
        '第一次升级'
        if iii % 2:
            up.selectversion_new()
            worksheet1.write(iii,2,shengji_input['mubiaobanben_new'])
        else:
            up.selectversion_old()
            worksheet1.write(iii,2,shengji_input['mubiaobanben_old'])
            
            
        '记录起始时间+执行结果'
        worksheet1.write(iii,1,shengji_input['jizhanliebiao'])
        worksheet1.write(iii,3,time.strftime('%Y-%m-%d %H:%M', time.localtime(time.time())))
        Upgrade_result,Task_Results = up.checkresult()
        worksheet1.write(iii,0,Upgrade_result)
        worksheet1.write(iii,4,time.strftime('%Y-%m-%d %H:%M', time.localtime(time.time())))
        worksheet1.write(iii,5,Task_Results)
        
        '删除结果'
        up.delete()
         
        '退出登录'
        logout.run(b)
        time.sleep(60)
        
        
    workbook.close()
    print 'over...'
    
    excel_perspective=excel_perspective()
    excel_perspective.run(excel_path)
    
    mail=mail()
    '打印一下Excel文件的具体名称'
#     print ('report'+os.sep+'perspective_'+str(excel_path.split(os.sep)[-1:]).split('\'')[1])
    mail.fasong('report'+os.sep+'perspective_'+str(excel_path.split(os.sep)[-1:]).split('\'')[1])        
    
    
    