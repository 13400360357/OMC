#coding:utf-8

from selenium import webdriver
from datetime import  *
import time,os,xlsxwriter,xlrd,ConfigParser
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains 
from selenium.webdriver.common.keys import Keys 

'''
# Created on 2016年9月29日
@author: MengLei
'''            

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
    
    def run(self,catalogue_xpath,son_catalogue_xpath):
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
            logoutbtton = WebDriverWait(b, 3).until(lambda c: b.find_element_by_xpath(son_catalogue_xpath)) 
            logoutbtton.click()

class CreateExcel:
    def run(self):
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
        
        return workbook,worksheet1
          
        '*******************************************************windows系统*******************************************************'
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
        # return workbook,worksheet1
        # '********************************************************Linux系统********************************************************'
        #===========================================================================

class CreateTxt:
    
    #===========================================================================
    # '********************************************************Linux系统********************************************************'
    # def run(self):
    #     '4.按当天日期创建txt文件'
    #     if  os.path.exists('/home/meng/workspace/OMC/report/err'+datetime.now().date().isoformat()+'.txt'):
    #         excelnum=0
    #         while 1:
    #             excelnum +=1
    #             if not os.path.exists('/home/meng/workspace/OMC/report/err'+datetime.now().date().isoformat()+'-'+str(excelnum)+'.txt'):
    #                 txt = open('/home/meng/workspace/OMC/report/err'+datetime.now().date().isoformat()+'-'+str(excelnum)+'.txt','yemianyuansu')
    #                 print 'Create txt:/home/meng/workspace/OMC/report/err'+datetime.now().date().isoformat()+'-'+str(excelnum)+'.txt'
    #                 break
    #     else:
    #         txt = open('/home/meng/workspace/OMC/report/err'+datetime.now().date().isoformat()+'.txt','yemianyuansu')
    #         print 'Create txt:/home/meng/workspace/OMC/report/err'+datetime.now().date().isoformat()+'.txt'
    #     return txt
    #     '********************************************************Linux系统********************************************************'
    #===========================================================================
    
    
    '*******************************************************windows系统*******************************************************'    
    def run(self):
        '4.按当天日期创建txt文件'
        if  os.path.exists(os.getcwd()+os.sep+'report'+os.sep+'err'+datetime.now().date().isoformat()+'.txt'):
            excelnum=0
            while 1:
                excelnum +=1
                if not os.path.exists(os.getcwd()+os.sep+'report'+os.sep+'err'+datetime.now().date().isoformat()+'-'+str(excelnum)+'.txt'):
                    txt = open(os.getcwd()+os.sep+'report'+os.sep+'err'+datetime.now().date().isoformat()+'-'+str(excelnum)+'.txt','yemianyuansu')
                    print 'Create txt： '+os.getcwd()+os.sep+'report'+os.sep+ 'err'+datetime.now().date().isoformat()+'-'+str(excelnum)+'.txt'
                    break
        else:
            txt = open(os.getcwd()+os.sep+'report'+os.sep+'err'+datetime.now().date().isoformat()+'.txt','yemianyuansu')
            print 'Create txt： '+os.getcwd()+os.sep+'report'+os.sep+'err'+datetime.now().date().isoformat()+'.txt'
        return txt
        '*******************************************************windows系统*******************************************************'

class login:
    def __init__(self):
        self.b = webdriver.Chrome()
#         self.b=webdriver.Firefox()
#         self.b=webdriver.PhantomJS(executable_path='/home/meng/Downloads/env/phantomjs-2.1.1-linux-x86_64/bin/phantomjs')
#         self.b=webdriver.PhantomJS(executable_path='/root/Downloads/phantomjs-2.1.1-linux-x86_64/bin/phantomjs')
    def run(self):
        b=self.b
        time.sleep(1)
        b.maximize_window()
        im=importinfo()
        inquire=im.run('web_ele.ini','login')
        b.get(inquire['url'])
        d=WebDriverWait(b, 10).until(lambda c:b.find_element_by_xpath(inquire['smcurl_userid']))
        d.clear()
        d.send_keys(inquire['username'])
        time.sleep(0.5)
        e=b.find_element_by_xpath(inquire['smcurl_pwdid'])
        e.clear()
        e.send_keys(inquire['userpassword'])
        b.find_element_by_xpath(inquire['smcurl_loginid']).click()
        return b

class prepare:
    def login(self):
        c=login()
        b=c.run()
        return b

    def opencatalogue(self,txt, sections, catalogue_option, son_catalogue_option):
        '2.2 打开左侧目录'
        open=opencatalogue()
        ini,catalogue_xpath,son_catalogue_xpath=open.importinfo(txt,sections,catalogue_option,son_catalogue_option)
        open.run(catalogue_xpath,son_catalogue_xpath)
        time.sleep(0.5)
        return ini
        
    def import_yemianyuansu(self,b,ini):
        '3.1 导入基站'
        yemianyuansu=ini.run('web_ele.ini','yemianyuansu')
        return yemianyuansu
    
    def select_enb(self,yemianyuansu):
        '3.2 选定，执行'
        time.sleep(1)
        ddd =WebDriverWait(b,20).until(lambda x: b.find_element_by_xpath(yemianyuansu['chaxunshuru']))
        ddd.send_keys(yemianyuansu['jizhan_kuandai'])
        time.sleep(1)
        b.find_element_by_xpath(yemianyuansu['chaxunanniu']).click()
        time.sleep(1)
        b.find_element_by_xpath(yemianyuansu['xuandinganniu']).click()
                 
    def excel(self):
        '4. 创建Excel'
        excel=CreateExcel()
        workbook,worksheet1=excel.run()
        return workbook,worksheet1
    
    def txt(self):
        txt=CreateTxt()
        txt.run()
        return txt
 
class inquire:
    def __init__(self,b,yemianyuansu,workbook,worksheet1):
        self.b=b
        self.yemianyuansu=yemianyuansu
        self.workbook=workbook
        self.worksheet1=worksheet1
        

    def run(self,ini,text,sections):
        '1. 导入要查询的元素'
        items=ini.run(text,sections)
        list=(items.values())
        
        '2. 循环打开下拉列表，选择各个元素，并保存查询结果'
        nrows = 0
#         for i in range(1,len(list)+1):
        for i in range(1,2):
            '1).打开下拉列表，选择查找元素，并点击执行'
            print '***************************start testing...      total %d      this is %d*******************************'%(i,len(list))
            aa=self.b.find_element_by_xpath(self.yemianyuansu['select'])
            time.sleep(1)
    #         ActionChains(self.b).move_to_element(aa).double_click().perform()
            aa.click()
            time.sleep(1)
            print 'ele is:',list[(i-1)]
            pp = b.find_element_by_xpath(list[(i-1)])
            ActionChains(self.b).double_click(pp).perform()
            self.b.find_element_by_xpath(self.yemianyuansu['zhixing']).click()
            time.sleep(1)
            nline=0
             
            '2).获得1、2行的查询结果，并保存到Excel表格'
            '/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[元素]/div[索引]/div[列]/span[行]'
            for m in range(1,3):
                '1)).获取第1和2行的结果（1查询项、2基站ID），并保存Excel'
                f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/span[%d]'%(i,m)
                try:
                    g=WebDriverWait(self.b,300).until(lambda x: self.b.find_element_by_xpath(f)).text
                    self.worksheet1.write(nrows,0,g)
                    nrows+=1
                except:
                    self.worksheet1.write(nrows, 0,'等待结果超时，本元素终止查询')
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
                    self.b.find_element_by_xpath(f)
                nline=1
                        
            except:
                '正常返回值，判断多少个索引'
                try:
                    for nlin in range(1,22):
                        f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/div[%d]'%(i,nlin)
                        self.b.find_element_by_xpath(f)
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
                    findele=self.b.find_element_by_xpath(f)
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
                            findele=self.b.find_element_by_xpath(f)
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
                            g = self.b.find_element_by_xpath(fff).text
                            self.worksheet1.write(nrows,n-1,g)
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
                                g = self.b.find_element_by_xpath(f).text
                                self.worksheet1.write(nrows,n-1,g)
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
        self.workbook.close()
        print '***************************test finished!   close workbook...'


class modify:
    def run(self,b,yemianyuansu,modifyobjectelepath,modifyobjectxpath,txt,workbook,worksheet1):
        '1 导入元素'
        modify_imp=importinfo()
        modifyobject=modify_imp.run('web_ele.ini','xiugai')
        modifyobjectxpath=modifyobject.values()
        modifyobjectelepath=modifyobject.keys()
        
        '2.开始修改功能测试'
        nrows = 0
        nround=1
        nline=0
        commad_success_flag = 1
        'commad_success_flag是判断点击执行后，是否成功下发。'
    #     for round in range(1,len(modifyobjectxpath)+1):
        for round in range(1,2):
            select=b.find_element_by_xpath(yemianyuansu['select'])
            ActionChains(b).move_to_element(select).double_click().perform()
            time.sleep(1)
            print '**************************** start test %d ************** total %d **************'%(round,len(modifyobjectxpath)),modifyobjectxpath[round-1]
            round_modify = b.find_element_by_xpath(modifyobjectxpath[round-1])
            ActionChains(b).move_to_element(round_modify).double_click().perform()
              
            print 'inquire element   {quantity | type(input or select) | necessary or not...}'
            
            '1)确定有多少要被修改的元素,下一步确定是input或者select'          
            all_modify_ele=[]
            for aaa in range(20):
                    f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div[2]/div/form/ul/li[%d]'%(aaa)
                    try:
                        findele=b.find_element_by_xpath(f)
                        all_modify_ele.append(findele)
                    except:
                        continue
    #         print '%d input or select'%(len(all_modify_ele))              
    #         print '有%d个要修改的元素,具体元素为：'%(len(all_modify_ele)),all_modify_ele
      
              
            '2)确定输入还是下拉'
            input_ele=[]
            select_ele=[]
            input_path=[]
            select_path=[]
            for modifyobject in range(1,len(all_modify_ele)+1):
                ele='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div[2]/div/form/ul/li[%d]/input'%(modifyobject)
                ele1='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div[2]/div/form/ul/li[%d]/select'%(modifyobject)
                try:
                    inputele=b.find_element_by_xpath(ele)
                    input_ele.append(inputele)
                    input_path.append(ele)
                except:
                    selectele=b.find_element_by_xpath(ele1)
                    select_ele.append(selectele)
                    select_path.append(ele1)
                    
    #         print'%d input'%(len(input_ele))
    #         print'%d select'%(len(select_ele))                
    #         print'有%d个输入框,下拉列表是:'%(len(input_ele)),input_ele
    #         print'有%d个下拉框,输入列表是:'%(len(select_ele)),select_ele
    #         print'有%d个输入框,下拉列表的xpath是:'%(len(input_ele)),input_path
    #         print'有%d个下拉框,输入列表的xpath是:'%(len(select_ele)),select_path
                  
    
            '3)确定有多少个必选项目,下一步判断多少个输入/下拉必选项'
            necessary_xpath=[]
            for modifyobject in range(1,len(all_modify_ele)+1):
                ele='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div[2]/div/form/ul/li[%d]/div'%(modifyobject)
                try:
                    necetele=b.find_element_by_xpath(ele)
                    if '*' in necetele.text:
                        necessary_xpath.append(ele)
                except:
                    continue
    #         print'%d necessary'%(len(necessary_xpath))
    #         print'有%d个必选项目,必选项目是'%(len(necessary_xpath)),necessary_xpath 
              
              
            '4)判断下拉或者输入框中，那些是必须项'
            select_necessary_xpath=[]
            select_no_necessary_xpath=[]
            input_necessary_xpath=[]
            inpout_no_necessary_xpath=[]
            if len(necessary_xpath)!=0:
                for m in range(len(necessary_xpath)):
                    '这是第%d个必选，判断是否为input。相应路径为：'%(m+1),(necessary_xpath[m][:-3]+'input')
                    for i in range(len(input_ele)):
    #                     print '第%d个input框，xpath为:'%(i),input_path[i]
                        if input_path[i] in (necessary_xpath[m][:-3]+'input'):
                            input_necessary_xpath.append(input_path[i])
    #                         print  '第%d个inputpath框，是必输入框'%(i)
                        else:
                            inpout_no_necessary_xpath.append(input_path[i])
    #                         print  '第%d个inputpath框，not必输入框:'%(i)
                               
                    '这是第%d个必选，判断是否为select。相应路径为：'%(m+1), (necessary_xpath[m][:-3]+'select')
                    for i in range(len(select_path)):
    #                     print '第%d个select框，xpath为:'%(i)+select_path[i]
                        if select_path[i] in (necessary_xpath[m][:-3]+'select'):
                            select_necessary_xpath.append(select_path[i])
    #                         print '第%d个select框，是必输入框:'%(i)
                        else:
                            select_no_necessary_xpath.append(select_path[i])
    #                         print '第%d个select框，not必输入框:'%(i)
                            
    #         print '%d necessary iput '%(len(input_necessary_xpath))
    #         print '%d necessary select'%(len(select_necessary_xpath))          
    #         print '输入框必选有%d个，具体为：'%(len(input_necessary_xpath)),input_necessary_xpath
    #         print '下拉框必选有%d个，具体为：'%(len(select_necessary_xpath)),select_necessary_xpath
    #         print '输入框not必选有%d个，具体为：'%(len(inpout_no_necessary_xpath)),inpout_no_necessary_xpath
    #         print '下拉框not必选有%d个，具体为：'%(len(select_no_necessary_xpath)),select_no_necessary_xpath          
              
               
            '***********************************************************判断、执行******************************************************************************************'
            '''
            1.全部不填写，执行，并查看结果
            2.所有input部分全部填写超出范围，判断是否提示。 
            3.input数值正常，界面遗漏1项必填内容（输入框/下拉列表），执行并查看执行结果。循环遍历所有必填项目。
            4.正常填写，执行并查看结果
            '''
            
            
            print 'start execute  {no member | illegal character | absent necessary | max value...}'                 
            
            '1.全部不填写，执行'
            b.find_element_by_xpath(yemianyuansu['zhixing']).click()
            time.sleep(1)
            '判断是否有弹窗'
            try:
                n = WebDriverWait(b, 3).until(lambda x: b.find_element_by_xpath('/html/body/div[54]/div[3]/yemianyuansu/span/span')).click()
                'nothing be wirtten. pass'
            except:
                print'nothing be wirtten. fault！！'
                txt.write('\n%d ele:%s,nothing be wirtten，no messagebox, failure。'%(round,modifyobjectelepath[round-1]))
                
                  
            '2所有input部分全部填写超出范围(特殊字符)，判断是否提示。 '
            time.sleep(5)
            if len(input_ele)>0:
                for num_outinput in range(1,len(input_ele)+1):
                    b.find_element_by_xpath(input_path[num_outinput-1]).clear()
                    b.find_element_by_xpath(input_path[num_outinput-1]).send_keys('&&&&%%%%%$$(@#$%^!@#$%^&*(#$%^&*($%^&*()@#$%^&*()_#$%^&*()$%^&*@#$%^&*()#$%^&*($%^&*&*()!@#$%^&*(@#$%^&*(@#$%^&*@#$%^&')
                    time.sleep(0.5)
                    b.find_element_by_xpath(yemianyuansu['zhixing']).click()
                    checkinp= input_path[num_outinput-1][:-5]+'div'
                    time.sleep(1)
                      
                    '2.1判断是否有弹窗，如果有，点击确定'
                    try:
                        n = WebDriverWait(b, 1).until(lambda x: b.find_element_by_xpath('/html/body/div[54]/div[3]/yemianyuansu/span/span')).click()
                        txt.write('\n%d ele:%s, %d inputbox，special context，messagebox,failure'%(round,modifyobjectelepath[round-1],num_outinput))
                    except:
                        'special context，no messagebox,pass'
                        pass
                       
                    '2.2判断是否有超出范围的提示'
                    try:
                        result=b.find_element_by_xpath(checkinp)
                    except:
                        print '%d input, special context,no prompt，fault'%(num_outinput)
                        txt.write('\n%d ele:(%s),%d inputbox，special context,no prompt，failure'%(round,modifyobjectelepath[round-1],num_outinput))
                    finally:
                        b.find_element_by_xpath(input_path[num_outinput-1]).clear()
                             
                    '2.3命令是否下发成功'
                    try:     
                        WebDriverWait(b,10).until(lambda x: b.find_element_by_xpath('/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/span[1]'%(commad_success_flag)))
                        commad_success_flag+=1
                        print 'all inputbox writting special context,click execute and it really happen.  %d ele of special context failure'%(num_outinput)
                        txt.write('\n special context，,click execute and it really happen. %d ele of special context failure'%(num_outinput))
                    except:
                        'all inputbox writting special context,pass'
                        pass
            else:
    #             print '2.no inputbox,ignore...'
                pass
              
              
            '***************************************************************判断、执行******************************************************************************************'
              
            '3.遗漏必填项，进行测试。--先正常填写（最大边界值），然后挨个删除必填项目，执行并查看结果'
            inputkey = ConfigParser.ConfigParser() 
            inputkey.read("web_input.ini") 
      
            '3.1进行输入框的填写'
            for num_input in range(1,len(input_ele)+1):
                input_ele[num_input-1].clear()
                input_ele[num_input-1].send_keys(inputkey.get(modifyobjectxpath[round-1][:-1],str(num_input)))
              
            '3.2进行下拉列表的选择'
            if len(select_ele)>0:
                for num_select in range(1,len(select_ele)+1):
                    ActionChains(b).move_to_element(select_ele[num_select-1]).click().perform()
                    time.sleep(1)
                    n_select=WebDriverWait(b,1).until(lambda x: b.find_element_by_xpath(inputkey.get(modifyobjectxpath[round-1][:-1],'select'+str(num_select))))
                    ActionChains(b).move_to_element(n_select).double_click().perform()
                    if len(select_ele)==1:
                        ActionChains(b).move_to_element(select_ele[num_select-1]).click().perform()
                        ActionChains(b).move_to_element(select_ele[num_select-1]).click().perform()
                          
            '3.3遗漏必选项'
            if len(necessary_xpath)>0:
                
                '3.3.1遗漏必选输入框，进行下发'
                if len(input_necessary_xpath)>0 :
                    for num_input in range(1,len(input_ele)+1):  
                        if input_path[num_input-1] in input_necessary_xpath:
                            input_ele[num_input-1].clear()
                            b.find_element_by_xpath(yemianyuansu['zhixing']).click()
                            time.sleep(3)
                            try:
                                WebDriverWait(b,5).until(lambda x: b.find_element_by_xpath('/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/span[1]'%(commad_success_flag)))
                                commad_success_flag+=1
    #                             print '遗漏必填项，input框部分，命令下发成功，用例测试失败'
                                txt.write('\n%d ele:(%s),%d necessary_inputbox，no context and excute success. failure'%(round,modifyobjectelepath[round-1],num_input))
                            except:
                                pass
    #                             print '遗漏必填项，input框部分，命令未下发，用例通过'                      
                            input_ele[num_input-1].send_keys(inputkey.get(modifyobjectxpath[round-1][:-1],str(num_input)))
                    time.sleep(0.5)
    #                 print '3.1遗漏必遗漏必选输入框，测试完成'
                
                else:
                    pass
    #                 print '3.1没有必选输入框，跳过遗漏'
                
                '3.3.2遗漏必选下拉框，进行下发'              
                if len(select_necessary_xpath)>0:
                    for num_select in range(1,len(select_ele)+1):
                        if (inputkey.get(modifyobjectxpath[round-1][:-1],'select'+str(num_select)))[:-10] in select_necessary_xpath:
                            ActionChains(b).move_to_element(select_ele[num_select-1]).click().perform()
                            time.sleep(1)
                            n_select=WebDriverWait(b,10).until(lambda x: b.find_element_by_xpath(inputkey.get(modifyobjectxpath[round-1][:-1],'select'+ str(num_select))[:-3]+'[1]'))
                            ActionChains(b).move_to_element(n_select).double_click().perform()
                            b.find_element_by_xpath(yemianyuansu['zhixing']).click()
                            time.sleep(3)
                              
                            try:
                                WebDriverWait(b,5).until(lambda x: b.find_element_by_xpath('/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/span[1]'%(commad_success_flag)))
                                commad_success_flag+=1
                                print '%d select，necessary_inputbox，no context and excute success. failure'%(num_select)
                                txt.write('\n%d ele:(%s), %d select，necessary_inputbox，no context and excute success. failure'%(round,modifyobjectelepath[round-1],num_select))
                            except:
                                pass
    #                             print '遗漏必填项，进行测试。select框部分,第%d个input部分填写超出范围用例，点击执行后命令未下发，用例通过'%(num_select)                        
      
                            '遗漏必填项，进行测试,重新填回遗漏部分'
                            ActionChains(b).move_to_element(select_ele[num_select-1]).click().perform()
                            time.sleep(1)
                            n_select=WebDriverWait(b,10).until(lambda x: b.find_element_by_xpath(inputkey.get(modifyobjectxpath[round-1][:-1],'select'+ str(num_select))))
                            ActionChains(b).move_to_element(n_select).double_click().perform()
    #                     print '3.2遗漏必遗漏必选下拉框，测试完成'
                else:
                    pass
    #                 print '3.2没有必选下拉框，跳过遗漏'                
            
            else:
                pass
    #             print '3.无必选项，跳过遗漏必选项测试' 
            
            '4.正常填写（最大边界值），执行并查看结果（接着3的结果，直接点击执行，判断即可）'
            b.find_element_by_xpath(yemianyuansu['zhixing']).click()
            time.sleep(1)
            
            '4.1判断是否有弹窗'
            try:
                n = WebDriverWait(b, 1).until(lambda x: b.find_element_by_xpath('/html/body/div[54]/div[3]/yemianyuansu/span/span')).click()
                txt.write('\n%d ele:%s，max value，messagebox . failure'%(round,modifyobjectelepath[round-1]))
                print 'max value,click on execute,pop-up window. failure'
            except:
                pass
              
            '4.2最大值，点击执行后，命令是否下发成功'
            try:
                WebDriverWait(b,300).until(lambda x: b.find_element_by_xpath('//*[@id="showParamValues"]/div[%d]/div'%(commad_success_flag)))
                commad_success_flag+=1
    #             print 'reslut get,prepare for writting...'
    #             print '填写最大边界值，命令下发成功。修改功能测试结束。即将进行结果读取并保存到Excel'
            except:
                print 'max value，after execute，no reslut.failure'%(round)    
                txt.write('\n %d ele:%s,max value，after execute，no reslut.failure'%(round,modifyobjectelepath[round-1]))
            
                           
      
               
            '5.读取执行结果，并写入Excel文件'
            
            '5.1读取前两行结果'
            for m in range(1,3):
                col = 0
                all_col=[]
                f='//*[@id="showParamValues"]/div[%d]/span[%d]'%(round,m)
                try:
                    g=WebDriverWait(b,300).until(lambda x: b.find_element_by_xpath(f)).text
                    worksheet1.write(nrows,0,g)
                    print '%d line writting into Excel'%(m)        
                except:
                    worksheet1.write(nrows, 0,'time out for reading  result')
                    print 'time out for reading  result'
                    continue
                nrows+=1
                   
                   
            '5.2获取3、4、5、6...的内容'
            
            '5.2.1判断多少个索引（行）'
            try:
    
                for nlin in range(1,22):
                    f='//*[@id="showParamValues"]/div[%d]/div[%d]'%(round,nlin)
                    b.find_element_by_xpath(f)
                    nline=nlin
            except:
                try:
                    '3）结果返回错误值，将索引数（行数）直接定义为1行'
                    for nlin in range(1,3):
                        f='//*[@id="showParamValues"]/div[%d]/span[%d]'%(round,nlin+2)
                        nline=1
                except:
                    print 'result incorrect,failure '
                    txt.write('\n result incorrect,failure ')
                        
            '%d个索引(1个索引有2行).'%(nline*2+2)
                      
                          
            '5.2.2判断有几列,后续打印具体细节'
            for n in range(1,22):
                f='//*[@id="showParamValues"]/div[%d]/div/div[%d]'%(round,n)
                try:
                    findele=b.find_element_by_xpath(f)
                    all_col.append(findele)
                except:
                    if len(all_col)>=1:
                        '证明正确的返回结果，不在进行循环'
                        break
                    else:
                        '查找错误的返回结果'
                        if m==3 and n ==1:
                            f='//*[@id="showParamValues"]/div[%d]/span[3]'%(round)
                        try:
                            findele=b.find_element_by_xpath(f)
                            all_col.append(findele)
                        except:
                            pass
    #         print '索引行有%d列'%(len(all_col))                    
    #         print 'all_col有%d列,具体元素为：'%(len(all_col)),all_col
                   
                          
          
            '5.2.3结果查询，并导出*********************************'
            
            '''round(项目)
            ii(索引)
            iii(行，1为第3行，2为第4行)
            n(列)'''
            for ii in range(1,nline+1):
                for iii in range(1,3):
                    for n in range(1,len(all_col)+1):
                        f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/div[%d]/div[%d]/span[%d]'%(round,ii,n,iii)
                        '1).获取第3和第4行的正常结果，具体数值'
                        try:
                            g = b.find_element_by_xpath(f).text
                            worksheet1.write(nrows,n-1,g)
    #                         print 'result is writting %d line，%d column'%(nrows,n-1)
                        except:
        
                            '2).获取第3和第4行的异常结果获取,，写入excel'
                            if ii==1 and iii==1 and n ==1:
                                'iii==1(第三行) and n ==1（第一列）'
                                f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/span[3]'%(round)
                            elif  ii==1 and iii==2 and n==1:
                                'iii==2（第四行） and n==1（第一列）'
                                f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/span[4]'%(round)
                            else:
                                f=''
        #                     print '查找错误的返回结果,当前的%d行，第%d个元素，xpath为%s'%(iii,n,f)
                            try:
                                g = b.find_element_by_xpath(f).text
                                worksheet1.write(nrows,n-1,g)
                            except:
        #                     print '*****异常结果数据写入********%d行，%d列'%(nrows,n-1)
                                break
                                        
                    '一行内容填写完成'
                    nrows+=1
                    time.sleep(0.2)
                    print '%d line writting into Excel'%(2*ii+iii)
    
            
            'start the next element......'
            
            
        time.sleep(1)
        workbook.close()


class omc_add:
#     def __init__(self,b,yemianyuansu,modifyobjectelepath,modifyobjectxpath,txt,workbook,worksheet1):
#         self.b=b
#         self.yemianyuansu=yemianyuansu
#         self.modifyobjectelepath=modifyobjectelepath
#         self.modifyobjectxpath=modifyobjectxpath
#         self.workbook=workbook
#         self.worksheet1=worksheet1  
#         self.txt=txt
#         
        
    def run(self):
        '3 导入元素'
        ini=ConfigParser.ConfigParser()
        ini.read('add.ini')
        elelist=ini.items('LST')
        print elelist
        dict={}
        dict.update(elelist)
        
        '4. 执行LST'
        nresult=0
        for e in range(1,len(dict)+1):
            '**************************************************执行LST******************************************'      
            
            '1.进行下拉框点击' 
            aa=b.find_element_by_xpath('/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div[1]/div/span/span/yemianyuansu')
            ActionChains(b).move_to_element(aa).double_click().perform()
            time.sleep(1)
            print '******************************************************************************这是第%d个元素，一共%d个。'%(e,len(dict))
            print '本次查询/增/删项目为:',dict['%d'%e]
            print '开始第一次查询，获得初始结果'
            if dict['%d'%e]!="LST IPSEC":
        #       round_modifyobjectxpath = b.find_element_by_xpath('//div[contains(text(),"LST CELL")]')
                round_modifyobjectxpath = b.find_element_by_xpath('//div[contains(text(),"%s")]'%dict['%d'%e]) 
            else:
                round_modifyobjectxpath = b.find_element_by_xpath('//div[contains(text(),"%s ")]'%dict['%d'%e])             
            ActionChains(b).move_to_element(round_modifyobjectxpath).double_click().perform()
            b.find_element_by_link_text('执行').click()
            nresult+=1
            time.sleep(1)
        
            '2 获取结果'
            cunzai=[]
            queshao=[]
            WebDriverWait(b, 60).until(lambda x:b.find_element_by_xpath('/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]'%nresult)).text
            
    #         print '1第一次查询结果/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]'%nresult
            
            
            for i in range(1,25):
                try:
    #                     print '第%d元素,获得初始查询结果的第%d行'%(e,i)
                    num=b.find_element_by_xpath('/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/div[%d]/div[1]/span[2]'%(nresult,i)).text
                    
                    cunzai.append(int(num))
                except:
                    pass
    
            cunzai=insert_sort(cunzai)
            print '初始结果：目前已存在的索引列表',cunzai
        
    
    
            '**************************************************执行ADD******************************************'          
            '1 导入ADD元素'   
            print  '执行添加功能测试，具体操作为：ADD%s'%(dict['%d'%e][3:])
            add_ele=ini.items('ADD%s'%(dict['%d'%e][3:]))
            dict_add={}
            dict_add.update(add_ele)
            
            '2.进行下拉框点击' 
            aa=b.find_element_by_xpath('/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div[1]/div/span/span/yemianyuansu')
            ActionChains(b).move_to_element(aa).double_click().perform()
            time.sleep(1)
            
    #       round_modifyobjectxpath = b.find_element_by_xpath('//div[contains(text(),"LST CELL")]')
            round_modifyobjectxpath = b.find_element_by_xpath('//div[contains(text(),"%s")]'%('ADD%s'%(dict['%d'%e][3:])))
            ActionChains(b).move_to_element(round_modifyobjectxpath).double_click().perform()
            time.sleep(1)  
            
             
            '3确定网页有多少要被修改的元素'          
            all_modify_ele=[]
            '确定有多少要被修改的元素,这里只给出数值，下一步确定是input或者select'
            for aaa in range(22):
    #             print '确定有多少要被修改的输入框/选择框'
                if aaa ==1:
                    try:
                        f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div[2]/div/form/ul/li[%d]'%(aaa)
                        findele=b.find_element_by_xpath(f)
                        all_modify_ele.append(findele)
                    except:
                        f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div[2]/div/form/ul/li'
                        findele1=b.find_element_by_xpath(f)
                        all_modify_ele.append(findele1)
    #                     print '只有1个输入框/选择框'
                        break
                else:
    #                 print '不是只有一个输入框/选择框，循环查找'
                    f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div[2]/div/form/ul/li[%d]'%(aaa)
                    try:
                        findele2=b.find_element_by_xpath(f)
                        all_modify_ele.append(findele2)
                    except:
                        continue
    #         print '有%d个要修改的元素,具体元素为：'%(len(all_modify_ele)),all_modify_ele
      
              
            '4确定输入还是下拉'          
            input_ele=[]
            select_ele=[]
            input_path=[]
            select_path=[]
            for modifyobject in range(1,len(all_modify_ele)+1):
    #             print'确定第%s个元素是输入还是下拉'%(modifyobject)
                if  len(all_modify_ele)!=1:
                    ele='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div[2]/div/form/ul/li[%d]/input'%(modifyobject)
                    ele1='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div[2]/div/form/ul/li[%d]/select'%(modifyobject)
                else:
                    ele='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div[2]/div/form/ul/li/input'
                    ele1='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div[2]/div/form/ul/li/select'
                try:
                    inputele=b.find_element_by_xpath(ele)
                    input_ele.append(inputele)
                    input_path.append(ele)
                except:
                    selectele=b.find_element_by_xpath(ele1)
                    select_ele.append(selectele)
                    select_path.append(ele1)
    #         print'有%d个输入框,下拉列表是:'%(len(input_ele)),input_ele
    #         print'有%d个下拉框,输入列表是:'%(len(select_ele)),select_ele
    #         print'有%d个输入框,下拉列表的xpath是:'%(len(input_ele)),input_path
    #         print'有%d个下拉框,输入列表的xpath是:'%(len(select_ele)),select_path
    
    
            '5进行填写'          
    #         print '进行输入框的填写'
            for num_input in range(1,len(input_ele)+1):
                input_ele[num_input-1].clear()
                input_ele[num_input-1].send_keys(dict_add['%d'%num_input])
              
    #         print '进行下拉列表的选择'
            if len(select_ele)>0:
                for num_select in range(1,len(select_ele)+1):
                    ActionChains(b).move_to_element(select_ele[num_select-1]).click().perform()
                    time.sleep(1)
                    n_select=WebDriverWait(b,1).until(lambda x: b.find_element_by_xpath(dict_add['select%d'%num_select]))
                    ActionChains(b).move_to_element(n_select).double_click().perform()
                    
                    '脚本进行选择会造成某些必选框无法正常识别，因此每个选择框选择两次，下面脚本重复选择一次'
                    ActionChains(b).move_to_element(select_ele[num_select-1]).click().perform()
                    time.sleep(1)
                    n_select=WebDriverWait(b,1).until(lambda x: b.find_element_by_xpath(dict_add['select%d'%num_select]))
                    ActionChains(b).move_to_element(n_select).double_click().perform()
                    
                    if len(select_ele)==1:
                        ActionChains(b).move_to_element(select_ele[num_select-1]).click().perform()
                        ActionChains(b).move_to_element(select_ele[num_select-1]).click().perform()            
                
                b.find_element_by_link_text('执行').click()
                nresult+=1
                time.sleep(1)
                
                WebDriverWait(b, 60).until(lambda x:b.find_element_by_xpath('/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]'%nresult)).text
    #             print '2 进行添加操作后/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]'%nresult
    
    
            '**********************再次查询，和第一次结果进行对比*********************************'          
        
            '1 再次执行LST'   
            print  '开始第二次查询，确定添加的索引号'
            aa=b.find_element_by_xpath('/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div[1]/div/span/span/yemianyuansu')
            ActionChains(b).move_to_element(aa).double_click().perform()
            time.sleep(1)
            if dict['%d'%e]!="LST IPSEC":
        #       round_modifyobjectxpath = b.find_element_by_xpath('//div[contains(text(),"LST CELL")]')
                round_modifyobjectxpath = b.find_element_by_xpath('//div[contains(text(),"%s")]'%dict['%d'%e]) 
            else:
                round_modifyobjectxpath = b.find_element_by_xpath('//div[contains(text(),"%s ")]'%dict['%d'%e])             
            ActionChains(b).move_to_element(round_modifyobjectxpath).double_click().perform()
            b.find_element_by_link_text('执行').click()
            nresult+=1
            time.sleep(1)
            
    
            WebDriverWait(b, 60).until(lambda x:b.find_element_by_xpath('/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]'%nresult)).text
    #         print '3 第二次查询（确定添加的索引）/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]'%nresult
            
            '2 获取结果'
            cunzai2=[]
            queshao2=[]
            for i in range(1,25):
                try:
    #                     print '第%d元素,第二次查询（确定添加的索引）,获得查询结果的第%d行'%(e,i)
                    num=b.find_element_by_xpath('/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/div[%d]/div[1]/span[2]'%(nresult,i)).text
                    cunzai2.append(int(num))
                except:
                    pass
        
            cunzai2=insert_sort(cunzai2)
            print '经过添加操作后，目前存在的索引为：',cunzai2
            
            
            '3 对比添加前后的结果'
            add_success=[]
            for i in cunzai2:
                if i in cunzai:
                    pass
                else:
                    add_success.append(i)
                    
            if len(add_success)!=1:
                print '添加操作执行失败，请人工检查操作前后序列号变化，跳过删除操作，直接进行下一项操作'
            else:
            
                
                
                '**********************执行删除操作*********************************'      
                
                '1.进行下拉框点击' 
                aa=b.find_element_by_xpath('/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div[1]/div/span/span/yemianyuansu')
                ActionChains(b).move_to_element(aa).double_click().perform()
                time.sleep(1)
                
                
                round_modifyobjectxpath = b.find_element_by_xpath('//div[contains(text(),"%s")]'%('RMV%s'%(dict['%d'%e][3:])))
                ActionChains(b).move_to_element(round_modifyobjectxpath).double_click().perform()
                time.sleep(1)  
                
                '2.空格框内输入序列号'
                print '执行删除功能，删除索引号(增加的索引号)为:',add_success[0]
                b.find_element_by_xpath('//*[@id="16"]').send_keys(add_success[0])
                time.sleep(1)
                b.find_element_by_link_text('执行').click()
                nresult+=1
                time.sleep(1)
                
                WebDriverWait(b, 60).until(lambda x:b.find_element_by_xpath('/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]'%nresult)).text
    #             print '4删除操作 /html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]'%nresult
    
                '**********************再次查询，和第一次结果进行对比*********************************'          
            
                '1 再次执行LST'   
                print  '开始第3次查询，确定删除后的索引号'
                
                
                '1.进行下拉框点击' 
                aa=b.find_element_by_xpath('/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div[1]/div/span/span/yemianyuansu')
                ActionChains(b).move_to_element(aa).double_click().perform()
                time.sleep(1)
                if dict['%d'%e]!="LST IPSEC":
                    # round_modifyobjectxpath = b.find_element_by_xpath('//div[contains(text(),"LST CELL")]')
                    round_modifyobjectxpath = b.find_element_by_xpath('//div[contains(text(),"%s")]'%dict['%d'%e]) 
                else:
                    round_modifyobjectxpath = b.find_element_by_xpath('//div[contains(text(),"%s ")]'%dict['%d'%e])             
                ActionChains(b).move_to_element(round_modifyobjectxpath).double_click().perform()
                b.find_element_by_link_text('执行').click()
                nresult+=1
                time.sleep(1)
                
            
                WebDriverWait(b, 60).until(lambda x:b.find_element_by_xpath('/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]'%nresult)).text
        #         print '5查询删除后结果 /html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]'%nresult
        
                '2 获取结果'
                cunzai3=[]
                queshao3=[]
                for i in range(1,25):
                    try:
    #                     print '第%d元素,删除后进行查询，目前提取第%d行'%(e,i)
                        num=b.find_element_by_xpath('/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/div[%d]/div[1]/span[2]'%(nresult,i)).text
                        cunzai3.append(int(num))
                    except:
                        pass
    
                cunzai3=insert_sort(cunzai3)
                print '删除完毕后目前存在的索引号为：',cunzai3
            
                for i in cunzai3:
                    if i in cunzai:
                        pass
                    else:
                        print '经添加和删除功能后，序列发生改变，请检查。'
                print '本项测试完成，即将进行下一项测试'
            

if __name__ == '__main__':
    session=prepare()
    b=session.login()
    try:
        b.find_element_by_xpath('/html/body/div[39]/div[3]/a/span/span').click()
    except:
        pass
    
    ini=session.opencatalogue('web_ele.ini','zuocemulu', 'jizhan', 'jizhan_peizhi')
    yemianyuansu=session.import_yemianyuansu(b, ini)
    session.select_enb(yemianyuansu)
    workbook,worksheet1=session.excel()

    session2=inquire(b, yemianyuansu,workbook, worksheet1)
    session2.run(ini,'web_ele.ini','inquire')



    