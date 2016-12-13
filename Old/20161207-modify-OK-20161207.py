#coding:utf-8

# import tkMessageBox
from selenium import webdriver
from datetime import  *
import time
from selenium.webdriver.common.action_chains import ActionChains 
from selenium.webdriver.common.keys import Keys 
from selenium.webdriver.support.ui import WebDriverWait
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
    item=cp.items(arg)
    args.update(item)
    return args
  
def logoin(b):
    '''登录网页'''
    arg=importinfo('login')   
    time.sleep(0.5)
    b.get(arg['url'])
    b.maximize_window()
    time.sleep(1)
    name=WebDriverWait(b, 10).until(lambda c:b.find_element_by_xpath(arg['smcurl_userid']))
    name.clear()
    name.send_keys(arg['username'])
    time.sleep(0.5)
    passw=WebDriverWait(b, 10).until(lambda c:b.find_element_by_xpath(arg['smcurl_pwdid']))
    passw.clear()
    passw.send_keys(arg['userpassword'])
    b.find_element_by_xpath(arg['smcurl_loginid']).click()
    return b
  
  
  
def opencatalogue(b,mydirectory,myrootdirectory):
    '''判断并打开网页子目录'''
    try:
        time.sleep(1)
        subtree = WebDriverWait(b, 15).until(lambda c: b.find_element_by_xpath(mydirectory))
        subtree.click()
    except:
        tree=b.find_element_by_xpath(myrootdirectory)
        time.sleep(1)
        ActionChains(b).move_to_element(tree).click().perform()
        time.sleep(1)
        subtree = WebDriverWait(b, 15).until(lambda c: b.find_element_by_xpath(mydirectory)) 
        subtree.click()
          

    
def opentreepath(b):
    '打开左侧目录'
    arg=importinfo('jibenpeizhi')
    time.sleep(1)
    myrootdirectory=arg.pop('jizhan')
    mydirectory=arg.pop('jizhanpeizhi')
    opencatalogue(b,mydirectory,myrootdirectory)
    time.sleep(0.5)


       
def modify(b):
    '1.打开目录，导入元素'
    opentreepath(b)
    webele=importinfo('yemianyuansu')
    modifyobject=importinfo('xiugai')
    modifyobjectxpath=modifyobject.values()
    modifyobjectelepath=modifyobject.keys()
#     print '修改项目的Xpath',modifyobjectxpath
#     print '修改项目的名称',modifyobjectelepath
  
    '2填入基站ID，并选定，执行'
    time.sleep(1)
    NBinput =WebDriverWait(b,20).until(lambda x: b.find_element_by_xpath(webele['chaxunshuru']))
    NBinput.send_keys(webele['jizhan_yidong'])
    time.sleep(1)
    b.find_element_by_xpath(webele['chaxunanniu']).click()
    time.sleep(1)
    b.find_element_by_xpath(webele['xuandinganniu']).click()
    
    
    '********************************************************Linux系统********************************************************'
    #===========================================================================
    # '3.按当天日期创建excel文件'
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
    #     
    # '4.按当天日期创建txt文件'
    # if  os.path.exists('/home/meng/workspace/OMC/report/err'+datetime.now().date().isoformat()+'.txt'):
    #     excelnum=0
    #     while 1:
    #         excelnum +=1
    #         if not os.path.exists('/home/meng/workspace/OMC/report/err'+datetime.now().date().isoformat()+'-'+str(excelnum)+'.txt'):
    #             txt = open('/home/meng/workspace/OMC/report/err'+datetime.now().date().isoformat()+'-'+str(excelnum)+'.txt','a')
    #             print 'Create txt:/home/meng/workspace/OMC/report/err'+datetime.now().date().isoformat()+'-'+str(excelnum)+'.txt'
    #             break
    # else:
    #     txt = open('/home/meng/workspace/OMC/report/err'+datetime.now().date().isoformat()+'.txt','a')
    #     print 'Create txt:/home/meng/workspace/OMC/report/err'+datetime.now().date().isoformat()+'.txt'
    #===========================================================================
    '********************************************************Linux系统********************************************************'
    
    
    '*******************************************************windows系统*******************************************************'
    '3.按当天日期创建excel文件'
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
      
      
    '4.按当天日期创建txt文件'
    if  os.path.exists(os.getcwd()+os.sep+'report'+os.sep+'err'+datetime.now().date().isoformat()+'.txt'):
        excelnum=0
        while 1:
            excelnum +=1
            if not os.path.exists(os.getcwd()+os.sep+'report'+os.sep+'err'+datetime.now().date().isoformat()+'-'+str(excelnum)+'.txt'):
                txt = open(os.getcwd()+os.sep+'report'+os.sep+'err'+datetime.now().date().isoformat()+'-'+str(excelnum)+'.txt','a')
                print 'Create txt： '+os.getcwd()+os.sep+'report'+os.sep+ 'err'+datetime.now().date().isoformat()+'-'+str(excelnum)+'.txt'
                break
    else:
        txt = open(os.getcwd()+os.sep+'report'+os.sep+'err'+datetime.now().date().isoformat()+'.txt','a')
        print 'Create txt： '+os.getcwd()+os.sep+'report'+os.sep+'err'+datetime.now().date().isoformat()+'.txt'
    '*******************************************************windows系统*******************************************************'
      
      
      
    '5.开始修改功能测试'
    nround=1
    nline=0
    commad_success_flag = 1
    'commad_success_flag是判断点击执行后，是否成功下发。'
#     for round in range(1,len(modifyobjectxpath)+1):
    for round in range(1,2):
        select=b.find_element_by_xpath(webele['select'])
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
        b.find_element_by_xpath(webele['zhixing']).click()
        time.sleep(1)
        '判断是否有弹窗'
        try:
            n = WebDriverWait(b, 3).until(lambda x: b.find_element_by_xpath('/html/body/div[54]/div[3]/a/span/span')).click()
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
                b.find_element_by_xpath(webele['zhixing']).click()
                checkinp= input_path[num_outinput-1][:-5]+'div'
                time.sleep(1)
                  
                '2.1判断是否有弹窗，如果有，点击确定'
                try:
                    n = WebDriverWait(b, 1).until(lambda x: b.find_element_by_xpath('/html/body/div[54]/div[3]/a/span/span')).click()
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
                        b.find_element_by_xpath(webele['zhixing']).click()
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
                        b.find_element_by_xpath(webele['zhixing']).click()
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
        b.find_element_by_xpath(webele['zhixing']).click()
        time.sleep(1)
        
        '4.1判断是否有弹窗'
        try:
            n = WebDriverWait(b, 1).until(lambda x: b.find_element_by_xpath('/html/body/div[54]/div[3]/a/span/span')).click()
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

#     '退出登录'
#     arg=importinfo('logout') 
#     time.sleep(3)
#     logoutbtton = WebDriverWait(b, 10).until(lambda c: b.find_element_by_xpath(arg['smcurl_logouid'])) 
#     ActionChains(b).click(logoutbtton).perform()
#     time.sleep(1)
#     logoutbtton2 =b.find_element_by_xpath(arg['smcurl_logouid2'])
#     ActionChains(b).move_to_element(logoutbtton2).click().perform()
#     print '退出系统登录'
    
    print '****************************all over ,close****************************'

    
    
if __name__ == '__main__':
    b = webdriver.Chrome()
#     b=webdriver.PhantomJS(executable_path='/home/meng/Downloads/env/phantomjs-2.1.1-linux-x86_64/bin/phantomjs')
    b = logoin(b)
    modify(b)
