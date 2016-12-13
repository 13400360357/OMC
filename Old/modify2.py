#coding:utf-8
import tkMessageBox
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
    cp.read(r'D:\OMC\web_ele.ini')
    args={}
    print arg.decode('utf-8').encode('gbk')
    item=cp.items(arg.decode('utf-8').encode('gbk'))
    args.update(item)
    return args
  
def logoin(b):
    '''登录网页'''
    arg=importinfo('login')   
    time.sleep(0.5)
    b.get(arg['url'])
    b.maximize_window()
    time.sleep(1)
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



def modify2(b):
    '1.打开目录，导入元素'
    opentreepath(b)
    webele=importinfo('页面元素')
    modifyobject=importinfo('修改')
    modifyobjectxpath=modifyobject.values()
    modifyobjectelepath=modifyobject.keys()
    print 'modifyobjectxpath',modifyobjectxpath
    print  'modifyobjectelepath',modifyobjectelepath


    '2填入基站ID，并选定，执行'
    ddd =b.find_element_by_xpath(webele[('查询输入').decode('utf-8').encode('gbk')])
    ddd.send_keys(webele[('基站_移动').decode('utf-8').encode('gbk')])
    time.sleep(1)
    b.find_element_by_xpath(webele[('查询按钮').decode('utf-8').encode('gbk')]).click()
    time.sleep(1)
    b.find_element_by_xpath(webele[('选定按钮').decode('utf-8').encode('gbk')]).click()
      
      
    '3查询界面输入框和必选项目'
    '****************************************************************查找网页元素***************************************************************************'          
    '1).按当天日期创建excel文件'
    nrows = 0
    command_success_result=1
    if  os.path.exists('MOD_RESULT_2-'+datetime.now().date().isoformat()+'.xlsx'):
        print '111'
        excelnum=0
        while 1:
            excelnum +=1
            #===============================================================
            # '目前不限制每日新建EXCEL文件数量，暂时注释掉'
            # if excelnum >20:
            #     break
            #===============================================================
            if not os.path.exists('MOD_RESULT_2-'+datetime.now().date().isoformat()+'-'+str(excelnum)+'.xlsx'):
                workbook = xlsxwriter.Workbook('MOD_RESULT_2-'+datetime.now().date().isoformat()+'-'+str(excelnum)+'.xlsx')
                worksheet1 = workbook.add_worksheet()
                print '222'
                print u'生成Excel文件为： '+ 'MOD_RESULT_2-'+datetime.now().date().isoformat()+'-'+str(excelnum)+'.xlsx'
                #===========================================================
                # '目前不需要读取模块，暂时注释掉'
                # readexcel = xlrd.open_workbook(u'基站信息查询'+datetime.now().date().isoformat()+u'号'+str(excelnum)+'.xlsx')
                #===========================================================
                break
    else:
        workbook = xlsxwriter.Workbook('MOD_RESULT_2-'+datetime.now().date().isoformat()+'.xlsx')
        worksheet1 = workbook.add_worksheet()
        print'333'
        print u'生成Excel文件为： '+'MOD_RESULT_2-'+datetime.now().date().isoformat()+'.xlsx'    


      
      
    print '*************************开始项修改功能测试***********************'
    nround=1
    nline=0
    commad_success_flag = 1
    'commad_success_flag是判断点击执行后，是否成功下发。'
    '2)开始总循环，导入查询的项，并点击跳转到具体界面'
    for round in range(1,len(modifyobjectxpath)+1):
#     for round in range(1,2):
        aa=b.find_element_by_xpath(webele['select'])
        ActionChains(b).move_to_element(aa).double_click().perform()
        time.sleep(1)
        print '***************************************************************这是第%d个元素，一共%d个。本次查询元素为'%(round,len(modifyobjectxpath)),modifyobjectxpath[round-1]
        round_modifyobjectxpath = b.find_element_by_xpath(modifyobjectxpath[round-1])
        ActionChains(b).move_to_element(round_modifyobjectxpath).double_click().perform()
          
          
        '**********************查找网页元素*********************************'          
        all_modify_ele=[]
        '确定有多少要被修改的元素,这里只给出数值，下一步确定是input或者select'
        for aaa in range(20):
            print '确定有多少要被修改的元素'
            if aaa ==1:
                try:
                    f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div[2]/div/form/ul/li[%d]'%(aaa)
                    findele=b.find_element_by_xpath(f)
                    all_modify_ele.append(findele)
                except:
                    f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div[2]/div/form/ul/li'
                    findele1=b.find_element_by_xpath(f)
                    all_modify_ele.append(findele1)
                    print '只有1个元素，后续不需要再判断元素数目'
                    break
            else:
                print '不是只有一个元素，循环查找'
                f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div[2]/div/form/ul/li[%d]'%(aaa)
                try:
                    findele2=b.find_element_by_xpath(f)
                    all_modify_ele.append(findele2)
                except:
                    continue
        print '有%d个要修改的元素,具体元素为：'%(len(all_modify_ele)),all_modify_ele
  
          
        '**********************查找网页元素*********************************'          
        input_ele=[]
        select_ele=[]
        input_path=[]
        select_path=[]
        '确定输入还是下拉'
        for modifyobject in range(1,len(all_modify_ele)+1):
            print'确定第%s个元素是输入还是下拉'%(modifyobject)
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
        print'有%d个输入框,下拉列表是:'%(len(input_ele)),input_ele
        print'有%d个下拉框,输入列表是:'%(len(select_ele)),select_ele
        print'有%d个输入框,下拉列表的xpath是:'%(len(input_ele)),input_path
        print'有%d个下拉框,输入列表的xpath是:'%(len(select_ele)),select_path
              
        '**********************查找网页元素*********************************'          
        necessary_xpath=[]
        '确定有多少个必选项目'
        for modifyobject in range(1,len(all_modify_ele)+1):
            print'确定有多少个必选项目' 
  
            if  len(all_modify_ele)!=1:
                ele='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div[2]/div/form/ul/li[%d]/div'%(modifyobject)
            else:
                ele='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div[2]/div/form/ul/li/div'
            try:
                necetele=b.find_element_by_xpath(ele)
                if '*' in necetele.text:
                    necessary_xpath.append(ele)
            except:
                continue
        print'有%d个必选项目,必选项目是'%(len(necessary_xpath)),necessary_xpath 
          
        '**********************查找网页元素*********************************'          
        select_necessary_xpath=[]
        select_no_necessary_xpath=[]
        input_necessary_xpath=[]
        inpout_no_necessary_xpath=[]
        '判断下拉或者输入框中，那些是必须项'
        if len(necessary_xpath)!=0:
            for m in range(len(necessary_xpath)):
                print  '这是第%d个必选，判断是否为input。相应路径为：'%(m+1),(necessary_xpath[m][:-3]+'input')
                for i in range(len(input_ele)):
                    print '第%d个input框，xpath为:'%(i),input_path[i]
                    if input_path[i] in (necessary_xpath[m][:-3]+'input'):
                        input_necessary_xpath.append(input_path[i])
                        print  '第%d个inputpath框，是必输入框'%(i)
                    else:
                        inpout_no_necessary_xpath.append(input_path[i])
                        print  '第%d个inputpath框，not必输入框:'%(i)
                           
                print  '这是第%d个必选，判断是否为select。相应路径为：'%(m+1), (necessary_xpath[m][:-3]+'select')
                for i in range(len(select_path)):
                    print '第%d个select框，xpath为:'%(i)+select_path[i]
                    if select_path[i] in (necessary_xpath[m][:-3]+'select'):
                        select_necessary_xpath.append(select_path[i])
                        print '第%d个select框，是必输入框:'%(i)
                    else:
                        select_no_necessary_xpath.append(select_path[i])
                        print '第%d个select框，not必输入框:'%(i)
        print '输入框必选有%d个，具体为：'%(len(input_necessary_xpath)),input_necessary_xpath
        print '下拉框必选有%d个，具体为：'%(len(select_necessary_xpath)),select_necessary_xpath
        print '输入框not必选有%d个，具体为：'%(len(inpout_no_necessary_xpath)),inpout_no_necessary_xpath
        print '下拉框not必选有%d个，具体为：'%(len(select_no_necessary_xpath)),select_no_necessary_xpath          
          
          
          
        '***************************************************************判断、执行******************************************************************************************'          
        '3.填写数值'
        print '首先填写所有参数（最大值），如果存在必选项，后续删除必选项，如果不存在，直接进行第4步--执行并查看结果'
        inputkey = ConfigParser.ConfigParser() 
        inputkey.read(r"D:\OMC\web_input2.ini") 
  
        print '进行输入框的填写'
        for num_input in range(1,len(input_ele)+1):
            input_ele[num_input-1].clear()
            input_ele[num_input-1].send_keys(inputkey.get(modifyobjectxpath[round-1][:-1],str(num_input)))
          
        print '进行下拉列表的选择'
        if len(select_ele)>0:
            for num_select in range(1,len(select_ele)+1):
                print inputkey.get(modifyobjectxpath[round-1][:-1],'select'+str(num_select))
                ActionChains(b).move_to_element(select_ele[num_select-1]).click().perform()
                time.sleep(1)
                n_select=WebDriverWait(b,1).until(lambda x: b.find_element_by_xpath(inputkey.get(modifyobjectxpath[round-1][:-1],'select'+str(num_select))))
                ActionChains(b).move_to_element(n_select).double_click().perform()
                if len(select_ele)==1:
                    ActionChains(b).move_to_element(select_ele[num_select-1]).click().perform()
                    ActionChains(b).move_to_element(select_ele[num_select-1]).click().perform()


        '4.执行并查看结果'
        b.find_element_by_xpath(webele[('执行').decode('utf-8').encode('gbk')]).click()
        time.sleep(1)
          
        print '最大值，点击执行后，是否有弹窗'
        try:
            bp= b.find_element_by_xpath('/html/body/div[48]/div[2]/div[2]').text
            WebDriverWait(b,1).until(lambda x: b.find_element_by_xpath('/html/body/div[48]/div[3]/a/span/span')).click()
            print '最大值，点击执行后，有弹窗'
        except:
            pass
            print '最大值，点击执行后，no有弹窗'
          
        print '最大值，点击执行后，命令是否下发成功'
        if commad_success_flag==1:
            try:
                WebDriverWait(b,300).until(lambda x: b.find_element_by_xpath('/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div/span[1]'))
                print '现在第%d个元素，目前填写最大边界值后，点击执行后，命令下发成功！！！！'%(round)
                commad_success_flag+=1
            except:
                print '现在第%d个元素，目前填写最大边界值后，点击执行后，命令下发成功失败，需提交研发人员处理！！！'%(round)
        else:
            try:     
                WebDriverWait(b,300).until(lambda x: b.find_element_by_xpath('/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/span[1]'%(commad_success_flag)))
                print '现在第%d个元素，目前填写最大边界值后，点击执行后，命令下发成功！！！！'%(round)
                commad_success_flag+=1
            except:
                print '现在第%d个元素，目前填写最大边界值后，点击执行后，命令下发成功失败，需提交研发人员处理！！！'%(round)       
           
                       
        '***************************************************************创建Excel并将结果写入******************************************************************************************'
          
  
           
        '5.读取执行结果，并写入Excel文件'
        for m in range(1,4):
            col = 0
            all_col=[]
            if  m < 3:
                '****************************************************************确定要查找的元素的Xpath*********************************'       
                print 'command_success_result  ',command_success_result
                print '当前第%d行'%m
                if command_success_result==1:
                    'command_success_result==1,意思是第一次显示查询结果。'
                    f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div/span[%d]'%(m)
      
                else:
                    f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/span[%d]'%(round,m)
      
                print 'm为：%d，查询元素的具体xpath为： %s'%(m,f)
                  
                  
                '****************************************************************等待执行的响应,并将第一行和第二行的结果写入Excel*********************************'    
                print '1.获取第1和2行的结果,抬头、基站名称部分'
                try:
                    g=WebDriverWait(b,300).until(lambda x: b.find_element_by_xpath(f)).text
                    worksheet1.write(nrows,0,g)
                except:
                    worksheet1.write(nrows, 0,'等待结果超时，本元素终止查询')
                    print '等待结果超时，本元素终止查询，后续不应该有任何本元素的动作，请注意！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！'
                    continue
                  
                nrows+=1
                print '1、2行的一行内容全部保存完毕'             
               
                '****************************************************************判断结果有多少行*********************************'   
            else:
                '此时m=3'
                print '1. 判断有几行,后续打印具体细节'
                try:
                    f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/div/'%(round)
                    b.find_element_by_xpath(f)
                    print '1）正常的返回值-----只有1行结果nline（行数）=%d'%nline
                except:
                    try:
                        for nlin in range(1,22):
                            f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/div[%d]'%(round,nlin)
                            b.find_element_by_xpath(f)
                            nline=nlin
                        print '2）有多行结果结果的情况nline（行数）=%d'%nline
                    except:
                        try:
                            for nlin in range(1,3):
                                f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/span[%d]'%(round,nlin+2)
                                b.find_element_by_xpath(f)
                                nline=1
                            print '3）结果返回错误值，将nline（行数）直接定义为1行'
                        except:
                            print '查询结果三四行，结果无法判断，请人工查询**************************************！！！！！！！！！！！！！！！！！'
                          
                    print 'nline is',nline
                      
                print'2判断有几列,后续打印具体细节'
                for n in range(1,22):
                    if command_success_result ==1:
                        f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div/div/div[%d]'%(n)
                        print'1）.command_success_result ==1（第一次查询结果）'
                    else:
                        f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/div/div[%d]'%(round,n)
                        print'2）.command_success_result ！=1，不是第一次出查询结果'
                    try:
                        print 'f is ',f
                        findele=b.find_element_by_xpath(f)
                        all_col.append(findele)
                    except:
                        if len(all_col)>0:
                            '证明是有结果，可能只有1列，则不在进行查找'
                            pass
                        else:
                            print '没有返回正确结果，xpath查找错误的返回结果'
                            if m==3 and n ==1:
                                f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/span[3]'%(round)
                            try:
                                findele=b.find_element_by_xpath(f)
                                all_col.append(findele)
                                print '错误的返回结果,强制定为只有1列（实际情况也只有1列）'
                            except:
                                pass
                print 'all_col有%d列,具体元素为：'%(len(all_col)),all_col
               
                      
  
                for ii in range(1,nline+1):
                    for iii in range(1,3):
                        '****************************************************************第3行（或第四行）结果查询，并导出*********************************'
                        for n in range(1,len(all_col)+1):
                            '1.定位元素部分。---i为查询元素，行iii  列n，/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[]/div/div[n]/span[iii]'
                            print 'command_success_result is ',command_success_result
                            print 'iii(1为第3行，2为第4行) is %d'%iii
                            if command_success_result==1:
                                'command_success_result==1,意思是第一次显示查询结果。'
                                f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div/div/div[%d]/span[%d]'%(n,iii)
      
                            else:
                                if nline==1:
                                    f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/div/div[%d]/span[%d]'%(round,n,iii)
                                else:
                                    f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/div[%d]/div[%d]/span[%d]'%(round,ii,n,iii)
                            print 'm为：%d，查询元素的具体xpath为： %s'%(iii,f)
                                    
                              
                            '2.获取第3和第4行的正常结果，具体数值'
                            try:
                                g = b.find_element_by_xpath(f).text
                                worksheet1.write(nrows,n-1,g)
                                print '2.获取第3和第4行的正常结果，写入excel'
                                print '当前的%d行，第%d个元素，xpath为%s'%(iii,n,f)
                                print '**************执行结果数据写入**************%d行，%d列'%(nrows,n-1)
                                  
                               
                                '3.获取第3和第4行的异常结果获取,，写入excel'
                            except:
                                print'3.获取第3和第4行的异常结果获取,，写入excel'
                                if ii==1 and iii==1 and n ==1:
                                    print 'iii==1(第三行) and n ==1（第一列）'
                                    f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/span[3]'%(round)
                                elif  ii==1 and iii==2 and n==1:
                                    print 'iii==2（第四行） and n==1（第一列）'
                                    f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/span[4]'%(round)
                                else:
                                    f=''
                                print '查找错误的返回结果,当前的%d行，第%d个元素，xpath为%s'%(iii,n,f)
                                try:
                                    g = b.find_element_by_xpath(f).text
                                    worksheet1.write(nrows,n-1,g)
                                except:
                                    continue
                                print '*****异常结果数据写入********%d行，%d列'%(nrows,n-1)
  
                        nrows+=1
                        time.sleep(0.2)
                        '3、4行的一行内容填写完成'
              
        if nrows >1:  
            command_success_result +=1  
        print '第%d项测试完毕，结果保存'%round
          
          
    time.sleep(1)
    print '测试完成，关闭测试！！'
    workbook.close()

    arg=importinfo('logout') 
    print 'logout start'
    time.sleep(3)
    logoutbtton = WebDriverWait(b, 10).until(lambda c: b.find_element_by_xpath(arg['smcurl_logouid'])) 
    ActionChains(b).click(logoutbtton).perform()
    time.sleep(1)
    logoutbtton2 =b.find_element_by_xpath(arg['smcurl_logouid2'])
    ActionChains(b).move_to_element(logoutbtton2).click().perform()
    print 'logout over'

if __name__ == '__main__':
    b = webdriver.Firefox()
    logoin(b)
    modify2(b)
#     inquire(b) 