#coding:utf-8
'''Created on 2016年12月6日
@author: MengLei'''

from basic import *

class modify:
    def run(self,b,yemianyuansu,workbook,worksheet1,txt,input_txt):
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
        for round in range(1,len(modifyobjectxpath)+1):
#         for round in range(1,2):
            select=b.find_element_by_xpath(yemianyuansu['select'])
            ActionChains(b).move_to_element(select).click().perform()
            time.sleep(1)
            print '**************************** start test %d ************** total %d **************'%(round,len(modifyobjectxpath)),modifyobjectxpath[round-1]
            round_modify = b.find_element_by_xpath(modifyobjectxpath[round-1])
            ActionChains(b).move_to_element(round_modify).click().perform()
            time.sleep(1)
              
            print 'inquire  {quantity | type(input or select) | necessary or not...}'

            '1)确定有多少要被修改的元素,下一步确定是input或者select'          
            all_modify_ele=[]
            for aaa in range(20):
                    f='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div[2]/div/form/ul/li[%d]'%(aaa)
                    try:
                        findele=b.find_element_by_xpath(f)
                        all_modify_ele.append(findele)
                    except:
                        continue
#             print '%d inputbox or selectbox'%(len(all_modify_ele))              
              
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
                    selectele=Select(b.find_element_by_xpath(ele1))
                    select_ele.append(selectele)
                    select_path.append(ele1)
                    
#             print'%d inputbox'%(len(input_ele))
#             print'%d selectbox'%(len(select_ele))                
    #         print'有%d个输入框,下拉列表的xpath是:'%(len(input_ele)),input_path
    #         print'有%d个下拉框,输入列表的xpath是:'%(len(select_ele)),select_path
    
            '3)确定有多少个必选项目'
            necessary_xpath=[]
            for modifyobject in range(1,len(all_modify_ele)+1):
                ele='/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div[2]/div/form/ul/li[%d]/div'%(modifyobject)
                try:
                    necetele=b.find_element_by_xpath(ele)
                    if  necetele.text:
                        necessary_xpath.append(ele)
                except:
                    continue
#             print'%d necessary'%(len(necessary_xpath))
    #         print'有%d个必选项目, 必选项目是'%(len(necessary_xpath)),necessary_xpath 
              
            '4)判断那些是必须输入框'
            input_necessary_xpath=[]
            if len(necessary_xpath)!=0:
                for m in range(len(necessary_xpath)):
                    '这是第%d个必选，判断是否为input。相应路径为：'%(m+1),(necessary_xpath[m][:-3]+'input')
                    for i in range(len(input_ele)):
    #                     print '第%d个input框, xpath为:'%(i), input_path[i]
                        if input_path[i] in (necessary_xpath[m][:-3]+'input'):
                            input_necessary_xpath.append(input_path[i])
    #                         print  '第%d个inputpath框, 是必输入框'%(i)
                        else:
                            pass
            print '%d necessary input'%(len(input_necessary_xpath))
              
            '***********************************************************判断、执行******************************************************************************************'
            '''
            1.全部不填写，执行，并查看结果
            2.所有input部分全部填写超出范围，判断是否提示。 
            3.input数值正常，界面遗漏1项必填内容（输入框/下拉列表），执行并查看执行结果。循环遍历所有必填项目。
            4.正常填写，执行并查看结果            '''
            
            print 'start no context...'                 
            
            '1.全部不填写，执行'
            b.find_element_by_xpath(yemianyuansu['zhixing']).click()
            time.sleep(1)
            '判断是否有弹窗'
            try:
                n = WebDriverWait(b, 3).until(lambda x: b.find_element_by_xpath('/html/body/div[57]/div[3]/a/span/span')).click()
                'nothing be wirtten. pass'
            except:
                print'nothing be wirtten.no messagebox. fault！！'
                txt.write('\n%d ele:%s,nothing be wirtten,no messagebox,failure.'%(round,modifyobjectelepath[round-1]))
                
            
            '2所有input部分全部填写超出范围(特殊字符)，判断是否提示。 '
            time.sleep(0.5)
            if len(input_ele)>0:
                
                print 'start  illegal character...'  
                for num_outinput in range(1,len(input_ele)+1):
                    b.find_element_by_xpath(input_path[num_outinput-1]).clear()
                    b.find_element_by_xpath(input_path[num_outinput-1]).send_keys('&&&&%%%%%$$(@#$%^!@#$%^&*(#$%^&*($%^&*()@#$%^&*()_#$%^&*()$%^&*@#$%^&*()#$%^&*($%^&*&*()!@#$%^&*(@#$%^&*(@#$%^&*@#$%^&')
                    time.sleep(0.5)
                    b.find_element_by_xpath(yemianyuansu['zhixing']).click()
                    checkinp= input_path[num_outinput-1][:-5]+'div'
                    time.sleep(1)
                      
                    '2.1判断是否有弹窗，如果有，点击确定'
                    try:
                        b.find_element_by_xpath('/html/body/div[57]/div[3]/a/span/span').click()
                        txt.write('\n%d ele:%s, %d inputbox,special context,messagebox,fault.'%(round,modifyobjectelepath[round-1],num_outinput))
                    except:
                        'special context, no messagebox, pass.'
                        pass
                       
                    '2.2判断是否有超出范围的提示'
                    try:
                        result=b.find_element_by_xpath(checkinp)
                    except:
                        print '%d input, special context, no prompt, fault'%(num_outinput)
                        txt.write('\n%d ele:(%s),%d inputbox,special context,no prompt,failure'%(round,modifyobjectelepath[round-1],num_outinput))
                    finally:
                        b.find_element_by_xpath(input_path[num_outinput-1]).clear()
            else:
                pass
              
            '***************************************************************判断、执行******************************************************************************************'
              
            '3.遗漏必填项，进行测试。--先正常填写（最大边界值），然后挨个删除必填项目，执行并查看结果'
            inputkey = ConfigParser.ConfigParser() 
#             inputkey.read("web_input.ini") 
            inputkey.read(input_txt)
            
            print 'write  max_value...'
            '3.1进行输入框的填写'
            for num_input in range(1,len(input_ele)+1):
                input_ele[num_input-1].clear()
                input_ele[num_input-1].send_keys(inputkey.get(modifyobjectxpath[round-1][:-1],str(num_input)))
                time.sleep(0.5)
              
            '3.2进行下拉列表的选择'
            if len(select_ele)>0:
                for num_select in range(1,len(select_ele)+1):
                    (select_ele[num_select-1]).select_by_value(inputkey.get(modifyobjectxpath[round-1][:-1],'select'+str(num_select)))
                    time.sleep(0.5)
                          
            '3.3 遗漏必选输入框，进行下发'
            if len(input_necessary_xpath)>0 :
                print 'start necessary_inputbox，no context...' 
                for num_input in range(1,len(input_ele)+1):  
                    if input_path[num_input-1] in input_necessary_xpath:
                        input_ele[num_input-1].clear()
                        b.find_element_by_xpath(yemianyuansu['zhixing']).click()
                        time.sleep(3)
                        try:
                            WebDriverWait(b,5).until(lambda x: b.find_element_by_xpath('/html/body/div[1]/div[4]/div/div/div[2]/div/div/div[1]/div[2]/div[%d]/span[1]'%(commad_success_flag)))
                            commad_success_flag+=1
#                             print '遗漏必填项，input框部分，命令下发成功，用例测试失败'
                            txt.write('\n%d ele:(%s), %d necessary_inputbox, no context and excute success, failure.'%(round,modifyobjectelepath[round-1],num_input))
                        except:
                            pass
#                             print '遗漏必填项，input框部分，命令未下发，用例通过'                      
                        input_ele[num_input-1].send_keys(inputkey.get(modifyobjectxpath[round-1][:-1],str(num_input)))
                time.sleep(0.5)
#                 print '3. 遗漏必遗漏必选输入框，测试完成'
            else:
                pass
#                 print '3. 没有必选输入框，跳过遗漏'
            
            '4.正常填写（最大边界值），执行并查看结果（接着3的结果，直接点击执行，判断即可）'
            print 'start  max_value...'
            b.find_element_by_xpath(yemianyuansu['zhixing']).click()
            time.sleep(1)
            
            '4.1判断是否有弹窗'
            try:
                b.find_element_by_xpath('/html/body/div[57]/div[3]/a/span/span').click()
                txt.write('\n%d ele:%s，max value，messagebox . failure'%(round,modifyobjectelepath[round-1]))
                print 'max value, click on execute, pop-up window, failure.'
            except:
                pass
              
            '4.2最大值，点击执行后，命令是否下发成功'
            try:
                WebDriverWait(b,300).until(lambda x: b.find_element_by_xpath('//*[@id="showParamValues"]/div[%d]/div'%(commad_success_flag)))
                commad_success_flag+=1
    #             print 'result get,prepare for writting...'
    #             print '填写最大边界值，命令下发成功。修改功能测试结束。即将进行结果读取并保存到Excel'
            except:
                print 'max value, after execute, no result, failure.'%(round)    
                txt.write('\n %d ele:%s, max value, after execute, no result, failure.'%(round,modifyobjectelepath[round-1]))
            
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
                    print 'time out for reading result.'
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
                    print 'result incorrect, failure.'
                    txt.write('\n result incorrect, failure.')
                        
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
                    print '%d line writting into Excel.'%(2*ii+iii)
            
            print 'start the next...'
            
        time.sleep(1)
        workbook.close()
        print '***************************test finished!   close workbook...'


if __name__ == '__main__':
    session=prepare()
    b=session.login()
    try:
        b.find_element_by_xpath('/html/body/div[39]/div[3]/a/span/span').click()
    except:
        pass
    ini=session.opencatalogue(b,'web_ele.ini','zuocemulu', 'jizhan', 'jizhan_peizhi')
    yemianyuansu=session.import_yemianyuansu(b, ini)
    session.select_enb(b,yemianyuansu)
    workbook,worksheet1=session.excel('modify')
    txt=session.txt('modify')
    
    session2=modify()
    session2.run(b,yemianyuansu,workbook,worksheet1,txt,'web_input2.ini')
    
    
    
    