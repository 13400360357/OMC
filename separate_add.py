#coding:utf-8
'''
Created on 2016年12月6日

@author: MengLei
'''

from basic import *



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
        
    def run(self,ini):
        '1. 导入元素'
        elelist=ini.read('add.ini','LST')
        dict={}
        dict.update(elelist)
        
        '2. 执行LST'
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

    session2=omc_add()
    session2.run(ini,'web_ele.ini','inquire')
    
    
