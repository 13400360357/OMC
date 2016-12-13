#coding:utf-8
'''Created on 2016年9月29日
@author: MengLei'''

from basic import *

def compare(excel1,excel2):
    '1. 创建Excel'
    excel=CreateExcel()
    workbook,worksheet=excel.run('compare')
    return workbook,worksheet
    
    '2. 读取文件，并对比'
    readexcel1=xlrd.open_workbook(excel1)
    readsheet1=readexcel1.sheets()[0]
    rows1=readsheet1.nrows
    
    readexcel2=xlrd.open_workbook(excel2)
    readsheel2=readexcel2.sheets()[0]
    rows2=readsheel2.nrows
    
    if rows1 != rows2:
        print 'diff rows,break!'
    else:
        
        '对每一行进行判断'
        n=0
        for row in range(0,rows1):
            context1=readsheet1.row_values(row)
            context2=readsheel2.row_values(row)
            
            '保存标题行数据'
            if 'RESILT' in context1:
                con=[]
                con.append(context1)
            
            '判断本行是否相等 ，不相等，将标题一同写入EXCEL'
            if context1==context2:
                pass
            else:
                    worksheet.write(n,0,con)
                    n+=1
                    worksheet.write(n,0,'line %d not equal,fault'%row)
                    n+=1
                
    workbook.close()
    print '测试结束！！！！！'
    
if __name__ == '__main__':
    
    
    
    
    
    