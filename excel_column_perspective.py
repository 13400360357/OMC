#coding:utf-8
'''Created on 2017年1月12日
@author: MengLei'''

import xlsxwriter,xlrd,time,os
class excel_perspective():
    def run(self,excel_path):
        
        '创建透视的Excel'
        workbook = xlsxwriter.Workbook('report'+os.sep+'perspective_'+str(excel_path.split(os.sep)[-1:]).split('\'')[1])
        worksheet = workbook.add_worksheet('RESULT')
        worksheet1 = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet()
        
        '开始xlrd模块的读取'
        data=xlrd.open_workbook(excel_path)
        
        '复制第一个sheet'
        table1 = data.sheet_by_index(0)
        nrows1 = table1.nrows
        for i in range(nrows1):
            worksheet1.write_row(i, 0, table1.row_values(i))
        
        '复制第2个sheet'
        try:
            table2 = data.sheet_by_index(1)
            if table2.nrows>0:
                for i in range(table2.nrows):
                    worksheet2.write_row(i, 0, table2.row_values(i))
        except:
            print 'no table2...'
            pass
        
        '透视第一个sheet'
        row_values=table1.col_values(0)
#         print row_values
        
        "依次筛选出想要的值"
        a=[]
        for i in range(len(row_values)):
            if a.count(row_values[i])>0:
                pass
            else:
                a.append(row_values[i])
#         print 'a',a
                
        "确定根据筛值出现的次数"
        b=[]
        for i in range(len(a)):
            b.append(row_values.count(a[i]))
#         print 'b',b
        
        '# #写入一行   '
        # worksheet.write_row('A1', headings, bold)  
        '# #写入一列  '
        worksheet.write_column('A1',a) 
        worksheet.set_column('A:A', len('Upgrade_result')+1)
        worksheet.write_column('B2',b[1:]) 
        worksheet.write(0,4,'total times:') 
        worksheet.set_column('E:E', len('total times: '))
        worksheet.write(0,5,nrows1-1)         
        workbook.close()
        
if __name__ == '__main__':
#     excel_perspective=excel_perspective()
#     excel_perspective.run(r'D:\Git\report\upgrade2017-01-13.xlsx')
    pass