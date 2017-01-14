#coding:utf-8
'''Created on 2017年1月11日
@author: MengLei'''
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import time,os,random
from email.utils import parseaddr, formataddr
from email.header import Header
from email import encoders

from basic import importinfo
class mail():
    def fasong(self,excel_path=''):
        '导入相关数据'
        ini=importinfo()
        mail = ini.run('mail.ini','mail_smtp')
        
        '#如名字所示Multipart就是分多个部分'
        msg = MIMEMultipart()
        if excel_path!='':
            '附件名称当作邮件主题'
            print '附件名称当作邮件主题%s'%str(excel_path.split(os.sep)[-1:]).split('\'')[1]
            msg['subject'] = '%s'%str(excel_path.split(os.sep)[-1:]).split('\'')[1]
        else:
            msg['subject'] = '请见邮件正文'+str(random.randint(1,99999999))
        msg["From"]  = mail['from_addr']
        msg["To"]   =  mail['to_addr']
           
            
        '#---这是内容部分---' 
        part = MIMEText(mail['mail_text'],'plain', 'utf-8') 
        msg.attach(part) 
             
        '#---这是附件部分,如有过附件则添加--- '
        if excel_path!='':
            part = MIMEApplication(open(excel_path,'rb').read()) 
            print 'excel_path is:',excel_path
            print 'excel_path.split(os.sep)[-1:] is',str(excel_path.split(os.sep)[-1:]).split('\'')[1]
            part.add_header('Content-Disposition', 'attachment', filename='%s'%str(excel_path.split(os.sep)[-1:]).split('\'')[1])
            msg.attach(part) 
             
        '发送邮件'
        try:
            print 'start email...'
            server = smtplib.SMTP(mail['smtp_server'], 25) # SMTP协议默认端口是25
#             server.set_debuglevel(1)
            server.login(msg["From"] , mail['password'])
            server.sendmail(msg["From"] , msg["To"].split(',')[0:], msg.as_string())
            time.sleep(0.5)
            server.quit()
            print 'email sucess...'
        except Exception, e:
            print  'email failed:',str(e) 
             
if __name__ == '__main__':
    '''可以发送附件，也可以不带附件。
        附件时,邮件主题为附件名称，
        不带附件时，邮件主题为：请见邮件正文'''
#     mail=mail()
    
    '1. 附件时,邮件主题为附件名称'
#     mail.fasong(r'D:\Git\report\inquire2017-01-11-3.xlsx')
    
    '2. 不带附件时，邮件主题为：请见邮件正文'
#     mail.fasong()








