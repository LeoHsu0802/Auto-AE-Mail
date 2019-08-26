# -*- coding: utf-8 -*-
"""
Created on Tue Jul 30 11:16:37 2019

@author: Leo
"""

# -*- coding: utf-8 -*-

import win32com.client as win32
from openpyxl import load_workbook

wb = load_workbook('CustomerAC.xlsx')
sheet = wb.active 
def AE_Mail(Part,AE,Date,Customer,ACC,Reason):
    
    Part = str(Part)
    if Part == "1":
        Part = "業務一部主管"
    elif Part =="2":
        Part = "業務二部主管"
    elif Part =="3":
     	 Part = "業務三部主管"
    
    AE = str(AE)            #A
    AEname =AE[-2:]
    Date = str(Date)        #C  
    Date = Date[0:10]    
    
    Customer = str(Customer)#D
    
    ACC = str(ACC)          #E
    ACC = ACC[3:]
    
    Reason = str(Reason)    #I
    print(AE,Date,Customer,ACC,Reason)
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = AE
    mail.CC = Part
    mail.Subject ='客戶國內期貨日報電寄失敗'
    mail.HTMLBody ='''\
                    <h3>Hi '''+AEname+'''</h3>
                    <div></div>
                    <h3>您的客戶<span style="color:#DB8F00;">'''+ACC+Customer+'''</span><h3>
                    <h3>'''+Date+'''國內期貨日對帳單寄送失敗，原因為<span style="color:#DB8F00;">'''+Reason+'''</span><h3>
                    <h3>請與客戶連絡，是否退信或信箱已滿而無法收件，本次電寄失敗將於七日後郵寄紙本予客戶</h3>
                    <h3 style="color:red;">請於今日下班前通知客戶後回覆此信件，以利補寄對帳單，謝謝!</h3>
                    '''
    mail.Send()
    
for i in range(2,60):
    i = str(i)
    AE_Mail(sheet[('A'+i)].value,sheet[('B'+i)].value,sheet[('C'+i)].value,sheet[('D'+i)].value,sheet[('E'+i)].value,sheet[('I'+i)].value)
