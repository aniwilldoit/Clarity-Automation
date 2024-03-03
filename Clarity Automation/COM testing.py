# -*- coding: utf-8 -*-
"""
Created on Mon Oct 22 01:31:48 2018

@author: pdas7
"""
import win32com.client
from win32com.client import Dispatch, constants

const=win32com.client.constants
olMailItem = 0x0
obj = win32com.client.Dispatch("Outlook.Application")
newMail = obj.CreateItem(olMailItem)
newMail.Subject = "Resource creation"
# newMail.Body = "I AM\nTHE BODY MESSAGE!"
newMail.BodyFormat = 2 # olFormatHTML https://msdn.microsoft.com/en-us/library/office/aa219371(v=office.11).aspx
newMail.HTMLBody = "<HTML><BODY>Please find <span style='color:red'>attached</span> file.</BODY></HTML>"
newMail.To = "anikesh.sinha@capgemini.com"
newMail.Cc = "anikesh.sinha@capgemini.com"
attachment1 = r"C:\\Users\\pdas7\\Desktop\\ResourceCreation1.xlsx"
newMail.Attachments.Add(Source=attachment1)
newMail.display(True)
newMail.send()