# -*- coding: utf-8 -*-
"""
Created on Tue Oct 23 20:27:36 2018

@author: pdas7
"""
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
                                    # the inbox. You can change that number to reference
                                    # any other folder
messages = inbox.Items
message = messages.GetLast()
body_content = message.body
print body_content