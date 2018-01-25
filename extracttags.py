import os
import win32com.client
import pandas as pd
import numpy as np
from datetime import datetime

#initialization 
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
sender = []
senderemail =[]
timestamp=[]
receiver=[]
cclist = []
bcclist = []
subject = []
msgbody = []

#read_mail(msg) reads .msg files and extracts relevant information
def read_mail(msg):
     msg=outlook.OpenSharedItem(msg)
     sender.append(msg.SenderName)
     senderemail.append(msg.SenderEmailAddress)
     timestamp.append(msg.SentOn)
     receiver.append(msg.To)
     cclist.append(msg.CC)
     bcclist.append(msg.BCC)
     subject.append(msg.Subject)
     msgbody.append(msg.body)
     msg.Close
     
     
#todatetime() converts Windows' datetime format (win32time) to pd.datetime format
def todatetime (pywin32time):
    x = datetime(
            year = int(pywin32time.Format("%Y")),
            month = int(pywin32time.Format("%m")),
            day = int(pywin32time.Format("%d")),
            hour = int(pywin32time.Format("%H")),
            minute =int(pywin32time.Format("%M")),
            second = int(pywin32time.Format("%S")))
    return x
    


timestamp = [todatetime(time) for time in timestamp]        

#creating the dataframe
df = pd.DataFrame(data = {'Timestamp': timestamp,
                          'Sender': sender,
                          "Recipients": receiver,
                          "CC": cclist,
                          "BCC": bcclist,
                          "Subject": subject,
                          "Message": msgbody})
