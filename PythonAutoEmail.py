# -*- coding: utf-8 -*-
"""
Created on Mon Nov 27 12:05:55 2017

@author: clahn
"""

# -*- coding: utf-8 -*-
"""
Created on Wed Sep 21 15:36:00 2016

@author: Deepesh.Singh
"""

import win32com.client as win32
import psutil
import os
import subprocess

# Drafting and sending email notification to senders. You can add other senders' email in the list


def send_notification():
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'christopher.lahn@sanfordhealth.org; Ryan.Bosca@SanfordHealth.org; Brent.Colby@sanfordhealth.org; Danielle.M.Goetz@SanfordHealth.org; William.Duppler@SanfordHealth.org; Kyle.McCallum@sanfordhealth.org; Lanny.Molstad@sanfordhealth.org; Steve.Cassola@sanfordhealth.org'
    mail.Subject = 'API Monday. Please review and sign off on your time cards.  Thanks!'
    mail.body = 'Please review and sign off on your time cards.  Thanks!  This email alert is auto generated. Please do not respond.'
    mail.send

# Open Outlook.exe. Path may vary according to system config
# Please check the path to .exe file and update below


def open_outlook():
    try:
        subprocess.call(['C:\Program Files\Microsoft Office\Office16\Outlook.exe'])
        os.system("C:\Program Files\Microsoft Office\Office16\Outlook.exe");
    except:
        print("Outlook didn't open successfully")


# Checking if outlook is already opened. If not, open Outlook.exe and send email
for item in psutil.pids():
    p = psutil.Process(item)
    if p.name() == "OUTLOOK.EXE":
        flag = 1
        break
    else:
        flag = 0

if (flag == 1):
    send_notification()
else:
    open_outlook()
    send_notification()
