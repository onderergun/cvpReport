#!/usr/bin/env python
#

import argparse
from getpass import getpass
import sys
import json
import requests
from requests import packages
import datetime
import time
import openpyxl
from openpyxl.styles import Font
from decimal import Decimal

import smtplib
import os.path as op
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import encoders

# CVP manipulation class

# Set up classes to interact with CVP API
# serverCVP exception class


class serverCvpError(Exception):
    def __init__(self, value):
        self.value = value
    def __str__(self):
        return repr(self.value)

# Create a session to the CVP server

class serverCvp(object):

    def __init__ (self,HOST,USER,PASS):
        self.url = "https://%s"%HOST
        self.authenticateData = {'userId' : USER, 'password' : PASS}
        requests.packages.urllib3.util.ssl_.DEFAULT_CIPHERS = 'ECDH+AESGCM:DH+AESGCM:ECDH+AES256:DH+AES256:ECDH+AES128:DH+AES:ECDH+3DES:DH+3DES:RSA+AESGCM:RSA+AES:RSA+3DES:!aNULL:!MD5:!DSS'
        from requests.packages.urllib3.exceptions import InsecureRequestWarning
        try:
            requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
        except packages.urllib3.exceptions.ProtocolError as e:
            if str(e) == "('Connection aborted.', gaierror(8, 'nodename nor servname provided, or not known'))":
                raise serverCvpError("DNS Error: The CVP Server %s can not be found" % CVPSERVER)
            elif str(e) == "('Connection aborted.', error(54, 'Connection reset by peer'))":
                raise serverCvpError( "Error, connection aborted")
            else:
                raise serverCvpError("Could not connect to Server")

    def logOn(self):
        try:
            headers = { 'Content-Type': 'application/json' }
            loginURL = "/web/login/authenticate.do"
            response = requests.post(self.url+loginURL,json=self.authenticateData,headers=headers,verify=False)
            if "errorMessage" in str(response.json()):
                text = "Error log on failed: %s" % response.json()['errorMessage']
                raise serverCvpError(text)
        except requests.HTTPError as e:
            raise serverCvpError("Error HTTP session to CVP Server: %s" % str(e))
        except requests.exceptions.ConnectionError as e:
            raise serverCvpError("Error connecting to CVP Server: %s" % str(e))
        except:
            raise serverCvpError("Error in session to CVP Server")
        self.cookies = response.cookies
        return response.json()

    def logOut(self):
        headers = { 'Content-Type':'application/json' }
        logoutURL = "/cvpservice/login/logout.do"
        response = requests.post(self.url+logoutURL, cookies=self.cookies, json=self.authenticateData,headers=headers,verify=False)
        return response.json()
    
    def getInventory(self):
       getURL = "/cvpservice/inventory/devices"
       response = requests.get(self.url+getURL,cookies=self.cookies,verify=False)
       if "errorMessage" in str(response.json()):
           text = "Error, retrieving inventory failed: %s" % response.json()['errorMessage']
           raise serverCvpError(text)
       inventoryList = response.json()
       return inventoryList

    def getSysinfo(self,deviceName,timestamp):
        getURL = "/api/v1/rest/"+deviceName+"/Kernel/sysinfo?pretty&end="+timestamp
        response = requests.get(self.url+getURL,cookies=self.cookies,verify=False)
        if "errorMessage" in str(response.json()):
            text = "Error retrieving Sysinfo failed: %s" % response.json()['errorMessage']
            raise serverCvpError(text)
        sysInfo = response.json()
        return sysInfo
        
    def getTasks(self):
        getURL = "/cvpservice/task/getTasks.do?"
        getParams = {"startIndex":0, "endIndex":0}
        response = requests.get(self.url+getURL,cookies=self.cookies,params=getParams,verify=False)
        if "errorMessage" in str(response.json()):
            text = "Error retrieving tasks failed: %s" % response.json()['errorMessage']
            raise serverCvpError(text)
        tasks = response.json()["data"]
        return tasks


def send_mail(send_from, send_to, subject, message, files,
              server, port, username, password,
              use_tls):

    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(message))

    for path in files:
        part = MIMEBase('application', "octet-stream")
        with open(path, 'rb') as file:
            part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        'attachment; filename="{}"'.format(op.basename(path)))
        msg.attach(part)

    smtp = smtplib.SMTP(server, port)
    if use_tls:
        smtp.starttls()
    smtp.login(username, password)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.quit()


def main():
    
    d1 = time.strftime("%Y_%m_%d_%H_%M_%S", time.gmtime())
    currenttime=time.time()
    
    parser = argparse.ArgumentParser()
    parser.add_argument('--username', required=True)
    parser.add_argument('--cvpServer', required=True)

    args = parser.parse_args()
    username = args.username
    password = getpass()
    cvpServer=args.cvpServer
    
    print ("Attaching to API on %s to get Data" %cvpServer )
    try:
        cvpSession = serverCvp(str(cvpServer),username,password)
        logOn = cvpSession.logOn()
    except serverCvpError as e:
        text = "serverCvp:(main1)-%s" % e.value
        print (text)
    print ("Login Complete")
    inventoryList = cvpSession.getInventory()
    numDevices = len(inventoryList)
    wb = openpyxl.Workbook()
    sheet = wb['Sheet']
    sheet['A1'] = 'Hostname'
    sheet['A1'].font = Font(size=14, bold=True)
    sheet.column_dimensions['A'].width = 20
    sheet['B1'] = 'Model'
    sheet['B1'].font = Font(size=14, bold=True)
    sheet.column_dimensions['B'].width = 20
    sheet['C1'] = 'SW Version'
    sheet['C1'].font = Font(size=14, bold=True)
    sheet.column_dimensions['C'].width = 20
    sheet['D1'] = 'IP Address'
    sheet['D1'].font = Font(size=14, bold=True)
    sheet.column_dimensions['D'].width = 15
    sheet['E1'] = 'Serial Number'
    sheet['E1'].font = Font(size=14, bold=True)
    sheet.column_dimensions['E'].width = 20
    sheet['F1'] = 'Up Time'
    sheet['F1'].font = Font(size=14, bold=True)
    sheet.column_dimensions['F'].width = 20
    sheet['G1'] = 'Daily Availability (%)'
    sheet['G1'].font = Font(size=14, bold=True)
    sheet.column_dimensions['G'].width = 25
    sheet['H1'] = 'CPU Load (%)'
    sheet['H1'].font = Font(size=14, bold=True)
    sheet.column_dimensions['H'].width = 15
    sheet['I1'] = 'Free Memory (%)'
    sheet['I1'].font = Font(size=14, bold=True)
    sheet.column_dimensions['I'].width = 20

    for num,device in enumerate(inventoryList):
        sheet.cell(row=num+2,column=1).value=device["hostname"]
        sheet.cell(row=num+2,column=2).value=device["modelName"]
        sheet.cell(row=num+2,column=3).value=device["version"]
        sheet.cell(row=num+2,column=4).value=device["ipAddress"]
        sheet.cell(row=num+2,column=5).value=device["serialNumber"]
        
        loadAvg15m=0
        percentram=0
        freeram=0
        totalram=0
        unavailabletime = 0
        t=0
        while t<96:
            sysInfo=cvpSession.getSysinfo(device["serialNumber"],str(int(currenttime)-t*900)+"000000000")
            for item in sysInfo["notifications"]:
                if "uptime" in item["updates"]:
                    seconds_input = item["updates"]["uptime"]["value"]["int"]
                    if seconds_input < 900:
                        unavailabletime = unavailabletime + 900 - seconds_input
                    if t==0:
                        conversion = datetime.timedelta(seconds=seconds_input)
                        converted_time = str(conversion)
                        sheet.cell(row=num+2,column=6).value = converted_time

                if "loadAvg15m" in item["updates"]:
                    loadAvg15m = loadAvg15m + Decimal(item["updates"]["loadAvg15m"]["value"]["float"])
                if "freeram" in item["updates"]:
                    freeram = int(item["updates"]["freeram"]["value"]["int"])
                if "totalram" in item["updates"]:
                    totalram = int(item["updates"]["totalram"]["value"]["int"])
                    percentram=percentram + Decimal(freeram)/Decimal(totalram)
            t+=1
        
        sheet.cell(row=num+2,column=7).value = round(100*(1-float(unavailabletime)/float(86400)),2)
        sheet.cell(row=num+2,column=8).value = round(loadAvg15m/96,2)
        sheet.cell(row=num+2,column=9).value = round(100*percentram/96,2)
        
#    Get the completed tasks and who executed it for the last 7 days"
    last7days = []
    for i in range (7):
        last7days.append((datetime.datetime.now() - datetime.timedelta(days=i)).strftime('%d-%m-%Y'))
    
    sheet['A'+str(numDevices+5)] = 'Username'
    sheet['A'+str(numDevices+5)].font = Font(size=14, bold=True)
    sheet['B'+str(numDevices+4)] = '# of config changes'
    sheet['B'+str(numDevices+4)].font = Font(size=14, bold=True)
    sheet['B'+str(numDevices+5)] = last7days[0]
    sheet['B'+str(numDevices+5)].font = Font(size=14, bold=True)
    sheet['C'+str(numDevices+5)] = last7days[1]
    sheet['C'+str(numDevices+5)].font = Font(size=14, bold=True)
    sheet['D'+str(numDevices+5)] = last7days[2]
    sheet['D'+str(numDevices+5)].font = Font(size=14, bold=True)
    sheet['E'+str(numDevices+5)] = last7days[3]
    sheet['E'+str(numDevices+5)].font = Font(size=14, bold=True)
    sheet['F'+str(numDevices+5)] = last7days[4]
    sheet['F'+str(numDevices+5)].font = Font(size=14, bold=True)
    sheet['G'+str(numDevices+5)] = last7days[5]
    sheet['G'+str(numDevices+5)].font = Font(size=14, bold=True)
    sheet['H'+str(numDevices+5)] = last7days[6]
    sheet['H'+str(numDevices+5)].font = Font(size=14, bold=True)
    
    taskList = cvpSession.getTasks()
    userTask = {}
    for task in taskList:
        if task["workOrderState"] == "COMPLETED":
            taskCreator = task["createdBy"]
            if not taskCreator in userTask:
                userTask[taskCreator] = {}
            taskCompletiontime = datetime.datetime.fromtimestamp(task["completedOnInLongFormat"]/1000).strftime('%d-%m-%Y')
            if not taskCompletiontime in userTask[taskCreator]:
                userTask[taskCreator][taskCompletiontime] = 0
            userTask[taskCreator][taskCompletiontime] = userTask[taskCreator][taskCompletiontime]+1

    for num,user in enumerate(userTask):
        sheet.cell(row=numDevices+6+num,column=1).value = user
        for iter,day in enumerate(last7days):
            if day in userTask[user]:
                sheet.cell(row=numDevices+6+num,column=iter+2).value = userTask[user][day]
            else:
                sheet.cell(row=numDevices+6+num,column=iter+2).value = 0

    filename = 'rapor_'+ d1 + '.xlsx'
    wb.save(filename)
    
    print ("Logout from CVP:%s" % cvpSession.logOut()['data'])

    send_from = "sender@domain.com"
    send_to = ["receiver@domain.com"]
    subject = "Daily Report "+d1
    message = ""
    files = [filename]
    server = "smtp.domain.com"
    username = "username"
    password = "password"
    port =587
    use_tls = True
    send_mail(send_from, send_to, subject, message, files, server, port, username, password, use_tls )


if __name__ == '__main__':
    main()

            
