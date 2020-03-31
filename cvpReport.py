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
    
    def getConfiglets(self):
        getURL = "/cvpservice/configlet/getConfiglets.do?"
        getParams = {"startIndex":0, "endIndex":0}
        response = requests.get(self.url+getURL,cookies=self.cookies,params=getParams,verify=False)
        if "errorMessage" in str(response.json()):
            text = "Error retrieving configlets failed: %s" % response.json()['errorMessage']
            raise serverCvpError(text)
        configlets = response.json()["data"]
        return configlets
  
    def getInventory(self):
       getURL = "/cvpservice/inventory/devices"
       response = requests.get(self.url+getURL,cookies=self.cookies,verify=False)
       if "errorMessage" in str(response.json()):
           text = "Error, retrieving tasks failed: %s" % response.json()['errorMessage']
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
        
    def snapshotDeviceConfig(self,deviceSerial):
        getURL = "/cvpservice/snapshot/deviceConfigs/"+deviceSerial+"?current=true"
        response = requests.get(self.url+getURL,cookies=self.cookies,verify=False)
        if "errorMessage" in str(response.json()):
            text = "Error retrieving running config failed: %s" % response.json()['errorMessage']
            raise serverCvpError(text)
        deviceConfig = response.json()["runningConfigInfo"]
        return deviceConfig
        
    def getTasks(self):
        getURL = "/cvpservice/task/getTasks.do?"
        getParams = {"startIndex":0, "endIndex":0}
        response = requests.get(self.url+getURL,cookies=self.cookies,params=getParams,verify=False)
        if "errorMessage" in str(response.json()):
            text = "Error retrieving tasks failed: %s" % response.json()['errorMessage']
            raise serverCvpError(text)
        tasks = response.json()["data"]
        return tasks


def send_mail(send_from, send_to, subject, message, files=[],
              server="localhost", port=587, username='', password='',
              use_tls=True):
    """Compose and send email with provided info and attachments.

    Args:
        send_from (str): from name
        send_to (str): to name
        subject (str): message title
        message (str): message body
        files (list[str]): list of file paths to be attached to email
        server (str): mail server host name
        port (int): port number
        username (str): server auth username
        password (str): server auth password
        use_tls (bool): use TLS mode
    """
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
    sheet = wb.get_sheet_by_name('Sheet')
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

    k=0
    for device in inventoryList:
        sheet.cell(row=k+2,column=1).value=inventoryList[k]["hostname"]
        sheet.cell(row=k+2,column=2).value=inventoryList[k]["modelName"]
        sheet.cell(row=k+2,column=3).value=inventoryList[k]["version"]
        sheet.cell(row=k+2,column=4).value=inventoryList[k]["ipAddress"]
        sheet.cell(row=k+2,column=5).value=inventoryList[k]["serialNumber"]
        
        t=0
        loadAvg15m=0
        percentram=0
        freeram=0
        totalram=0
        unavailabletime = 0
        while t<96:
            sysInfo=cvpSession.getSysinfo(inventoryList[k]["serialNumber"],str(int(currenttime)-t*900)+"000000000")
            j=0

            for item in sysInfo["notifications"]:
                if "uptime" in sysInfo["notifications"][j]["updates"]:
                    seconds_input = sysInfo["notifications"][j]["updates"]["uptime"]["value"]["int"]
                    print (seconds_input)
                    if seconds_input < 900:
                        unavailabletime = unavailabletime + 900 - seconds_input
                    if t==0:
                        conversion = datetime.timedelta(seconds=seconds_input)
                        converted_time = str(conversion)
                        sheet.cell(row=k+2,column=6).value = converted_time

                if "loadAvg15m" in sysInfo["notifications"][j]["updates"]:
                    loadAvg15m = loadAvg15m + Decimal(sysInfo["notifications"][j]["updates"]["loadAvg15m"]["value"]["float"])
                if "freeram" in sysInfo["notifications"][j]["updates"]:
                    freeram = int(sysInfo["notifications"][j]["updates"]["freeram"]["value"]["int"])
                if "totalram" in sysInfo["notifications"][j]["updates"]:
                    totalram = int(sysInfo["notifications"][j]["updates"]["totalram"]["value"]["int"])
                    percentram=percentram + Decimal(freeram)/Decimal(totalram)
                j+=1
            t+=1

        print (unavailabletime)
        
        sheet.cell(row=k+2,column=7).value = round(100*(1-float(unavailabletime)/float(86400)),2)
        sheet.cell(row=k+2,column=8).value = round(loadAvg15m/96,2)
        sheet.cell(row=k+2,column=9).value = round(100*percentram/96,2)
    
        k=k+1
        
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
    k=0
    userTask = {}
    for task in taskList:
        if taskList[k]["workOrderState"] == "COMPLETED":
            taskCreator = taskList[k]["createdBy"]
            if not taskCreator in userTask:
                userTask[taskCreator] = {}
            taskCompletiontime = datetime.datetime.fromtimestamp(taskList[k]["completedOnInLongFormat"]/1000).strftime('%d-%m-%Y')
            if not taskCompletiontime in userTask[taskCreator]:
                userTask[taskCreator][taskCompletiontime] = 0
            userTask[taskCreator][taskCompletiontime] = userTask[taskCreator][taskCompletiontime]+1
        k=k+1
    k=0
    for user in userTask:
        sheet.cell(row=numDevices+6+k,column=1).value = user
        j=2
        for day in last7days:
            if day in userTask[user]:
                sheet.cell(row=numDevices+6+k,column=j).value = userTask[user][day]
            else:
                sheet.cell(row=numDevices+6+k,column=j).value = 0
            j+=1
        k+=1
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

            
