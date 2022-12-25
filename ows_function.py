#functions associated with auto_escalate ows
from pathlib import Path
import re
from tkinter import *
from tkinter import messagebox
from tkinter.scrolledtext import ScrolledText
import tkinter.ttk 
from tkinter import filedialog as fd
import pandas as pd
import numpy as np
import os
from datetime import *
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException as Exception
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import ElementNotInteractableException
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
from time import sleep
from datetime import datetime,date,time
from PIL import ImageTk, Image
import sys


#excel database to reference
rd_base=pd.read_excel('T:\\SEUN\\REPORTDBASE2.xlsm')[['SITE_ID','TECH STATE']]


def time_criteria():
    return datetime.now()-timedelta(minutes=75)

def clean_alarm(string):
    if 'OML Fault' in string:
        return string.replace('OML Fault','2G DOWN:')
    elif 'CSL Fault' in string:
        return string.replace('CSL Fault','2G DOWN:')
    elif 'Main' in string:
        matches=re.findall('M\w{3}',string)
        return matches[0].replace('Main','Rectifier Main')
    elif 'Urgent' in string:
        matches=re.findall('U\w{5}',string)
        return matches[0].replace('Urgent','Rectifier Urgent')
    elif 'System' in string:
        matches=re.findall('S\w{5}',string)
        return matches[0].replace('System','Gen System')    
    elif 'Gen' in string or 'Gen2' in string:
        return (string.replace(string,'Gen System'))
    elif 'NodeB' in string:
        return string.replace(string,'3G DOWN:')


def test_run(main):
    alarm_arrangement=['Rectifier Main','Rectifier Urgent','Gen System','2G DOWN:','3G DOWN:']
    c=[]
    for i in main:
        if i in alarm_arrangement:
            if i not in c:
                c.append(i)
            else:
                pass
        else:
            c.append(i)
    return c  

def clean_site(string):
    if 'E' in string  and len(string)>6 and string.index('E')==0:
        return string.lstrip('E')
    elif 'D' in string  and len(string)>6 and string.index('D')==0:
        return string.replace('D',' ',1).strip()
    elif 'U' in string and len(string)>6 and string.index('U')==0:
        return string.replace('U',' ',1).strip()
    elif 'C' in string and len(string)>6 and string.index('C')==0:
        return string.replace('C',' ',1).strip()
    else:
        return string



def ows_send_button(driver,window_driver):
    now=datetime.now()
    url='http://10.100.111.22/smsfeeder'
    now=datetime.now()
    if now.hour in range(6,21):
        url2='http://10.100.111.222/smsfeeder'
    else:
        url2='http://10.100.111.221/smsfeeder'
        
    
    rd_base=pd.read_excel('T:\\SEUN\\REPORTDBASE2.xlsm')[['SITE_ID','TECH STATE']]
    sleep(45)
    path1=Path('C:\\Users\\user\\Downloads')
    bg_files=[file for file in path1.glob('*.xlsx')]
    bg_report=max(bg_files,key=os.path.getmtime)
    bg_report=pd.read_excel(bg_report)
    y=time_criteria()
    bg_report['lastoccurrence']=pd.to_datetime(bg_report['lastoccurrence'])
    bg_report=bg_report[(bg_report['lastoccurrence']>=y) & (bg_report['tt duration'].isnull()==True) & (~bg_report['node'].str.contains('_BSC'))]
    bg_report['alarmname']=bg_report['alarmname'].apply(lambda x : clean_site(x))
    bg_report.rename(columns={'node':'SITE_ID','alarmname':'Alarm'},inplace=True)
    sheet=bg_report.merge(rd_base,on='SITE_ID',how='left')
    sheet.rename(columns={'TECH STATE':'STATE'},inplace=True)
    for index,row in sheet.iterrows():
        if row.Alarm=='CSL Fault':
            sheet.at[index,'SITE_ID']='D' + row.SITE_ID
        
    sheet.Alarm=sheet.Alarm.apply(lambda x: clean_alarm(x))
    alarm_arrangement=['Rectifier Main','Rectifier Urgent','Gen System','2G DOWN:','3G DOWN:']
    sheet['Alarm']=pd.Categorical(sheet.Alarm,alarm_arrangement,ordered=True)
    sheet.sort_values(['STATE','Alarm'],inplace=True)
    sheet['Alarm']=sheet.Alarm.astype(str)
    sheet=sheet.drop_duplicates(subset=['SITE_ID','Alarm','STATE'],keep='first')
    sheet.fillna('others',inplace=True)
    sheet['comb']=''
    for index,row in sheet.iterrows():
        sheet.at[index,'comb']=[row.Alarm] + [row.SITE_ID]
    grouping=sheet.groupby('STATE')['comb'].sum().reset_index()
    grouping.comb=grouping.comb.apply(lambda x: test_run(x))
    for a,b in zip(grouping.STATE,grouping.comb):
        if 'nan' in b:
            b.remove('nan')
    no_details=grouping[grouping.STATE=='others']
    b=[]
    for i in no_details.comb:
        for a in i:
            if a not in alarm_arrangement:
                b.append(a)
    with open('Z:\\noc\\SITE NOT IN RDBASE\\no_details.txt','a') as f:
        f.writelines('%s\n' % a for a in b)
    driver.execute_script('window.open();')
    driver.switch_to.window(driver.window_handles[1])
    try:
        sms_window=driver.get(url)  
        user_input=driver.find_elements_by_class_name("TextField")
        input_username=user_input[0].send_keys('nmc')
        input_password=user_input[1].send_keys('****') #protected
        login_button = driver.find_element_by_class_name("Button" )
        login_button.click()
    except:
        sms_window=driver.get(url2)  
        user_input=driver.find_elements_by_class_name("TextField")
        input_username=user_input[0].send_keys('nmc')
        input_password=user_input[1].send_keys('****') #protected
        login_button = driver.find_element_by_class_name("Button" )
        login_button.click()
    message_type=driver.find_element_by_xpath("//*[@id='mtype']/option[text()='Text Message']").click()
    sender_id=driver.find_element_by_xpath("/html/body/div[1]/div[2]/form/div[2]/input").send_keys('NMC')
    for row in grouping.itertuples():
        if row.STATE=='others':
            pass
        else:
            with open('test1.txt','w+') as filehandle:
                filehandle.writelines('%s\n' % a for a in row.comb)
                filehandle.seek(0)
                msg=filehandle.read()
                
            if len(msg)>=391:
                message=driver.find_element_by_xpath("//*[@id='msg']").send_keys(msg[:392])
                region_selection=driver.find_element_by_name('userfile').send_keys(f'T:\\SEUN\\BULK SMS NEWKINGZ\\NEW STRUCTURE REG&EMC\\{row.STATE}.txt')
                driver.find_element_by_name('btn1').click()
                driver.back()
                driver.find_element_by_xpath("//*[@id='msg']").clear()
                message=driver.find_element_by_xpath("//*[@id='msg']").send_keys(msg[390:])
                region_selection=driver.find_element_by_name('userfile').send_keys(f'T:\\SEUN\\BULK SMS NEWKINGZ\\NEW STRUCTURE REG&EMC\\{row.STATE}.txt')
                driver.find_element_by_name('btn1').click()
                driver.back()
                driver.find_element_by_xpath("//*[@id='msg']").clear()
            else:
                message=driver.find_element_by_xpath("//*[@id='msg']").send_keys(msg)
                region_selection=driver.find_element_by_name('userfile').send_keys(f'T:\\SEUN\\BULK SMS NEWKINGZ\\NEW STRUCTURE REG&EMC\\{row.STATE}.txt')
                driver.find_element_by_name('btn1').click()
                driver.back()
                driver.find_element_by_xpath("//*[@id='msg']").clear()
    
        
def export(driver):
    try:
        freeze=driver.find_elements_by_xpath('/html/body/div[1]/div/div/div/div[1]/form/div/div[1]/div/div[1]/div/div/span')[0].click()
        extend=driver.find_element_by_xpath('/html/body/div[1]/div/div/div/div[1]/form/div/div[1]/div/div[16]/div/div/span').click()
        export=driver.find_element_by_xpath('/html/body/div[1]/div/div/div/div[1]/form/div/div[1]/div/div[14]/div[17]/div/div/span[1]').click()
        sleep(20)
    except ElementNotInteractableException:
        freeze=driver.find_elements_by_xpath('/html/body/div[1]/div/div/div/div[1]/form/div/div[1]/div/div[1]/div/div/span')[0].click()
        export=driver.find_element_by_xpath('/html/body/div[1]/div/div/div/div[1]/form/div/div[1]/div/div[14]/div[17]/div/div/span[1]').click()
        sleep(20)
    
def ows_login(driver,window_driver):
    driver.get('https://sgdev.teleows.com')
    user_name=driver.find_element_by_id('usernameInput').send_keys('******') #protected
    password=driver.find_element_by_id('password').send_keys('*******') #protected
    driver.find_element_by_class_name('login_button').click()
    try:
        driver.find_element_by_xpath('//*[@id="app-nav"]/div[1]/div[6]/div/div[1]')
        sleep(5)
        menu=driver.find_element_by_xpath('//*[@id="app-nav"]/div[1]/div[6]/div/div[1]').click()
        driver.find_element_by_xpath('//*[@id="app-nav"]/div[1]/div[10]/div[1]/div/div/div[2]/div[1]/div/div[1]').click()
        driver.switch_to.frame(driver.find_element_by_id('ext-gen27'))
        sleep(20)
        frame2=driver.find_element_by_id('center_panel')
        sleep(5)
        driver.switch_to.frame(frame2)
        driver.find_element_by_xpath('//*[@id="filterSelect"]/div/div[1]/input[2]').click()
        sleep(1)
        share_filter=driver.find_element_by_xpath('/html/body/div[1]/div/div/div/div[1]/form/div/div[1]/div/div[21]/div/div/div/div[2]/div[1]/ul/li[3]/div').click()
        driver.find_element_by_xpath('/html/body/div[1]/div/div/div/div[1]/form/div/div[1]/div/div[21]/div/div/div/div[2]/div[2]/ul/li[1]/div').click()
        sleep(4)
        try:
            ack=driver.find_element_by_class_name('row_init').text
            sort_by_acked=driver.find_element_by_xpath('/html/body/div[1]/div/div/div/div[2]/div[1]/div/div/div/div[3]/div[1]/div[1]/div/table/thead/tr/td[2]/div')
            if ack=='Yes':
                sort_by_acked.click()
        except Exception:
            ack=driver.find_element_by_class_name('row_init').text
            sort_by_acked=driver.find_element_by_xpath('/html/body/div[1]/div/div/div/div[2]/div[1]/div/div/div/div[3]/div[1]/div[1]/div/table/thead/tr/td[2]/div')
            if ack=='Yes':
                sort_by_acked.click()
    except:
        print('Error','incorrect username or password, enter the correct combination')
        sleep(10)
        driver.find_element_by_xpath('//*[@id="app-nav"]/div[1]/div[6]/div/div[1]')
        sleep(5)
        menu=driver.find_element_by_xpath('//*[@id="app-nav"]/div[1]/div[6]/div/div[1]').click()
        driver.find_element_by_xpath('//*[@id="app-nav"]/div[1]/div[10]/div[1]/div/div/div[2]/div[1]/div/div[1]').click()
        driver.switch_to.frame(driver.find_element_by_id('ext-gen27'))
        sleep(20)
        frame2=driver.find_element_by_id('center_panel')
        sleep(5)
        driver.switch_to.frame(frame2)
        driver.find_element_by_xpath('//*[@id="filterSelect"]/div/div[1]/input[2]').click()
        sleep(1)
        share_filter=driver.find_element_by_xpath('/html/body/div[1]/div/div/div/div[1]/form/div/div[1]/div/div[21]/div/div/div/div[2]/div[1]/ul/li[3]/div').click()
        driver.find_element_by_xpath('/html/body/div[1]/div/div/div/div[1]/form/div/div[1]/div/div[21]/div/div/div/div[2]/div[2]/ul/li[1]/div').click()
        sleep(4)
        try:
            ack=driver.find_element_by_class_name('row_init').text
            sort_by_acked=driver.find_element_by_xpath('/html/body/div[1]/div/div/div/div[2]/div[1]/div/div/div/div[3]/div[1]/div[1]/div/table/thead/tr/td[2]/div')
            if ack=='Yes':
                sort_by_acked.click()
        except Exception:
            ack=driver.find_element_by_class_name('row_init').text
            sort_by_acked=driver.find_element_by_xpath('/html/body/div[1]/div/div/div/div[2]/div[1]/div/div/div/div[3]/div[1]/div[1]/div/table/thead/tr/td[2]/div')
            if ack=='Yes':
                sort_by_acked.click()
    export(driver)
    ows_send_button(driver,window_driver)
   