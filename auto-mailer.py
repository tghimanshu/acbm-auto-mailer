print('*'*50)
print('WELCOME TO AUTO EMAILER FOR ACBM')
print('ver. 0.1.0')
print('CREATED BY: TGHIMANSHU')
print('*'*50)
print('\nPLEASE WAIT..........\n')
# All Imports

import pandas as pd
import os
import json

import smtplib
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders



# Loading the Configuration
with open('config.json', 'r') as content:
    config = json.load(content)

# Getting Student data

students = pd.read_excel(config['data'])
students.head(2)


students_len = len(students)


files = config['files']
all_files = []
file_types = []
for cfile in files:
    all_files.append({'type':cfile['type'], 'extension': cfile['extension'], 'files':[]})

dir_files = os.listdir(config['folder'])
for currFile in dir_files:
    for cfile in all_files:
        if os.path.splitext(currFile)[1] == cfile['extension']:
            cfile['files'].append(currFile)


for cfiles in all_files:
    sorted_files = sorted(cfiles['files'], key=lambda doc: int(os.path.splitext(doc)[0]))
    students[cfiles['type']] = sorted_files[:students_len]
    file_types.append(cfiles['type'])


students.head(2)


# attach Files

def attachFiles(msg, filename):
    attachment = open(config['folder']+filename, "rb")
    p = MIMEBase('application', 'octet-stream') 
    p.set_payload((attachment).read()) 
    encoders.encode_base64(p) 
    p.add_header('Content-Disposition', "attachment; filename= %s" % filename)  
    msg.attach(p) 

# Initializing SMTP Server

s = smtplib.SMTP('smtp.gmail.com', 587)
s.starttls()
s.login(config['email'], config['password'])


with open(config['body'], 'r') as content:
    body = content.read()

def mainMail(name, mailTo, filenames):
    msg = MIMEMultipart() 
    msg['From'] = config['email']  
    msg['Subject'] = config['subject']
    msg.attach(MIMEText(body, 'plain'))
    msg['To'] = mailTo
    for filename in filenames:
        attachFiles(msg, filename)
    text = msg.as_string()
    s.sendmail(config['email'], mailTo, text)
    print(f'Mail Successfully Send to {name}!')


for index, student in students.iterrows():
    studentEmail = student['Email']
    studentName = student['Name']
    student_files = []
    for file_type in file_types:
        student_files.append(student[file_type])
    mainMail(studentName, studentEmail, student_files)
    # print(studentEmail, student_files)

s.quit()
input('\nTHANKS FOR USING IT, PRESS ANY KEY TO QUIT....')
