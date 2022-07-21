#!/usr/bin/env python
# coding: utf-8


import pandas as pd

import pathlib
from pathlib import Path
downloadPath = str(Path.home() / "Downloads") + "/"

import time


#Exchange Library: necessary to send emails with  Office365 

#https://ecederstrand.github.io/exchangelib/
from exchangelib import Credentials, Account, Configuration
from exchangelib import DELEGATE, IMPERSONATION, Account, Credentials, EWSDateTime, EWSTimeZone, Configuration, NTLM, GSSAPI, CalendarItem, Message, Mailbox, Attendee, Q, ExtendedProperty, FileAttachment, ItemAttachment, HTMLBody, Build, Version, FolderCollection
from exchangelib import Account, FileAttachment, ItemAttachment, Message

import csv




#The name should be found in your Download folder
#Based on grade report file
NameOfTheFile= "gradeReport.csv"
df = pd.read_csv (downloadPath + NameOfTheFile)




#Function to clean the data i.e remove EPFL people from your email.
def removeThisData(dataframe, columnName, itemToRemove):
    dataframe = dataframe[dataframe[columnName] != itemToRemove]
    return dataframe

#Function to get rid off certain columns
def dropColByName(dataframe, columns):
    for colum in columns:
        dataframe.drop(dataframe.filter(regex=colum).columns, axis=1, inplace=True)




#Remove EPFL cohort (e.g EPFL personel)
df = removeThisData(df, 'Cohort Name', 'EPFL')

#Remove unenrolled participants
df = removeThisData(df, 'Enrollment Status', 'unenrolled')

#Remove Students that haven't done anything
df = removeThisData(df, 'Grade', 0)


#Droping of the columns by name by the function define before
colToDrop = ["Cohort Name", "Enrollment Track", "Verification Status", "Certificate Eligible", "Certificate Type", "Enrollment Status", "Certificate Delivered", "Unnamed"]
dropColByName(df, colToDrop)





#Define first column and the last column (of the dataframe)
firstcolFull = "Grade"
lastcolFull = df.iloc[: , -1].name
start = df.columns.get_loc(firstcolFull)
end = df.columns.get_loc(lastcolFull)+1

#Replace NaN par 0
#Replace Not Available by 0
#Replace Not Attempted by -1

df = df.fillna(0)
for num in list(range(start,end)):
    nomColonne = df.columns[num]
    df[nomColonne] = df[nomColonne].replace('Not Attempted',-1)
    df[nomColonne] = df[nomColonne].replace('Not Available',0)
    df[nomColonne] = df[nomColonne].astype(float)
df.to_csv('check.csv')



#Number of modules of your mooc
mod1 = {}
mod2 = {}
mod3 = {}
mod4 = {}
mod5 = {}
mod6 = {}
mod7 = {}
mod8 = {}
mod9 = {}
mod10 = {}



#Function that determines the start and the end of a module
def sections(dataframe, mod, colStart, colEnd):
    start = dataframe.columns.get_loc(colStart)
    end = dataframe.columns.get_loc(colEnd)
    mod['start'] = start
    mod['end'] = end+1



#Define the modules (start and end) from their names.
sections(df,mod1, 'Ice Breaker (mod1IcB)', 'KeyConcept (mod1Kec)')
sections(df,mod2, 'Assess readiness (mod2Rea)', 'Stakeholders Analysis (mod2Sta)')
sections(df,mod3,  'Design&Learning Activities (mod3DLA) 1: Design', 'Lesson Plan (mod3LPP)')
sections(df,mod4,  'Intellectual Property (mod4ipq)', 'Flipped Classroom (mod4fc)')
sections(df,mod5, 'Theoretical Foundations (mod5tf)', 'Tests and Assessments (mod5ta)  (Avg)')
sections(df,mod6,  'Best practices (mod6bp)', 'Filming a video (mod6fv)')
sections(df,mod7,  'Content Publication quiz (mod7cp)', 'Flipped Classroom Sandbox (mod7fls)')
sections(df,mod8,  'Implementation various quiz (mod8iq) 1: Implementation', 'Implementation plan assignment (mod8ip)')
sections(df,mod9,  'Implementation and Motivation (mod9im)', 'Quiz (mod9q)')
sections(df,mod10,  'Design a questionnaire (mod10daq)', 'Wrap up quiz (mod10wuq)')



#Other way to define the different modules
"""
df.filter(regex="mod1", axis=1)
df.filter(regex="mod2", axis=1)
df.filter(regex="mod3", axis=1)
df.filter(regex="mod4", axis=1)
df.filter(regex="mod5", axis=1)
df.filter(regex="mod6", axis=1)
df.filter(regex="mod7", axis=1)
df.filter(regex="mod8", axis=1)
df.filter(regex="mod9", axis=1)
df.filter(regex="mod10", axis=1)
 """


#Threshold is 
threshold = 0.7
start = mod1['start']
end = mod10['end']
df['Grade'] = df['Grade'].mul(100)

for activity in list(range(start, end)):
    nomActivity = df.columns[activity]
    achieved = "<span style='color:green;''>Achieved</span>"
    notYetAchieved = "<span style='color:orange;''>Not (yet) achieved*</span>"
    notAttempted = "<span style='color:red;''>Not attempted*</span>"
    notAvailable = "<span style='color:black;''>Not available*</span>"
    NAN = "<span style='color:black;''>N/A*</span>"
    df[nomActivity] = df[nomActivity].apply(lambda x: achieved if x >= threshold else notYetAchieved if x < threshold and x>0 else notAttempted if x == 0 else notAvailable if x ==-1 else NAN )


#Those two lines could help to check the data before sending.

df = pd.read_csv ('check.csv')
df = df[df.columns.drop(list(df.filter(regex='Unnamed')))]



#Create the HTML in the mail
def creationHTML(dataframe, moduleName, module, line):
    HTML.append("<h3>"+moduleName+" </h3><ul>")
    for activity in list(range(module['start'],module['end'] )):
        HTML.append('<li>'+ dataframe.columns[activity] + ": <strong>" +  line[activity+1] + "</strong>" + '</li>')
    HTML.append('</ul>')



emailAdress = 'user@epfl.ch'
username = 'user'




#Exchange protocols to send email : EPFL 
answer = input("Enter yes or no: ") 
if answer == "yes":
    emails = df['Email']
    names = df['Username']
    #THis can be improve in order to hide the password 
    password = input("Type your password and press enter: ")
    credentials = Credentials(username, password)
    account = Account(emailAdress, credentials=credentials, autodiscover=True)
    num = 0
    HTML = []

    with open('check.csv', 'r') as csv_file:
        csv_reader = csv.reader(csv_file)
        next(csv_reader, None)  # skip the headers

        for line in csv_reader:
            creationHTML(df, 'Module 1', mod1, line)
            creationHTML(df, 'Module 2', mod2, line)
            creationHTML(df, 'Module 3', mod3, line)
            creationHTML(df, 'Module 4', mod4, line)
            creationHTML(df, 'Module 5', mod5, line)
            creationHTML(df, 'Module 6', mod6, line)
            creationHTML(df, 'Module 7', mod7, line)
            creationHTML(df, 'Module 8', mod8, line)
            creationHTML(df, 'Module 9', mod9, line)
            creationHTML(df, 'Module 10', mod10, line)
           
            #Sender
            #Sender
            nameUser = names[num]
            toEmail = emails[num]

            a = account
            m = Message(
                account=a,
                folder=a.sent,
                subject="Individual Follow up DEM2S_EN- ",
                to_recipients=[toEmail]
                #to_recipients=['email@epfl.ch'],
                #cc_recipients=['email@epfl.ch'],
            )
            
            m.body = HTMLBody("""
            <html>
                <body>
                <p>Hello <b>""" + nameUser + """,</b>
                <br>
                <p>This is your Digital Education Masterclass course progress so far:</p>
            
                """+(' '.join(HTML))+"""
              
                <p>This reflects the situation at """+' GMT on <strong>'+"""</strong>, and bears in mind that the deadline for some of this week's activities is the end of Friday. </p>

    <p>If you haven't done so, please complete any pending activities as soon as possible. Working on your course regularly gives you the best chance of completing it successfully.</p> 
    <p>If you need help with any activities, please contact your facilitator. S/he would be glad to help! Our goal is to support you and make sure you get a rewarding learning experience.</p>
<p>*Some activities may not have been corrected yet </p>
                 <p>This is an automatic email, so please do not reply to it. If you have any questions, don't hesitate to reach out to your facilitator.</p>
                <p>Best wishes,<br>Your team of EXAF facilitators</p>
                </body>
            </html>
            """)
            m.send_and_save()
            HTML = []
            print(num, toEmail)
            num +=1
            time.sleep(0.5)
        print('END')
elif answer == "no":
    print('Ok no problem')
else:
    print("Please enter yes or no.")




#It is possible to send with other protocols - this is for GMAIL (SMTP protocol)

"""
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
mail_content = '''Hello,
This is a simple mail. There is only text, no attachments are there The mail is sent using Python SMTP library.
Thank You
' ' '
#The mail addresses and password
sender_address = 'sender123@gmail.com'
sender_pass = 'xxxxxxxx'
receiver_address = 'receiver567@gmail.com'
#Setup the MIME
message = MIMEMultipart()
message['From'] = sender_address
message['To'] = receiver_address
message['Subject'] = 'A test mail sent by Python. It has an attachment.'   #The subject line
#The body and the attachments for the mail
message.attach(MIMEText(mail_content, 'plain'))
#Create SMTP session for sending the mail
session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
session.starttls() #enable security
session.login(sender_address, sender_pass) #login with mail_id and password
text = message.as_string()
session.sendmail(sender_address, receiver_address, text)
session.quit()
print('Mail Sent')

"""

