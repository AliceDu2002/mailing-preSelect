##########################################################################
# File:         mailer_invite.py                                         #
# Purpose:      Automatic send 專題說明會 invitation mails to professors   #
# Last changed: 2015/06/21                                               #
# Author:       Yi-Lin Juang (B02 學術長)                                 #
# Edited:       2021/07/01 Eleson Chuang (B08 Python大佬)                 #
#               2018/05/22 Joey Wang (B05 學術長)                         #
# Copyleft:     (ɔ)NTUEE                                                 #
##########################################################################
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr
from email.mime.image import MIMEImage
import sys
import time
import re
import configparser as cp
import os
import os.path
import csv
import email.message
from string import Template
from pathlib import Path
import json

# load config.json
with open("config.json", 'r', encoding = "utf-8") as f:
    datas = json.load(f)

# get the reciever list from json
email_list = datas["recipientList"]

# load email account info
config = cp.ConfigParser()
config.read("account.ini")  # reading sender account information
try:
    user = config["ACCOUNT"]["user"]
    pw = config["ACCOUNT"]["pw"]
except:
    print("Reading config file fail!\nPlease check the config file...")
    exit()

def connectSMTP():
    # Send the message via NTU SMTP server.
    # For students ID larger than 09
    s = smtplib.SMTP_SSL("smtps.ntu.edu.tw", 465)
    # For students ID smaller than 08 i.e. elders
    # s = smtplib.SMTP('mail.ntu.edu.tw', 587)
    s.set_debuglevel(False)
    # Uncomment this line to go through SMTP connection status.
    s.ehlo()
    if s.has_extn("STARTTLS"):
        s.starttls()
        s.ehlo()
    s.login(user, pw)
    print("SMTP Connected!")
    return s

def disconnect(server):
    server.quit()

def read_list(file_name):
    obj = list()
    with open(file_name, "r", encoding="utf-8") as f:
        for line in f:
            t = line.split()
            if t is not None:
                obj.append(t)
    return obj

def send_mail(msg, server):
    server.sendmail(msg["From"], msg["To"], msg.as_string())
    print("Sent message from {} to {}".format(msg["From"], msg["To"]))

def attach_files_METHOD1(msg):
    '''This method will attach all the files in the ./attach folder.'''
    dir_path = os.getcwd()+"/attach"
    files = os.listdir("attach")
    for f in files:  # add files to the message
        file_path = os.path.join(dir_path, f)
        attachment = MIMEApplication(open(file_path, "rb").read())
        attachment.add_header('Content-Disposition', 'inline', filename=f)
        msg.attach(attachment)

def attach_files_METHOD2(msg):
    '''Reading attachment, put file_path in args'''
    for argvs in sys.argv[1:]:
        attachment = MIMEApplication(open(str(argvs), "rb").read())
        attachment.add_header("Content-Disposition", "attachment", filename=str(argvs))
        msg.attach(attachment)

sender = "{}@ntu.edu.tw".format(user)
server = connectSMTP()
count = 0

with open(email_list, 'r', newline = '', encoding="utf-8") as csvfile:
    recipients = csv.reader(csvfile)

    for recipient in recipients:
        if count % 10 == 0 and count > 0:
            print("{} mails sent, resting...".format(count))
            time.sleep(10)  # for mail server limitation
        if count % 130 == 0 and count > 0:
            print("{} mails sent, resting...".format(count))
            time.sleep(20)  # for mail server limitation
        if count % 260 == 0 and count > 0:
            print("{} mails sent, resting...".format(count))
            time.sleep(20)  # for mail server limitation
        # from    
        msg = MIMEMultipart("alternative")
        FROM = datas["from"]
        msg['From'] = formataddr((Header(FROM, 'utf-8').encode(), sender))
        # subject
        msg["Subject"] = datas["subject"]

        # letter content
        template = Template(Path(datas["letter"]).read_text( encoding="utf-8") )
        # substitute
        s = list(datas["substitute"].items())
        substitute = {}
        for key, value in s:
            substitute[key] = recipient[value]
        body = template.substitute(substitute)
        msg.attach(MIMEText(body, "html"))

        # ./attach folder METHOD
        attach_files_METHOD1(msg)

        # sys.argv METHOD
        # attach_files_METHOD2(msg)

        if datas["email"]["@ntu.edu.tw"]:
            msg["To"] = (recipient[datas["email"]["index"]] + "@ntu.edu.tw")
        else: 
            msg["To"] = (recipient[datas["email"]["index"]])
        send_mail(msg, server)
        count += 1

    disconnect(server)
    print("{} mails sent. Exiting...".format(count))
