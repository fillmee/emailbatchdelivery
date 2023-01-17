import pandas as pd
import smtplib
import ssl
from email.mime.text import MIMEText
from email.message import EmailMessage

#送信先を取得して出力 [送信先address、送信元address、BCCaddress（空白可）]
def get(path):
    df = pd.read_excel(path)
    df_array = df.to_numpy()
    return df_array

#送信文章を取得して出力
def getmsg(path):
    msg_tmp = open(path,encoding="utf-8")
    msg = msg_tmp.read()
    msg_tmp.close()
    return msg

#メール送信スクリプト
def sendmail(msg_to,msg_account,msg_bcc,msg_message,msg_subject):
    host = 'af127.secure.ne.jp'
    port = 465
    accountpass = 'UoCcq5TN'
    timeout_time = 10  
    cset = 'utf-8'
    msg = MIMEText(msg_message,'plain',cset)
    msg["Subject"] = msg_subject
    msg["To"] = msg_to
    msg["From"] = msg_account
    msg["Bcc"] = msg_bcc
    try:
        server = smtplib.SMTP_SSL(host,port,timeout=timeout_time,context=ssl.create_default_context())
        #server.starttls()
        server.login(msg_account,accountpass)
        server.send_message(msg)
        server.quit()
        return [True,msg_to]
    except:
        return [False,msg_to]

