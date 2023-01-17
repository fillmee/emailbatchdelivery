import tkinter as tk
from tkinter import filedialog
import tkinter.messagebox as tkmsg
import time
import math
import senddelivery
import threading
import pandas as pd
import datetime

#送信先リストを取得
def set_deliveryfile():
    filetype = [("", "*xlsx")]
    deliveryfile_path = tk.filedialog.askopenfilename(filetype=filetype,initialdir='')
    deliveryfile_label['text'] = deliveryfile_path

#送信文章を保管     
def set_sendfile():
    if(deliveryfile_label['text'] != ''):
        filetype = [("", "*txt")]
        sendfile_path = tk.filedialog.askopenfilename(filetype=filetype,initialdir='')
        sendfile_label['text'] = sendfile_path
        msg_text = senddelivery.getmsg(sendfile_label['text'])
        delivery_array = senddelivery.get(deliveryfile_label['text'])
        try:
            bodymsg = msg_text.format(
                CORPNAME=delivery_array[0][4],
                PICNAME=delivery_array[0][5],
                MYNAME=delivery_array[0][6],
                CONTEXT01=delivery_array[0][7],
                CONTEXT02=delivery_array[0][8],
                CONTEXT03=delivery_array[0][9],
                CONTEXT04=delivery_array[0][10],
                CONTEXT05=delivery_array[0][11],
                )
            preview_label['text'] = bodymsg
        except KeyError as e:
            preview_label['text'] = '差し込みフィードが正しくありません\n'+str(e)+'\n正しい値をセットしてからもう一度ファイルの読み込みを行って下さい'
    else:
        tkmsg.showinfo(message="事前に送信リストを設定してください")

#送信
result_true = []
result_false = []

def count():
    delivery_array = senddelivery.get(deliveryfile_label['text'])
    len(delivery_array) #合計
    result_sum = 0
    while result_sum < len(delivery_array):
        result_sum = len(result_true) + len(result_false)
        preview_label['text'] = f'{str(len(delivery_array))}メッセージを送信中です。\nアプリケーションを終了しないでください\n'+f'送信完了:{str(len(result_true))}件 / 送信失敗:{str(len(result_false))}件'   
    preview_label['text'] = 'すべての送信が完了しました\n'+f'送信完了:{str(len(result_true))}件 / 送信失敗:{str(len(result_false))}件'

def write_log():
    backupfilename_value = f'log_送信結果_{get_backupfilename()}.xlsx'
    today = datetime.datetime.now()
    hoge = today.strftime('%Y年%m月%d日 %H:%M:%S')
    df_true = pd.DataFrame({'送信先':result_true,'送信結果':'成功','送信時刻':hoge})
    df_false = pd.DataFrame({'送信先':result_false,'送信結果':'失敗','送信時刻':hoge})
    df = pd.concat([df_true,df_false])
    #df.to_csv(backupfilename_value,encoding='cp932',mode='w')
    df.to_excel(backupfilename_value)

def copyfile_log():
    backupfilename_value = get_backupfilename()
    with open(f'log_テキスト_{backupfilename_value}.txt','w') as file:
        file.write(senddelivery.getmsg(sendfile_label['text']))
    delivery_array = senddelivery.get(deliveryfile_label['text'])
    df = pd.DataFrame(delivery_array,columns=[
        '送信先','送信元','BCC','件名','企業名','担当者名','私の名前','差し込み','差し込み','差し込み','差し込み','差し込み'
        ])
    df.to_excel(f'log_配信リスト_{backupfilename_value}.xlsx')
        
def submit():
    preview_label['text'] = '処理を開始します'
    copyfile_log()
    count_th = threading.Thread(target=count)
    submit_run_th = threading.Thread(target=submit_run)
    submit_run_th.start()
    count_th.start()

def get_backupfilename():
    backupfilename = backupfilename_entry.get()
    if(backupfilename==""):
        return 'DefaultBackupFileName'
    else:
        return backupfilename
      
def submit_run():
    msg_text = senddelivery.getmsg(sendfile_label['text'])
    delivery_array = senddelivery.get(deliveryfile_label['text'])
    for sendmsg in delivery_array:
        bodymsg = msg_text.format(
            CORPNAME=sendmsg[4],
            PICNAME=sendmsg[5],
            MYNAME=sendmsg[6],
            CONTEXT01=sendmsg[7],
            CONTEXT02=sendmsg[8],
            CONTEXT03=sendmsg[9],
            CONTEXT04=sendmsg[10],
            CONTEXT05=sendmsg[11],
            )
        if ('@' in str(sendmsg[2])) == False:
            bcc = ''
        else:
            bcc = sendmsg[2]
        resulet = senddelivery.sendmail(sendmsg[0],sendmsg[1],bcc,bodymsg,sendmsg[3])
        if(resulet[0]==True):
            result_true.append(resulet[1])
        else:
            result_false.append(resulet[1])
        write_log()

#GUIの組み込み  
root = tk.Tk()
root.title('一括メール送信')
calc_geometry = (f'800x{math.trunc(root.winfo_screenheight()*0.8)}+25+25')
root.geometry(calc_geometry)
root.resizable(width=0,height=math.trunc(root.winfo_screenheight()*0.9))

#row 0　送信先のセット
set_deliveryfile_btn = tk.Button(text=u'送信先をセットする',command=set_deliveryfile,width=30)
set_deliveryfile_btn.grid(row=0,column=0,padx=10,pady=10,ipadx=10,ipady=10)

deliveryfile_label = tk.Label(text=u'',width=70,background='#ffffff')
deliveryfile_label.grid(row=0,column=1,padx=10,pady=10,ipadx=10,ipady=10)

#row 1　送信文のセット
set_sendfile_btn = tk.Button(text=u'送信文をセットする',command=set_sendfile,width=30)
set_sendfile_btn.grid(row=1,column=0,padx=10,ipadx=10,ipady=10)
sendfile_label = tk.Label(text=u'',width=70,background='#ffffff')
sendfile_label.grid(row=1,column=1,padx=10,ipadx=10,ipady=10)

backupfilename_label = tk.Label(text=u'バックアップファイル名を入力してください。※拡張子不要',anchor='w', justify='left')
backupfilename_label.grid(row=2,column=0,columnspan=2,padx=10,pady=0,ipadx=10,ipady=10,sticky=tk.W+tk.E)

backupfilename_entry = tk.Entry()
backupfilename_entry.grid(row=3,column=0,columnspan=2,padx=10,pady=0,ipadx=10,ipady=10,sticky=tk.W+tk.E)
#row 3
submit = tk.Button(text=u'送信する',command=submit).grid(row=4,column=0,columnspan=2,padx=10,pady=10,ipadx=10,ipady=10,sticky=tk.W+tk.E)

#row 4
preview_label = tk.Label(text=u'',anchor='e', justify='left')
preview_label.grid(row=5,column=0,columnspan=2,padx=10,pady=10,ipadx=10,ipady=10)

root.mainloop()