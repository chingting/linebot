import requests
import re
import random
import os
import configparser
from bs4 import BeautifulSoup
from flask import Flask, request, abort
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
from datetime import datetime, timedelta

from linebot import (
    LineBotApi, WebhookHandler
)
from linebot.exceptions import (
    InvalidSignatureError
)
from linebot.models import *

app = Flask(__name__)
config = configparser.ConfigParser()
config.read("config.ini")

line_bot_api = LineBotApi(config['line_bot']['Channel_Access_Token'])
handler = WebhookHandler(config['line_bot']['Channel_Secret'])
client_id = config['imgur_api']['Client_ID']
client_secret = config['imgur_api']['Client_Secret']
album_id = config['imgur_api']['Album_ID']
API_Get_Image = config['other_api']['API_Get_Image']


@app.route("/callback", methods=['POST'])
def callback():
    # get X-Line-Signature header value
    signature = request.headers['X-Line-Signature']

    # get request body as text
    body = request.get_data(as_text=True)
    # print("body:",body)
    app.logger.info("Request body: " + body)

    # handle webhook body
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)

    return 'ok'

def over_time(count):   #加班
    scope = ["https://spreadsheets.google.com/feeds"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('auth.json', scope)
    client = gspread.authorize(creds)
    sheet = client.open('test').get_worksheet(0)
    content=sheet.cell(1, count).value
    return content

def off_time(count):   #請假
    scope = ["https://spreadsheets.google.com/feeds"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('auth.json', scope)
    client = gspread.authorize(creds)
    sheet = client.open('test').get_worksheet(0)
    content=sheet.cell(2, count).value
    return content

def write_in(answers):   #加班存入
    scope = ["https://spreadsheets.google.com/feeds"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('auth.json', scope)
    client = gspread.authorize(creds)
    sheet = client.open('test').get_worksheet(1)
    sheet.insert_row(answers, 2)
    print('done')

def write_to(answers):    #請假存入
    scope = ["https://spreadsheets.google.com/feeds"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('auth.json', scope)
    client = gspread.authorize(creds)
    sheet = client.open('test').get_worksheet(3)
    sheet.insert_row(answers, 2)
    print('done')

def search_over(count):   #查詢加班
    scope = ["https://spreadsheets.google.com/feeds"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('auth.json', scope)
    client = gspread.authorize(creds)
    sheet = client.open('test').get_worksheet(0)
    content = sheet.cell(3, count).value
    return content

def search_off(count):   #查詢請假
    scope = ["https://spreadsheets.google.com/feeds"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('auth.json', scope)
    client = gspread.authorize(creds)
    sheet = client.open('test').get_worksheet(0)
    content = sheet.cell(4, count).value
    return content

def general():
    scope = ["https://spreadsheets.google.com/feeds"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('auth.json', scope)
    client = gspread.authorize(creds)
    return client

def love(name, year, breaks):   #看今年已請假別與天數
    client = general()
    sheet = client.open('test').get_worksheet(5)
    col = sheet.col_values(1)
    for p in range(len(col)):
        p = p + 1
        if (sheet.cell(p, 1).value) == name and (sheet.cell(p, 2).value) == year:
            if breaks=="事假":
                c = (sheet.cell(p, 3).value)
            if breaks == "公傷病假":
                c = (sheet.cell(p, 4).value)
            if breaks == "普通傷病假":
                c = (sheet.cell(p, 5).value)
            if breaks == "生理假":
                c = (sheet.cell(p, 6).value)
            if breaks == "喪假":
                c = (sheet.cell(p, 7).value)
            if breaks == "婚假":
                c = (sheet.cell(p, 8).value)
            if breaks == "特休":
                c = (sheet.cell(p, 9).value)
            if breaks == "補休":
                c = (sheet.cell(p, 10).value)
            if breaks == "公假":
                c = (sheet.cell(p, 11).value)
            if 0 <= int(c) < 9:
                a = ("%s: %s小時") % (breaks, c)
            if int(c)==9:
                a = ("%s: 1天") % (breaks)
            if int(c) > 9:
                dayt = int(c) // 9
                hourt = int(c) % 9
                a = ("%s: %s天%s小時") % (breaks, dayt, hourt)
    return a



def love1(name, year, month, breaks):   #看某月出勤狀況
    client = general()
    sheet = client.open('test').get_worksheet(4)
    col = sheet.col_values(1)
    for p in range(len(col)):
        p = p + 1
        if (sheet.cell(p, 1).value) == name and (sheet.cell(p, 2).value) == year:
            if (sheet.cell(p, 3).value) == month:
                if breaks == "事假":
                    c = (sheet.cell(p, 4).value)
                if breaks == "公傷病假":
                    c = (sheet.cell(p, 5).value)
                if breaks == "普通傷病假":
                    c = (sheet.cell(p, 6).value)
                if breaks == "生理假":
                    c = (sheet.cell(p, 7).value)
                if breaks == "喪假":
                    c = (sheet.cell(p, 8).value)
                if breaks == "婚假":
                    c = (sheet.cell(p, 9).value)
                if breaks == "特休":
                    c = (sheet.cell(p, 10).value)
                if breaks == "補休":
                    c = (sheet.cell(p, 11).value)
                if breaks == "公假":
                    c = (sheet.cell(p, 12).value)
                if 0 <= int(c) < 9:
                    a = ("%s: %s小時") % (breaks, c)
                if int(c) == 9:
                    a = ("%s: 1天") % (breaks)
                if int(c) > 9:
                    dayt = int(c) // 9
                    hourt = int(c) % 9
                    a = ("%s: %s天%s小時") % (breaks, dayt, hourt)
    return a


@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    print("event.reply_token:", event.reply_token)
    print("event.message.text:", event.message.text)
    user_id = json.loads(str(event.source))['userId']

    #建一個虛擬的資料夾Users，裡面暫存每個使用者的檔案，資料夾不能為空，所以先隨便放一個345的檔案。
    #每個檔案命名以使用者的user id為主。檔案存成json檔
    users = os.listdir("Users")  #找出資料夾Users裡所有的檔案
    answers =[]
    name = user_id + ".json"
    path = "Users/" + name


    if name in users:
        print("first")
        with open(path, encoding="utf-8") as json_data:
            d = json.load(json_data)
        count = d['count']
        answers = d['answers']

    if name not in users:
        print("second")
        d = {"user_id": user_id, "count": 0, "answers": []}  # 記得要建answer的部分
        with open(path, 'w', encoding='UTF-8') as f:
            json.dump(d, f, ensure_ascii=False)
        count = d['count']
        answers = d['answers']

    if event.message.text == "加班申請單" or event.message.text == "請假申請單" or event.message.text == "加班查詢" or event.message.text == "請假查詢":
        if count > 0:
            d['count'] = 0
            d['answers'] = []
            answers=d['answers']
        print("third")
        print(d)
        answers.append(event.message.text)
        d['count'] = 1
        d['answers'] = answers
        with open(path, 'w') as f:
            json.dump(d, f, ensure_ascii=False)
        client = general()
        sheet = client.open('test').get_worksheet(2)
        col = sheet.col_values(2)
        if user_id not in col:
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="目前人事資料表裡沒有你的line使用者id資料，你的line使用者id為"+user_id+"請提供給相關負責人以便於建檔"))
            d['count'] = 0
            d['answers'] = []
            with open(path, 'w') as f:
                json.dump(d, f, ensure_ascii=False)
            return 0
        if user_id in col:
            value_index = col.index(user_id) + 1
            u_name = sheet.cell(value_index, 1).value
            print(u_name)
            buttons_template = TemplateSendMessage(
                alt_text='名字 template',
                template=ButtonsTemplate(
                text='請選擇你的名字或手動輸入',
                actions=[
                    MessageTemplateAction(
                        label=u_name,
                        text=u_name
                    )
                ]
            )
        )
        line_bot_api.reply_message(event.reply_token, buttons_template)
        return 0





    if count != 0:
        if answers[0] == '請假查詢':
            print("forth_off_1")
            if count == 1:
                client = general()
                sheet = client.open('test').get_worksheet(2)
                col = sheet.col_values(1)
                if event.message.text not in col:
                    event.message.text = "none"

            if count==2:
                if event.message.text!="某月出勤狀況" and event.message.text!="今年已請假別與天數":
                    event.message.text="wrong"


            if count==3:
                print(event.message.text)
                if int(event.message.text) not in range(2017,2019):
                    event.message.text="wyear"

            if count==4:
                print(event.message.text)
                event.message.text=int(event.message.text)
                if event.message.text not in range(1,13):
                    event.message.text="wmonth"

            if event.message.text == "none":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="查無此人，請重新輸入"))

            if event.message.text == "wrong":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="格式輸入錯誤，請重新輸入"))

            if event.message.text == "wmonth":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="輸入錯誤，請輸入1~12的數字"))

            if event.message.text == "wyear":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="輸入錯誤，請輸入四位數字，預設初始年份為2017，目前可以查至2018年"))

            if event.message.text == "今年已請假別與天數":
                now = datetime.now().strftime('%Y-%m-%dT%H:%M')
                name = answers[1]
                year = "%s" % (datetime.strptime(now, "%Y-%m-%dT%H:%M")).year
                f = love(name, year, "事假")
                f1 = love(name, year, "公傷病假")
                f2 = love(name, year, "普通傷病假")
                f3 = love(name, year, "生理假")
                f4 = love(name, year, "喪假")
                f5 = love(name, year, "婚假")
                f6 = love(name, year, "特休")
                f7 = love(name, year, "補休")
                f8 = love(name, year, "公假")
                ans="%s總共請了\n%s\n%s\n%s\n%s\n%s\n%s\n%s\n%s\n%s" % (
                                year, f, f1, f2, f3, f4, f5, f6, f7, f8)
                line_bot_api.push_message(user_id, TextSendMessage(text=ans))
                d['count'] = 0
                d['answers'] = []
                with open(path, 'w') as f:
                    json.dump(d, f, ensure_ascii=False)
                return 0

            if type(event.message.text)==int:
                answers.append(str(event.message.text))
                name = answers[1]
                year = answers[3]
                month = answers[4]
                if year=="2017" and int(month) not in range(11,13):
                    line_bot_api.push_message(user_id, TextSendMessage(text="查無此月份的資料"))
                    d['count'] = 0
                    d['answers'] = []
                    with open(path, 'w') as f:
                        json.dump(d, f, ensure_ascii=False)
                    return 0
                c = love1(name, year, month, "事假")
                c1 = love1(name, year, month, "公傷病假")
                c2 = love1(name, year, month, "普通傷病假")
                c3 = love1(name, year, month, "生理假")
                c4 = love1(name, year, month, "喪假")
                c5 = love1(name, year, month, "婚假")
                c6 = love1(name, year, month, "特休")
                c7 = love1(name, year, month, "補休")
                c8 = love1(name, year, month, "公假")
                ans = "%s年%s月的請假狀況\n%s\n%s\n%s\n%s\n%s\n%s\n%s\n%s\n%s" % (
                    year,month, c, c1, c2, c3, c4, c5, c6, c7, c8)
                line_bot_api.push_message(user_id, TextSendMessage(text=ans))
                d['count'] = 0
                d['answers'] = []
                with open(path, 'w') as f:
                    json.dump(d, f, ensure_ascii=False)
                return 0

            answers.append(event.message.text)
            print(answers)
            if "none" in answers:
                answers.remove("none")
            if "wrong" in answers:
                answers.remove("wrong")
            if "wmonth" in answers:
                answers.remove("wmonth")
            if "wyear" in answers:
                answers.remove("wyear")
            print(answers)
            d['answers'] = answers
            print(d)
            count += 1
            print(count)

            print("next")
            content = search_off(count)
            print(content)
            d['count'] = count
            with open(path, 'w') as f:
                json.dump(d, f, ensure_ascii=False)

            if content == "請問要查詢哪個年份的請假狀況?":
                line_bot_api.reply_message(
                    event.reply_token,
                    TextSendMessage(text=content))
                return 0

            if content == "請問要查詢哪個月份的請假狀況?":
                line_bot_api.reply_message(
                    event.reply_token,
                    TextSendMessage(text=content))
                return 0


            if content=="查詢選項":
                buttons_template = TemplateSendMessage(
                    alt_text='查詢 template',
                    template=ButtonsTemplate(
                        text='請選擇',
                        actions=[
                            MessageTemplateAction(
                                label='某月出勤狀況',
                                text='某月出勤狀況'
                            ),
                            MessageTemplateAction(
                                label='今年已請假別與天數',
                                text='今年已請假別與天數'
                            )
                        ]
                    )
                )
                line_bot_api.reply_message(event.reply_token, buttons_template)
                return 0

        if answers[0] == '加班查詢':
            print("forth_off_1")
            if count == 1:
                client = general()
                sheet = client.open('test').get_worksheet(2)
                col = sheet.col_values(1)
                if event.message.text not in col:
                    event.message.text = "none"

            if count==2:
                print(event.message.text)
                if int(event.message.text) not in range(2017,2019):
                    event.message.text="wyear"

            if count==3:
                print(event.message.text)
                if int(event.message.text) not in range(1,13):
                    event.message.text="wmonth"

            if count ==4:
                if event.message.text!="累積加班時數" and event.message.text!="完整加班紀錄":
                    event.message.text="wrong"

            if event.message.text == "none":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="查無此人，請重新輸入"))

            if event.message.text == "wmonth":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="輸入錯誤，請輸入1~12的數字"))

            if event.message.text == "wyear":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="輸入錯誤，請輸入四位數字，預設初始年份為2017，目前可以查至2018年"))

            if event.message.text == "wrong":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="格式輸入錯誤，請重新輸入"))

            if event.message.text =="累積加班時數":
                print("here")
                client=general()
                sheet = client.open('test').get_worksheet(6)
                col = sheet.col_values(2)
                year=answers[2]
                month=answers[3]
                name=answers[1]
                for jr in range(len(col)):
                    jr = jr + 1
                    if (sheet.cell(jr, 1).value) == name and (sheet.cell(jr, 2).value) == year:  # 找出名字一樣且審核通過的
                        if (sheet.cell(jr, 3).value) == month:
                            total= (sheet.cell(jr, 4).value)
                            total_d = (sheet.cell(jr, 5).value)
                            total_m = (sheet.cell(jr, 6).value)
                            ans = "%s年%s月份\n姓名:%s\n累積加班時數:%s小時\n\n=====轉換成=====\n補修時數:%s小時\n加班費時數:%s小時" % (
                                year, month, name, total, total_d, total_m)
                        if year=="2017" and month not in ["11", "12"]:
                            ans = "此月份無任何加班紀錄"
                        #else:
                            #ans = "此月份無任何加班紀錄"
                line_bot_api.push_message(user_id, TextSendMessage(text=ans))
                d['count'] = 0
                d['answers'] = []
                with open(path, 'w') as f:
                    json.dump(d, f, ensure_ascii=False)
                return 0

            if event.message.text =="完整加班紀錄":
                client=general()
                sheet = client.open('test').get_worksheet(1)
                col = sheet.col_values(10)
                col2 = sheet.col_values(9)
                print(col)
                year = answers[2]
                month=answers[3]
                print(month)
                print(type(month))
                name=answers[1]
                b=[]
                g=[] ######
                for j in range(len(col)):
                    j = j + 1
                    if (sheet.cell(j, 2).value) == name and (sheet.cell(j, 11).value) == "是":  # 找出名字一樣且審核通過的
                        time = (sheet.cell(j, 3).value)
                        date = time.split("T")[0]
                        hour = (sheet.cell(j, 7).value)
                        if int(sheet.cell(j, 9).value) == int(year):
                            g.append(sheet.cell(j, 10).value)   #########
                            if int(sheet.cell(j, 10).value) == int(month):   #####
                                b.append(("加班日期: %s，加班時數: %s" % (date, hour)))
                                ans=("\n".join(b))
                        #if year not in col2 or month not in col:  #如果我選2018和1，但裡面沒有資料的話
                            #ans="此月份無任何加班紀錄"
                if month not in g:
                    ans="此月份無任何加班紀錄"
                line_bot_api.push_message(user_id, TextSendMessage(text=ans))
                #count = count + 1
                #d['count'] = count
                #with open(path, 'w') as f:
                #    json.dump(d, f, ensure_ascii=False)
                d['count'] = 0
                d['answers'] = []
                with open(path, 'w') as f:
                    json.dump(d, f, ensure_ascii=False)
                return 0


            answers.append(event.message.text)
            print(answers)
            if "none" in answers:
                answers.remove("none")
            if "wmonth" in answers:
                answers.remove("wmonth")
            if "wrong" in answers:
                answers.remove("wrong")
            if "wyear" in answers:
                answers.remove("wyear")
            print(answers)
            d['answers'] = answers
            print(d)
            count += 1
            print(count)

            print("next")
            content = search_over(count)
            print(content)
            d['count'] = count
            with open(path, 'w') as f:
                json.dump(d, f, ensure_ascii=False)

            if content=="請問要查詢哪個年份的加班資訊?":
                line_bot_api.reply_message(
                    event.reply_token,
                    TextSendMessage(text=content))
                return 0

            if content=="請問要查詢哪個月份的加班資訊?":
                line_bot_api.reply_message(
                    event.reply_token,
                    TextSendMessage(text=content))
                return 0

            if content=="加班細節":
                buttons_template = TemplateSendMessage(
                    alt_text='查詢 template',
                    template=ButtonsTemplate(
                        text='請選擇',
                        actions=[
                            MessageTemplateAction(
                                label='累積加班時數',
                                text='累積加班時數'
                            ),
                            MessageTemplateAction(
                                label='完整加班紀錄',
                                text='完整加班紀錄'
                            )
                        ]
                    )
                )
                line_bot_api.reply_message(event.reply_token, buttons_template)
                return 0






        if answers[0]=="請假申請單":            ########################################################################
            print("forth_off")
            #######################################################
            if count == 1:
                client = general()
                sheet = client.open('test').get_worksheet(2)
                col = sheet.col_values(1)
                if event.message.text not in col:
                    event.message.text = "none"

            if count == 2:
                if event.message.text not in ["事假", "公傷病假", "普通傷病假", "生理假", "喪假", "婚假", "特休", "補休", "公假"]:
                    event.message.text = "wrong"


            if count == 3:
                client = general()
                sheet = client.open('test').get_worksheet(2)
                cols = sheet.col_values(1)
                if event.message.text not in cols:
                    event.message.text = "none"
                if event.message.text in cols:
                    client = general()
                    sheet = client.open('test').get_worksheet(2)
                    col = sheet.col_values(2)
                    col2 = sheet.col_values(5)
                    print(len(col2))
                    value_index = col.index(user_id) + 1
                    print(value_index)
                    depart = sheet.cell(value_index, 5).value
                    anse = []
                    for i in range(len(col2)):
                        i = i + 1
                        print(i)
                        if i == value_index:
                            continue
                        if sheet.cell(i, 5).value == depart:
                            print(i)
                            print(sheet.cell(i, 1).value)
                            anse.append(sheet.cell(i, 1).value)
                    if event.message.text not in anse:
                        event.message.text = "else"

            if count == 4 or count == 5:
                try:
                    datetime.strptime(event.message.text, "%Y-%m-%dT%H:%M")
                    hour = "%s" % (datetime.strptime(event.message.text, "%Y-%m-%dT%H:%M")).hour
                    minute = "%s" % (datetime.strptime(event.message.text, "%Y-%m-%dT%H:%M")).minute
                    print(hour)
                    if 0 <= (datetime.strptime(event.message.text, "%Y-%m-%dT%H:%M")).weekday() <= 5:
                        if 0 <= int(hour) < 9 or 18 < int(hour) < 24:
                            event.message.text = "error"
                        if int(hour) == 18 and int(minute) != 0:
                            event.message.text = "error"
                    else:
                        print("周末")
                        event.message.text = "off"
                except ValueError:
                    print("格式輸入錯誤")
                    event.message.text = "wrong"

            if count == 6:
                if event.message.text != "確定" and event.message.text != "請假申請單":
                    print("格式輸入錯誤")
                    event.message.text = "wrong"



            if event.message.text == "none":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="查無此人，請重新輸入"))

            if event.message.text == "off":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="此時間為周末時間，應該沒有請假的必要，有疑問請詢問人事處，並在下方重新輸入"))

            if event.message.text == "error":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="輸入的資料不正確，請假時間應為平日上班時間，請重新輸入"))

            if event.message.text == "wrong":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="格式輸入錯誤，請重新輸入"))

            if event.message.text == "else":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="請輸入同部門的同仁"))

            answers.append(event.message.text)
            if "wrong" in answers:
                answers.remove("wrong")
            if "off" in answers:
                answers.remove("off")
            if "error" in answers:
                answers.remove("error")
            if "none" in answers:
                answers.remove("none")
            if "else" in answers:
                answers.remove("else")
            d['answers'] = answers
            count += 1

            if count >= 7:
                print("over")
                time = str((datetime.strptime(answers[5], "%Y-%m-%dT%H:%M") - datetime.strptime(answers[4], "%Y-%m-%dT%H:%M")))
                e = time.split(":")[0]
                if "day" in time:
                    if time.split(":")[1] == "00":
                        day = (time.split(",")[0]).split(" ")[0]
                        h = e.split(",")[1].replace(" ", "")
                        if h != "9":
                            if int(h) > 15:
                                h = int(h) - 15
                            j = int(day) * 9 + int(h)
                        if h == "9":
                            day = int(day) + 1
                            j = day * 9
                        if h == "0":
                            j = int(day) * 9
                    else:
                        day = (time.split(",")[0]).split(" ")[0]
                        h = e.split(",")[1].replace(" ", "")
                        z = int(h) + 1
                        if int(z) > 15:
                            z = int(z) - 15
                            j = int(day) * 9 + int(z)
                        elif z == "9":
                            day = int(day) + 1
                            j = int(day) * 9
                        else:
                            j = int(day) * 9 + int(z)
                else:
                    if time.split(":")[1] == "00":
                        e = time.split(":")[0]
                        j = e
                        if int(e) >= 15:
                            e = int(e) - 15
                            j = e
                        if int(e) == 9:
                            j = e
                    else:
                        e = time.split(":")[0]
                        s = int(e) + 1
                        j = s
                        if int(s) >= 15:
                            s = int(s) - 15
                            j = s
                        if int(s) == 9:
                            j = s
                answers.append(j)
                if int(j) < 9:
                    m = "%s小時" % j
                else:
                    if int(j) % 9 == 0:
                        d2 = int(j) // 9
                        m = "%s天" % d2
                    else:
                        d1 = int(j) // 9
                        h1 = int(j) % 9
                        m = "%s天%s小時" % (d1, h1)
                unique = (datetime.now() + timedelta(hours=8)).strftime('%Y_%m_%d_%H_%M_%S')
                name = d['answers'][1]
                id = name + "_" + unique
                print(id)
                ###############################################
                client = general()
                sheet = client.open('test').get_worksheet(2)
                col = sheet.col_values(1)
                name =answers[1]
                if name in col:
                    value_index = col.index(name) + 1
                    boss_id = sheet.cell(value_index, 4).value
                    print(boss_id)

                ###############################################
                name = d['answers'][1]
                typ = d['answers'][2]
                name_1 = d['answers'][3]
                start = d['answers'][4]
                end = d['answers'][5]
                ans = "請假單\n請假人姓名:%s\n請假類型:%s\n請假起始時間:%s\n請假結束時間:%s\n總計請假時數:%s\n請問是否核准請假申請?" % (
                    name, typ, name_1, start, end, m)
                confirm_template_message = TemplateSendMessage(
                    alt_text='請假申請審核通知',
                    template=ConfirmTemplate(
                        text=ans,
                        actions=[
                            PostbackTemplateAction(
                                label='是',
                                data=id + "_" + "yes"
                            ),
                            PostbackTemplateAction(
                                label='否',
                                data=id + "_" + "no"
                            )
                        ]
                    )
                )
                line_bot_api.push_message(boss_id, confirm_template_message)
                line_bot_api.push_message(user_id, TextSendMessage(text='已將申請單遞交給相關人員，請靜待審核通知，謝謝'))

                answers.append(id)
                yeara = (datetime.strptime(answers[4], "%Y-%m-%dT%H:%M")).year  ####
                montha = (datetime.strptime(answers[4], "%Y-%m-%dT%H:%M")).month  ####
                answers.append(yeara)
                answers.append(montha)
                d['count'] = 0
                d['answers'] = []
                with open(path, 'w') as f:
                    json.dump(d, f, ensure_ascii=False)
                write_to(answers)
                return 0

            print("next")
            print(count)
            content = off_time(count)
            d['count'] = count
            with open(path, 'w') as f:
                json.dump(d, f, ensure_ascii=False)

            if content == "請選擇你的名字或手動輸入":
                client = general()
                sheet = client.open('test').get_worksheet(2)
                col = sheet.col_values(2)
                if user_id in col:
                    value_index = col.index(user_id) + 1
                    u_name = sheet.cell(value_index, 1).value
                    print(u_name)
                buttons_template = TemplateSendMessage(
                    alt_text='名字 template',
                    template=ButtonsTemplate(
                        text='請選擇你的名字或手動輸入',
                        actions=[
                            MessageTemplateAction(
                                label=u_name,
                                text=u_name
                            )
                        ]
                    )
                )
                line_bot_api.reply_message(event.reply_token, buttons_template)
                return 0

            if content == "指定代理人是?":
                client = general()
                sheet = client.open('test').get_worksheet(2)
                col = sheet.col_values(2)
                col2 = sheet.col_values(5)
                print(len(col2))
                if user_id in col:
                    value_index = col.index(user_id) + 1
                    print(value_index)
                    depart = sheet.cell(value_index, 5).value
                    anse = []
                    for i in range(len(col2)):
                        i = i + 1
                        print(i)
                        if i == value_index:
                            continue
                        if sheet.cell(i, 5).value == depart:
                            print(i)
                            print(sheet.cell(i, 1).value)
                            anse.append(sheet.cell(i, 1).value)
                    print(anse)
                if len(anse) == 2:
                    buttons_template = TemplateSendMessage(
                        alt_text='名字 template',
                        template=ButtonsTemplate(
                            text='指定代理人是?',
                            actions=[
                                MessageTemplateAction(
                                    label=anse[0],
                                    text=anse[0]
                                ),
                                MessageTemplateAction(
                                    label=anse[1],
                                    text=anse[1]
                                )
                            ]
                        )
                    )
                    line_bot_api.reply_message(event.reply_token, buttons_template)
                    return 0

                if len(anse) == 3:
                    buttons_template = TemplateSendMessage(
                        alt_text='名字 template',
                        template=ButtonsTemplate(
                            text='指定代理人是?',
                            actions=[
                                MessageTemplateAction(
                                    label=anse[0],
                                    text=anse[0]
                                ),
                                MessageTemplateAction(
                                    label=anse[1],
                                    text=anse[1]
                                ),
                                MessageTemplateAction(
                                    label=anse[2],
                                    text=anse[2]
                                )
                            ]
                        )
                    )
                    line_bot_api.reply_message(event.reply_token, buttons_template)
                    return 0

                if len(anse) == 4:
                    buttons_template = TemplateSendMessage(
                        alt_text='名字 template',
                        template=ButtonsTemplate(
                            text='指定代理人是?',
                            actions=[
                                MessageTemplateAction(
                                    label=anse[0],
                                    text=anse[0]
                                ),
                                MessageTemplateAction(
                                    label=anse[1],
                                    text=anse[1]
                                ),
                                MessageTemplateAction(
                                    label=anse[2],
                                    text=anse[2]
                                ),
                                MessageTemplateAction(
                                    label=anse[3],
                                    text=anse[3]
                                )
                            ]
                        )
                    )
                    line_bot_api.reply_message(event.reply_token, buttons_template)
                    return 0

            if content == "請問要請什麼假?":
                print("yo!")
                carousel_template_message = TemplateSendMessage(
                    alt_text='Carousel template',
                    template=CarouselTemplate(
                        columns=[
                            CarouselColumn(
                                text='請問要請什麼假?',
                                actions=[
                                    PostbackTemplateAction(
                                        label='事假',
                                        data='事假'
                                    ),
                                    PostbackTemplateAction(
                                        label='公傷病假',
                                        data='公傷病假'
                                    ),
                                    PostbackTemplateAction(
                                        label='普通傷病假',
                                        data='普通傷病假'
                                    )
                                ]
                            ),
                            CarouselColumn(
                                text=' ',
                                actions=[
                                    PostbackTemplateAction(
                                        label='生理假',
                                        data='生理假'
                                    ),
                                    PostbackTemplateAction(
                                        label='喪假',
                                        data='喪假'
                                    ),
                                    PostbackTemplateAction(
                                        label='婚假',
                                        data='婚假'
                                    )
                                ]
                            ),
                            CarouselColumn(
                                text=' ',
                                actions=[
                                    PostbackTemplateAction(
                                        label='特休',
                                        data='特休'
                                    ),
                                    PostbackTemplateAction(
                                        label='補休',
                                        data='補休'
                                    ),
                                    PostbackTemplateAction(
                                        label='公假',
                                        data='公假'
                                    )
                                ]
                            )
                        ]
                    )
                )
                line_bot_api.reply_message(event.reply_token, carousel_template_message)
                return 0


            if content == "確認嗎?":
                name = d['answers'][1]
                start = d['answers'][4]
                end = d['answers'][5]
                print(start)
                print(end)
                print(type(start))
                print(type(end))
                start_time = (datetime.strptime(start, "%Y-%m-%dT%H:%M"))
                end_time = (datetime.strptime(end, "%Y-%m-%dT%H:%M"))
                print(start_time)
                time = end_time - start_time
                time = str(time)
                e = time.split(":")[0]
                if "day" in time:
                    if time.split(":")[1] == "00":
                        d = (time.split(",")[0]).split(" ")[0]
                        h = e.split(",")[1].replace(" ", "")
                        if h != "9":
                            j="%s天%s小時" %(d,h)
                        if time.split(":")[2] == "00" and h == "9":
                            d = int(d) + 1
                            j = "%s天" %d
                        if h == "0":
                            print("只顯示時間")
                            j = "%s天" %d
                    else:
                        d = (time.split(",")[0]).split(" ")[0]
                        h = e.split(",")[1].replace(" ", "")
                        z = int(h) + 1
                        j = "%s天%s小時" % (d, z)
                else:
                    if time.split(":")[1] == "00":
                        e = time.split(":")[0]
                        if int(e) >= 15:
                            e = int(e) - 15
                            j = "%s小時" %e
                        if int(e) == 9:
                            e = 1
                            j = "%s天" % e
                    else:
                        e = time.split(":")[0]
                        s = int(e) + 1
                        if int(e) >= 15:
                            s = int(s) - 15
                            j = "%s小時" %s
                ans = "請假單\n請假人姓名:%s\n請假起始時間:%s\n請假結束時間:%s\n總計請假:%s\n請問是否確認送出?" % (
                    name, start, end, j)
                confirm_template_message = TemplateSendMessage(
                    alt_text='Confirm template',
                    template=ConfirmTemplate(
                        text=ans,
                        actions=[
                            PostbackTemplateAction(
                                label='確定',
                                data='確定'
                            ),
                            MessageTemplateAction(
                                label='重新填寫',
                                text='請假申請單'
                            )
                        ]
                    )
                )
                line_bot_api.reply_message(event.reply_token, confirm_template_message)
                return 0

            if content == "從哪天開始請假?":
                print("yeso!")  #
                buttons_template_message = TemplateSendMessage(
                    alt_text='Buttons template',
                    template=ButtonsTemplate(
                        text='從哪天開始請假?',
                        actions=[
                            DatetimePickerTemplateAction(
                                label='請選擇',
                                data='datetime',
                                mode="datetime",
                                min=(datetime.now() - timedelta(days=30)).strftime("%Y-%m-%dT09:00"),
                                max=(datetime.now() + timedelta(days=30)).strftime("%Y-%m-%dT18:00")
                            )
                        ]
                    )
                )
                line_bot_api.reply_message(event.reply_token, buttons_template_message)
                return 0

            if content == "請到哪一天結束?":
                print("noa!")  #
                buttons_template_message = TemplateSendMessage(
                    alt_text='Buttons template',
                    template=ButtonsTemplate(
                        text='請到哪一天結束?',
                        actions=[
                            DatetimePickerTemplateAction(
                                label='請選擇',
                                data='datetime',
                                mode="datetime"
                            )
                        ]
                    )
                )
                line_bot_api.reply_message(event.reply_token, buttons_template_message)
                return 0

            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text=content))
            return 0





        if answers[0]=="加班申請單":   ##################################################################################
            print("forth")
            #######################################################
            if count == 1:
                client = general()
                sheet = client.open('test').get_worksheet(2)
                col = sheet.col_values(1)
                if event.message.text not in col:
                    event.message.text = "none"

            if count == 2 or count == 3:
                try:
                    datetime.strptime(event.message.text, "%Y-%m-%dT%H:%M")
                    hour = "%s" % (datetime.strptime(event.message.text, "%Y-%m-%dT%H:%M")).hour
                    minute = "%s" % (datetime.strptime(event.message.text, "%Y-%m-%dT%H:%M")).minute
                    print(hour)
                    if 0 <= (datetime.strptime(event.message.text, "%Y-%m-%dT%H:%M")).weekday() <= 4:
                        print("上班時間")
                        if 9 < int(hour) < 18:
                            event.message.text = "error"
                        if int(hour) == 9 and int(minute) != 0:
                            event.message.text = "error"
                except ValueError:
                    print("格式輸入錯誤")
                    event.message.text = "wrong"

            if count == 4:
                if event.message.text != "加班費" and event.message.text != "補修時數":
                    print("格式輸入錯誤")
                    event.message.text = "wrong"

            if count == 5:
                if event.message.text != "確定" and event.message.text != "加班申請單":
                    print("格式輸入錯誤")
                    event.message.text = "wrong"

            if event.message.text == "none":
                count = count - 1
                # content = over_time(count)
                line_bot_api.push_message(user_id, TextSendMessage(text="查無此人，請重新輸入"))

            if event.message.text == "error":
                count = count - 1
                # content = over_time(count)
                line_bot_api.push_message(user_id, TextSendMessage(text="輸入的資料不正確，加班時間應為非上班時間，請重新輸入"))

            if event.message.text == "wrong":
                count = count - 1
                # content = over_time(count)
                line_bot_api.push_message(user_id, TextSendMessage(text="格式輸入錯誤，請重新輸入"))

            answers.append(event.message.text)
            if "wrong" in answers:
                answers.remove("wrong")
            if "error" in answers:
                answers.remove("error")
            if "none" in answers:
                answers.remove("none")
            d['answers'] = answers
            count += 1

            if count >= 6:
                print("over")
                sub = str(
                    (datetime.strptime(answers[3], "%Y-%m-%dT%H:%M") - datetime.strptime(answers[2], "%Y-%m-%dT%H:%M")))
                e = sub.split(":")[0]
                answers.append(e)
                unique = (datetime.now() + timedelta(hours=8)).strftime('%Y_%m_%d_%H_%M_%S')
                name = d['answers'][1]
                id = name + "_" + unique
                print(id)

                ###############################################
                client = general()
                sheet = client.open('test').get_worksheet(2)
                col = sheet.col_values(2)
                if user_id in col:
                    value_index = col.index(user_id) + 1
                    boss_id = sheet.cell(value_index, 4).value
                    print(boss_id)

                ###############################################
                name = d['answers'][1]
                start = d['answers'][2]
                end = d['answers'][3]
                typ = d['answers'][4]
                month = "%s" % (datetime.strptime(start, "%Y-%m-%dT%H:%M")).month
                day = "%s" % (datetime.strptime(start, "%Y-%m-%dT%H:%M")).day
                ans = "%s月%s日的加班單\n加班人姓名:%s\n加班起始時間:%s\n加班結束時間:%s\n總計加班時數:%s\n轉換成:%s\n請問是否核准加班申請?" % (
                month, day, name, start, end, e, typ)
                confirm_template_message = TemplateSendMessage(
                    alt_text='加班申請審核通知',
                    template=ConfirmTemplate(
                        text=ans,
                        actions=[
                            PostbackTemplateAction(
                                label='是',
                                data=id + "_" + "yes"
                            ),
                            PostbackTemplateAction(
                                label='否',
                                data=id + "_" + "no"
                            )
                        ]
                    )
                )
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text='已將申請單遞交給相關人員，請靜待審核通知，謝謝'))
                line_bot_api.push_message(boss_id, confirm_template_message)


                answers.append(id)
                d['count'] = 0
                d['answers'] = []
                with open(path, 'w') as f:
                    json.dump(d, f, ensure_ascii=False)
                write_in(answers)
                return 0

            print("next")
            content = over_time(count)
            d['count'] = count
            with open(path, 'w') as f:
                json.dump(d, f, ensure_ascii=False)

            if content == "請選擇你的名字或手動輸入":
                client = general()
                sheet = client.open('test').get_worksheet(2)
                col = sheet.col_values(2)
                if user_id in col:
                    value_index = col.index(user_id) + 1
                    u_name = sheet.cell(value_index, 1).value
                    print(u_name)
                buttons_template = TemplateSendMessage(
                    alt_text='名字 template',
                    template=ButtonsTemplate(
                        text='請選擇你的名字或手動輸入',
                        actions=[
                            MessageTemplateAction(
                                label=u_name,
                                text=u_name
                            )
                        ]
                    )
                )
                line_bot_api.reply_message(event.reply_token, buttons_template)
                return 0

            if content == "轉換費用或補假":
                start = d['answers'][2]
                end = d['answers'][3]
                sub = str((datetime.strptime(end, "%Y-%m-%dT%H:%M") - datetime.strptime(start, "%Y-%m-%dT%H:%M")))
                e = sub.split(":")[0]
                if e == 0:
                    line_bot_api.push_message(user_id, TextSendMessage(text='加班時間未滿一小時不予以計算，請重新輸入'))  #
                    d['count'] = 0  #
                    d['answers'] = []  #
                    return 0  #
                ans = "總計加班時數共%s小時，請問想轉換成?" % e
                confirm_template_message = TemplateSendMessage(
                    alt_text='Confirm template',
                    template=ConfirmTemplate(
                        text=ans,
                        actions=[
                            PostbackTemplateAction(
                                label='加班費',
                                data='加班費'
                            ),
                            PostbackTemplateAction(
                                label='補修時數',
                                data='補修時數'
                            )
                        ]
                    )
                )
                line_bot_api.reply_message(event.reply_token, confirm_template_message)
                return 0

            if content == "確認嗎?":
                name = d['answers'][1]
                start = d['answers'][2]
                end = d['answers'][3]
                sub = str((datetime.strptime(end, "%Y-%m-%dT%H:%M") - datetime.strptime(start, "%Y-%m-%dT%H:%M")))
                e = sub.split(":")[0]
                typ = d['answers'][4]
                month = "%s" % (datetime.strptime(start, "%Y-%m-%dT%H:%M")).month
                day = "%s" % (datetime.strptime(start, "%Y-%m-%dT%H:%M")).day
                ans = "%s月%s日的加班單\n加班人姓名:%s\n加班起始時間:%s\n加班結束時間:%s\n總計加班時數:%s\n轉換成:%s\n請問是否確認送出?" % (
                month, day, name, start, end, e, typ)
                confirm_template_message = TemplateSendMessage(
                    alt_text='Confirm template',
                    template=ConfirmTemplate(
                        text=ans,
                        actions=[
                            PostbackTemplateAction(
                                label='確定',
                                data='確定'
                            ),
                            MessageTemplateAction(
                                label='重新填寫',
                                text='加班申請單'
                            )
                        ]
                    )
                )
                line_bot_api.reply_message(event.reply_token, confirm_template_message)
                return 0

            if content == "加班的起始時間是?":
                print("yeso!")  #
                buttons_template_message = TemplateSendMessage(
                    alt_text='Buttons template',
                    template=ButtonsTemplate(
                        text='加班的起始時間是?',
                        actions=[
                            DatetimePickerTemplateAction(
                                label='請選擇',
                                data='datetime',
                                mode="datetime",
                                min=(datetime.now() - timedelta(days=30)).strftime("%Y-%m-%dT18:00"),
                                max=(datetime.now() + timedelta(days=30)).strftime("%Y-%m-%dT09:00")
                            )
                        ]
                    )
                )
                line_bot_api.reply_message(event.reply_token, buttons_template_message)
                return 0

            if content == "加班的結束時間是?":
                print("noa!")  #
                buttons_template_message = TemplateSendMessage(
                    alt_text='Buttons template',
                    template=ButtonsTemplate(
                        text='加班的結束時間是?',
                        actions=[
                            DatetimePickerTemplateAction(
                                label='請選擇',
                                data='datetime',
                                mode="datetime"
                            )
                        ]
                    )
                )
                line_bot_api.reply_message(event.reply_token, buttons_template_message)
                return 0

            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text=content))
            return 0

    buttons_template = TemplateSendMessage(
        alt_text='目錄 template',
        template=ButtonsTemplate(
            title='申請單',
            text='請選擇',
            thumbnail_image_url='https://i.imgur.com/DzqsWsO.jpg',
            actions=[
                MessageTemplateAction(
                    label='我要加班',
                    text='加班申請單'
                ),
                MessageTemplateAction(
                    label='我要請假',
                    text='請假申請單'
                ),
                MessageTemplateAction(
                    label='查詢加班紀錄',
                    text='加班查詢'
                ),
                MessageTemplateAction(
                    label='查詢請假紀錄',
                    text='請假查詢'
                )
            ]
        )
    )
    line_bot_api.reply_message(event.reply_token, buttons_template)


@handler.add(PostbackEvent)
def Postback_message(event):
    print("event.reply_token:", event.reply_token)
    data=(str(event.postback.data))
    user_id = json.loads(str(event.source))['userId']

    ###################主管同意或拒絕加班申請單##########################
    client= general()
    sheet = client.open('test').get_worksheet(3)
    col=sheet.col_values(9)   #根據表單id做確認
    sheet2 = client.open('test').get_worksheet(1)
    col2=sheet2.col_values(8)
    if data.rsplit('_', 1)[0] in col and data.split("_")[-1] in["yes","no"]:   #
        value_index = col.index(data.rsplit('_', 1)[0])
        location = value_index + 1
        u_id=(data.rsplit('_', 1)[0]).split("_")[0]
        if sheet.cell(location, 12).value:  #如果已經有值的話，就停止繼續做，不然會一直重複寫入   ###
            return 0
        if data.split("_")[-1] == "yes":
            sheet.update_cell(location, 12, "是")  #答案是"是"的話，寫進去並發通知
            # @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@\
            anser = sheet.row_values(location)
            print(anser)
            name = anser[1]
            typ = anser[2]
            name_1 = anser[3]
            start = anser[4]
            end = anser[5]
            j = anser[7]
            if int(j) < 9:
                m = "%s小時" % j
            else:
                if int(j) % 9 == 0:
                    d2 = int(j) // 9
                    m = "%s天" % d2
                else:
                    d1 = int(j) // 9
                    h1 = int(j) % 9
                    m = "%s天%s小時" % (d1, h1)
            boss = anser[11]
            ans_end = "請假單\n請假人姓名:%s\n假別:%s\n代理人姓名:%s\n請假起始時間:%s\n請假結束時間:%s\n總計請假:%s\n審核是否通過:%s" % (
                name, typ, name_1, start, end, m, boss)
            #@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@\
            line_bot_api.push_message(u_id, TextSendMessage(text=ans_end))
            line_bot_api.push_message(user_id, TextSendMessage(text="你同意請假申請審核"))
        if data.split("_")[-1] == "no":
            sheet.update_cell(location, 12, "否")
            # @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!
            anser = sheet.row_values(location)
            print(anser)
            name = anser[1]
            typ = anser[2]
            name_1 = anser[3]
            start = anser[4]
            end = anser[5]
            j = anser[7]
            if int(j) < 9:
                m = "%s小時" % j
            else:
                if int(j) % 9 == 0:
                    d2 = j // 9
                    m = "%s天" % d2
                else:
                    d1 = j // 9
                    h1 = j % 9
                    m = "%s天%s小時" % (d1, h1)
            boss = anser[11]
            ans_end = "請假單\n請假人姓名:%s\n假別:%s\n代理人姓名:%s\n請假起始時間:%s\n請假結束時間:%s\n總計請假:%s\n審核是否通過:%s" % (
                name, typ, name_1, start, end, m, boss)
            # @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!
            line_bot_api.push_message(u_id, TextSendMessage(text=ans_end))
            line_bot_api.push_message(user_id, TextSendMessage(text="你已拒絕請假申請審核"))
        return 0
    if data.rsplit('_', 1)[0] in col2 and data.split("_")[-1] in["yes","no"]:
        value_index = col2.index(data.rsplit('_', 1)[0])
        location = value_index + 1
        u_id=(data.rsplit('_', 1)[0]).split("_")[0]
        if sheet2.cell(location, 11).value:  #如果已經有值的話，就停止繼續做，不然會一直重複寫入
            return 0
        if data.split("_")[-1] == "yes":
            sheet2.update_cell(location, 11, "是")  #答案是"是"的話，寫進去並發通知
            # @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@\
            anser = sheet2.row_values(location)
            print(anser)
            name = anser[1]
            start = anser[2]
            end = anser[3]
            typ = anser[4]
            j= anser[6]
            boss = anser[10]
            month = "%s" % (datetime.strptime(start, "%Y-%m-%dT%H:%M")).month
            day = "%s" % (datetime.strptime(start, "%Y-%m-%dT%H:%M")).day
            ans_end = "%s月%s日的加班單\n加班人姓名:%s\n加班起始時間:%s\n加班結束時間:%s\n總計加班時數:%s\n轉換成:%s\n審核是否通過:%s" % (month, day, name, start, end, j, typ, boss)
            #@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@\
            line_bot_api.push_message(u_id, TextSendMessage(text=ans_end))
            line_bot_api.push_message(user_id, TextSendMessage(text="你同意加班申請審核"))
        if data.split("_")[-1] == "no":
            sheet2.update_cell(location, 11, "否")
            # @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!
            anser = sheet2.row_values(location)
            print(anser)
            name = anser[1]
            start = anser[2]
            end = anser[3]
            typ = anser[4]
            j = anser[6]
            boss = anser[10]
            month = "%s" % (datetime.strptime(start, "%Y-%m-%dT%H:%M")).month
            day = "%s" % (datetime.strptime(start, "%Y-%m-%dT%H:%M")).day
            ans_end = "%s月%s日的加班單\n加班人姓名:%s\n加班起始時間:%s\n加班結束時間:%s\n總計加班時數:%s\n轉換成:%s\n審核是否通過:%s" % (
            month, day, name, start, end, j, typ, boss)
            # @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!
            line_bot_api.push_message(u_id, TextSendMessage(text=ans_end))
            line_bot_api.push_message(user_id, TextSendMessage(text="你已拒絕加班申請審核"))
        return 0

    ######如果data是時間的話，要做另外處理(因為只有時間格式才是datetime，其他不是，會抓錯)#######
    if data== "datetime":
        time = str(event.postback.params['datetime'])
        data=time


    users = os.listdir("Users")
    answers = []
    name = user_id + ".json"
    path = "Users/" + name
    print(path)
    if name in users:
        print("first_1")
        with open(path, encoding="utf-8") as json_data:
            d = json.load(json_data)
        count = d['count']
        answers = d['answers']
        print(d)
        print(type(d))
        print(answers)

    if name not in users:
        print("second_1")
        d = {"user_id": user_id, "count": 0, "answers": []}  # 記得要建answer的部分
        with open(path, 'w', encoding='UTF-8') as f:
            json.dump(d, f, ensure_ascii=False)
        count = d['count']
        answers = d['answers']


    if data == "加班申請單" or data == "請假申請單" or data=="加班查詢" or data=="請假查詢":
        if count > 0:    #只要是按加班申請單，就是從頭開始做，得把之前的紀錄刪掉
            d['count'] = 0
            d['answers'] = []
            answers=d['answers']
        print("third_1")
        count = 1
        answers.append(data)
        content = over_time(count)
        d['count'] = 1
        d['answers'] = answers
        with open(path, 'w') as f:
            json.dump(d, f, ensure_ascii=False)
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=content))
        return 0


    if count != 0:

        if answers[0] == '請假查詢':
            print("forth_off_1")
            if count == 1:
                client = general()
                sheet = client.open('test').get_worksheet(2)
                col = sheet.col_values(1)
                if data not in col:
                    data = "none"

            if count==2:
                if data!="某月出勤狀況" and data!="今年已請假別與天數":
                    data="wrong"


            if count==3:
                print(data)
                if int(data) not in range(2017,2019):
                    data="wyear"

            if count==4:
                print(data)
                if int(data) not in range(1,13):
                    data="wmonth"

            if data == "none":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="查無此人，請重新輸入"))

            if data == "wrong":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="格式輸入錯誤，請重新輸入"))

            if data == "wmonth":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="輸入錯誤，請輸入1~12的數字"))

            if data == "wyear":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="輸入錯誤，請輸入四位數字，預設初始年份為2017，目前可以查至2018年"))

            if data == "今年已請假別與天數":
                now = datetime.now().strftime('%Y-%m-%dT%H:%M')
                name = answers[1]
                year = "%s" % (datetime.strptime(now, "%Y-%m-%dT%H:%M")).year
                c = love(name, year, "事假")
                c1= love(name, year, "公傷病假")
                c2 = love(name, year, "普通傷病假")
                c3 = love(name, year, "生理假")
                c4 = love(name, year, "喪假")
                c5 = love(name, year, "婚假")
                c6 = love(name, year, "特休")
                c7 = love(name, year, "補休")
                c8 = love(name, year, "公假")
                ans="%s總共請了\n%s\n%s\n%s\n%s\n%s\n%s\n%s\n%s\n%s" % (
                                year, c, c1, c2, c3, c4, c5, c6, c7, c8)
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text=ans))
                return 0

            answers.append(data)
            print(answers)
            if "none" in answers:
                answers.remove("none")
            if "wrong" in answers:
                answers.remove("wrong")
            if "wmonth" in answers:
                answers.remove("wmonth")
            if "wyear" in answers:
                answers.remove("wyear")
            print(answers)
            d['answers'] = answers
            print(d)
            count += 1
            print(count)

            print("next")
            content = search_off(count)
            print(content)
            d['count'] = count
            with open(path, 'w') as f:
                json.dump(d, f, ensure_ascii=False)

            if content == "請問要查詢哪個年份的請假狀況?":
                line_bot_api.reply_message(
                    event.reply_token,
                    TextSendMessage(text=content))
                return 0

            if content == "請問要查詢哪個月份的請假狀況?":
                line_bot_api.reply_message(
                    event.reply_token,
                    TextSendMessage(text=content))
                return 0


            if content=="查詢選項":
                buttons_template = TemplateSendMessage(
                    alt_text='查詢 template',
                    template=ButtonsTemplate(
                        text='請選擇',
                        actions=[
                            MessageTemplateAction(
                                label='某月出勤狀況',
                                text='某月出勤狀況'
                            ),
                            MessageTemplateAction(
                                label='今年已請假別與天數',
                                text='今年已請假別與天數'
                            )
                        ]
                    )
                )
                line_bot_api.reply_message(event.reply_token, buttons_template)
                return 0



        if answers[0] == '加班查詢':
            print("forth_off_1")
            if count == 1:
                client = general()
                sheet = client.open('test').get_worksheet(2)
                col = sheet.col_values(1)
                if data not in col:
                    data = "none"

            if count==2:
                print(data)
                if int(data) not in range(2017,2019):
                    data="wyear"

            if count==3:
                print(data)
                if int(data) not in range(1,13):
                    data="wmonth"

            if count ==4:
                if data!="累積加班時數" and data!="完整加班紀錄":
                    data="wrong"


            if data == "none":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="查無此人，請重新輸入"))

            if data == "wmonth":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="輸入錯誤，請輸入1~12的數字"))

            if data == "wyear":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="輸入錯誤，請輸入四位數字，預設初始年份為2017，目前可以查至2018年"))

            if data == "wrong":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="格式輸入錯誤，請重新輸入"))

            if data =="累積加班時數":
                client=general()
                sheet = client.open('test').get_worksheet(1)
                col = sheet.col_values(2)
                year=answers[2]
                month=answers[3]
                print(month)
                name=answers[1]
                total = 0
                total_d = 0
                total_m = 0
                for j in range(len(col)):
                    j = j + 1
                    if (sheet.cell(j, 2).value) == name and (sheet.cell(j, 9).value) == "是":  # 找出名字一樣且審核通過的
                        time = (sheet.cell(j, 3).value)
                        if (datetime.strptime(time, "%Y-%m-%dT%H:%M")).year == int(year) and (datetime.strptime(time, "%Y-%m-%dT%H:%M")).month == int(month):  # 12月份的
                            total = total + int(sheet.cell(j, 7).value)
                            if total == 46:
                                ans="%s年%s月份\n姓名:%s\n累積加班時數:%s小時\n已達勞基法上限\n\n=====轉換成=====\n補修時數:%s小時\n加班費時數:%s小時" % (
                                year, month, name, total, total_d, total_m)
                                line_bot_api.reply_message(event.reply_token, TextSendMessage(text=ans))
                                return 0
                            print(type(total))
                            if (sheet.cell(j, 5).value) == "補修時數":
                                total_d = total_d + int(sheet.cell(j, 7).value)
                            if (sheet.cell(j, 5).value) == "加班費":
                                total_m = total_m + int(sheet.cell(j, 7).value)
                            ans = "%s年%s月份\n姓名:%s\n累積加班時數:%s小時\n\n=====轉換成=====\n補修時數:%s小時\n加班費時數:%s小時" % (
                                year, month, name, total, total_d, total_m)
                        else:
                            ans = "此月份無任何加班紀錄"
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text=ans))
                return 0

            if data =="完整加班紀錄":
                client=general()
                sheet = client.open('test').get_worksheet(1)
                col = sheet.col_values(2)
                year = answers[2]
                month=answers[3]
                name=answers[1]
                b=[]
                for j in range(len(col)):
                    j = j + 1
                    if (sheet.cell(j, 2).value) == name and (sheet.cell(j, 9).value) == "是":  # 找出名字一樣且審核通過的
                        time = (sheet.cell(j, 3).value)
                        date = time.split("T")[0]
                        #end = (sheet.cell(j, 4).value)
                        hour = (sheet.cell(j, 7).value)
                        if (datetime.strptime(time, "%Y-%m-%dT%H:%M")).year == int(year) and (datetime.strptime(time, "%Y-%m-%dT%H:%M")).month == int(month):  # 12月份的
                            b.append(("加班日期: %s，加班時數: %s" % (date, hour)))
                            print(b)
                            ans=("\n".join(b))
                        else:
                            ans="此月份無任何加班紀錄"
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text=ans))
                return 0


            answers.append(data)
            print(answers)
            if "none" in answers:
                answers.remove("none")
            if "wmonth" in answers:
                answers.remove("wmonth")
            if "wrong" in answers:
                answers.remove("wrong")
            if "wyear" in answers:
                answers.remove("wyear")
            print(answers)
            d['answers'] = answers
            print(d)
            count += 1
            print(count)

            print("next")
            content = search_over(count)
            print(content)
            d['count'] = count
            with open(path, 'w') as f:
                json.dump(d, f, ensure_ascii=False)

            if content=="請問要查詢哪個年份的加班資訊?":
                line_bot_api.reply_message(
                    event.reply_token,
                    TextSendMessage(text=content))
                return 0

            if content=="請問要查詢哪個月份的加班資訊?":
                line_bot_api.reply_message(
                    event.reply_token,
                    TextSendMessage(text=content))
                return 0

            if content=="加班細節":
                buttons_template = TemplateSendMessage(
                    alt_text='查詢 template',
                    template=ButtonsTemplate(
                        text='請選擇',
                        actions=[
                            MessageTemplateAction(
                                label='累積加班時數',
                                text='累積加班時數'
                            ),
                            MessageTemplateAction(
                                label='完整加班紀錄',
                                text='完整加班紀錄'
                            )
                        ]
                    )
                )
                line_bot_api.reply_message(event.reply_token, buttons_template)
                return 0





        if answers[0] == '請假申請單':
            print("8th")
            #######################################################
            if count == 1:
                client = general()
                sheet = client.open('test').get_worksheet(2)
                col = sheet.col_values(1)
                if data not in col:
                    data = "none"

            if count == 2:
                if data not in ["事假", "公傷病假", "普通傷病假", "生理假", "喪假", "婚假", "特休", "補休", "公假"]:
                    data = "wrong"

            if count == 3:
                client = general()
                sheet = client.open('test').get_worksheet(2)
                cols = sheet.col_values(1)
                if data not in cols:
                    data = "none"
                if data in cols:
                    client = general()
                    sheet = client.open('test').get_worksheet(2)
                    col = sheet.col_values(2)
                    col2 = sheet.col_values(5)
                    print(len(col2))
                    value_index = col.index(user_id) + 1
                    print(value_index)
                    depart = sheet.cell(value_index, 5).value
                    anse = []
                    for i in range(len(col2)):
                        i = i + 1
                        print(i)
                        if i == value_index:
                            continue
                        if sheet.cell(i, 5).value == depart:
                            print(i)
                            print(sheet.cell(i, 1).value)
                            anse.append(sheet.cell(i, 1).value)
                    if data not in anse:
                        data = "else"

            if count == 4 or count == 5:
                try:
                    datetime.strptime(data, "%Y-%m-%dT%H:%M")
                    hour = "%s" % (datetime.strptime(data, "%Y-%m-%dT%H:%M")).hour
                    minute = "%s" % (datetime.strptime(data, "%Y-%m-%dT%H:%M")).minute
                    print(hour)
                    if 0 <= (datetime.strptime(data, "%Y-%m-%dT%H:%M")).weekday() <= 5:
                        print("周間")
                        if 0 <= int(hour) < 9 or 18 < int(hour) < 24:
                            data = "error"
                        if int(hour) == 18 and int(minute) != 0:
                            data = "error"
                    else:
                        print("周末")
                        data = "off"
                except ValueError:
                    print("格式輸入錯誤")
                    data = "wrong"

            if count == 6:
                if data != "確定" and data != "請假申請單":
                    print("格式輸入錯誤")
                    data = "wrong"

            if data == "none":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="查無此人，請重新輸入"))

            if data == "off":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="此時間為周末時間，應該沒有請假的必要，有疑問請詢問人事處，並在下方重新輸入"))

            if data == "error":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="輸入的資料不正確，請假時間應為平日上班時間，請重新輸入"))

            if data == "wrong":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="格式輸入錯誤，請重新輸入"))

            if data == "else":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="請輸入同部門的同仁"))

            answers.append(data)
            print(answers)
            if "wrong" in answers:
                answers.remove("wrong")
            if "off" in answers:
                answers.remove("off")
            if "error" in answers:
                answers.remove("error")
            if "none" in answers:
                answers.remove("none")
            if "else" in answers:
                answers.remove("else")
            print(answers)
            d['answers'] = answers
            d['count'] = count
            count += 1

            if count >= 7:
                print("over")
                line_bot_api.push_message(user_id, TextSendMessage(text='已將申請單遞交給相關人員，請靜待審核通知，謝謝'))
                time = (datetime.strptime(answers[5], "%Y-%m-%dT%H:%M")) - (datetime.strptime(answers[4], "%Y-%m-%dT%H:%M"))
                time = str(time)
                e = time.split(":")[0]
                if "day" in time:
                    if time.split(":")[1] == "00":
                        day = (time.split(",")[0]).split(" ")[0]
                        h = e.split(",")[1].replace(" ", "")
                        if h != "9":
                            if int(h) > 15:
                                h = int(h) - 15
                            j = int(day) * 9 + int(h)
                        if h == "9":
                            day = int(day) + 1
                            j = day * 9
                        if h == "0":
                            j = int(day) * 9
                    else:
                        day = (time.split(",")[0]).split(" ")[0]
                        h = e.split(",")[1].replace(" ", "")
                        z = int(h) + 1
                        if int(z) > 15:
                            z = int(z) - 15
                            j = int(day) * 9 + int(z)
                        elif z == "9":
                            day = int(day) + 1
                            j = int(day) * 9
                        else:
                            j = int(day) * 9 + int(z)
                else:
                    if time.split(":")[1] == "00":
                        e = time.split(":")[0]
                        j = e
                        if int(e) >= 15:
                            e = int(e) - 15
                            j = e
                        if int(e) == 9:
                            j = e
                    else:
                        e = time.split(":")[0]
                        s = int(e) + 1
                        j = s
                        if int(s) >= 15:
                            s = int(s) - 15
                            j = s
                        if int(s) == 9:
                            j = s
                answers.append(j)
                unique = (datetime.now() + timedelta(hours=8)).strftime('%Y_%m_%d_%H_%M_%S')
                id = user_id + "_" + unique
                answers.append(id)
                yeara= (datetime.strptime(answers[4], "%Y-%m-%dT%H:%M")).year   ####
                montha = (datetime.strptime(answers[4], "%Y-%m-%dT%H:%M")).month    ####
                answers.append(yeara)
                answers.append(montha)
                ############找出主管id#################
                client = general()
                sheet = client.open('test').get_worksheet(2)
                col = sheet.col_values(1)
                person= answers[1]
                if person in col:
                    value_index = col.index(person) + 1
                    boss_id = sheet.cell(value_index, 4).value

                # 將表單的暫存內容歸零
                d['count'] = 0
                d['answers'] = []
                with open(path, 'w') as f:
                    json.dump(d, f, ensure_ascii=False)
                write_to(answers)

                name = answers[1]
                typ = answers[2]
                name_1 = answers[3]
                start = answers[4]
                end = answers[5]
                j=answers[7]
                if int(j) < 9:
                    m = "%s小時" % j
                else:
                    if int(j) % 9 == 0:
                        d2 = int(j) // 9
                        m = "%s天" % d2
                    else:
                        d1 = int(j) // 9
                        h1 = int(j) % 9
                        m = "%s天%s小時" % (d1, h1)

                ans = "請假單\n請假人姓名:%s\n請假類型:%s\n指定代理人是:%s\n請假起始時間:%s\n請假結束時間:%s\n總計請假:%s\n請問是否核准請假申請?" % (
                    name, typ, name_1, start, end, m)
                confirm_template_message = TemplateSendMessage(
                    alt_text='請假申請審核通知',
                    template=ConfirmTemplate(
                        text=ans,
                        actions=[
                            PostbackTemplateAction(
                                label='是',
                                data=id + "_" + "yes"
                            ),
                            PostbackTemplateAction(
                                label='否',
                                data=id + "_" + "no"
                            )
                        ]
                    )
                )
                line_bot_api.push_message(boss_id, confirm_template_message)
                return 0

            print("next")
            print(count)
            content = off_time(count)
            d['count'] = count
            with open(path, 'w') as f:
                json.dump(d, f, ensure_ascii=False)

            if content == "請選擇你的名字或手動輸入":
                client = general()
                sheet = client.open('test').get_worksheet(2)
                col = sheet.col_values(2)
                if user_id in col:
                    value_index = col.index(user_id) + 1
                    u_name = sheet.cell(value_index, 1).value
                    print(u_name)
                buttons_template = TemplateSendMessage(
                    alt_text='名字 template',
                    template=ButtonsTemplate(
                        text='請選擇你的名字或手動輸入',
                        actions=[
                            MessageTemplateAction(
                                label=u_name,
                                text=u_name
                            )
                        ]
                    )
                )
                line_bot_api.reply_message(event.reply_token, buttons_template)
                return 0

            if content == "指定代理人是?":
                client = general()
                sheet = client.open('test').get_worksheet(2)
                col = sheet.col_values(2)
                col2 = sheet.col_values(5)
                print(len(col2))
                if user_id in col:
                    value_index = col.index(user_id) + 1
                    print(value_index)
                    depart = sheet.cell(value_index, 5).value
                    anse = []
                    for i in range(len(col2)):
                        i = i + 1
                        print(i)
                        if i == value_index:
                            continue
                        if sheet.cell(i, 5).value == depart:
                            print(i)
                            print(sheet.cell(i, 1).value)
                            anse.append(sheet.cell(i, 1).value)
                    print(anse)
                if len(anse) == 2:
                    buttons_template = TemplateSendMessage(
                        alt_text='名字 template',
                        template=ButtonsTemplate(
                            text='指定代理人是?',
                            actions=[
                                MessageTemplateAction(
                                    label=anse[0],
                                    text=anse[0]
                                ),
                                MessageTemplateAction(
                                    label=anse[1],
                                    text=anse[1]
                                )
                            ]
                        )
                    )
                    line_bot_api.reply_message(event.reply_token, buttons_template)
                    return 0

                if len(anse) == 3:
                    buttons_template = TemplateSendMessage(
                        alt_text='名字 template',
                        template=ButtonsTemplate(
                            text='指定代理人是?',
                            actions=[
                                MessageTemplateAction(
                                    label=anse[0],
                                    text=anse[0]
                                ),
                                MessageTemplateAction(
                                    label=anse[1],
                                    text=anse[1]
                                ),
                                MessageTemplateAction(
                                    label=anse[2],
                                    text=anse[2]
                                )
                            ]
                        )
                    )
                    line_bot_api.reply_message(event.reply_token, buttons_template)
                    return 0

                if len(anse) == 4:
                    buttons_template = TemplateSendMessage(
                        alt_text='名字 template',
                        template=ButtonsTemplate(
                            text='指定代理人是?',
                            actions=[
                                MessageTemplateAction(
                                    label=anse[0],
                                    text=anse[0]
                                ),
                                MessageTemplateAction(
                                    label=anse[1],
                                    text=anse[1]
                                ),
                                MessageTemplateAction(
                                    label=anse[2],
                                    text=anse[2]
                                ),
                                MessageTemplateAction(
                                    label=anse[3],
                                    text=anse[3]
                                )
                            ]
                        )
                    )
                    line_bot_api.reply_message(event.reply_token, buttons_template)
                    return 0

            if content == "請問要請什麼假?":
                carousel_template_message = TemplateSendMessage(
                    alt_text='Carousel template',
                    template=CarouselTemplate(
                        columns=[
                            CarouselColumn(
                                text='請問要請什麼假?',
                                actions=[
                                    PostbackTemplateAction(
                                        label='事假',
                                        data='事假'
                                    ),
                                    PostbackTemplateAction(
                                        label='公傷病假',
                                        data='公傷病假'
                                    ),
                                    PostbackTemplateAction(
                                        label='普通傷病假',
                                        data='普通傷病假'
                                    )
                                ]
                            ),
                            CarouselColumn(
                                text=' ',
                                actions=[
                                    PostbackTemplateAction(
                                        label='生理假',
                                        data='生理假'
                                    ),
                                    PostbackTemplateAction(
                                        label='喪假',
                                        data='喪假'
                                    ),
                                    PostbackTemplateAction(
                                        label='婚假',
                                        data='婚假'
                                    )
                                ]
                            ),
                            CarouselColumn(
                                text=' ',
                                actions=[
                                    PostbackTemplateAction(
                                        label='特休',
                                        data='特休'
                                    ),
                                    PostbackTemplateAction(
                                        label='補休',
                                        data='補休'
                                    ),
                                    PostbackTemplateAction(
                                        label='公假',
                                        data='公假'
                                    )
                                ]
                            )
                        ]
                    )
                )
                line_bot_api.reply_message(event.reply_token, carousel_template_message)
                return 0

            if content == "確認嗎?":
                name = d['answers'][1]
                start = d['answers'][4]
                end = d['answers'][5]
                start_time = (datetime.strptime(start, "%Y-%m-%dT%H:%M"))
                end_time = (datetime.strptime(end, "%Y-%m-%dT%H:%M"))
                time = end_time - start_time
                time = str(time)
                e = time.split(":")[0]
                if "day" in time:
                    if time.split(":")[1] == "00":
                        day = (time.split(",")[0]).split(" ")[0]
                        print(day)
                        h = e.split(",")[1].replace(" ", "")
                        print(h)
                        if h != "9":
                            if int(h) > 15:
                                h = int(h) - 15
                            jn = int(day) * 9 + int(h)
                        if h == "9":
                            day = int(day) + 1
                            jn = day * 9
                        if h == "0":
                            print("只顯示時間")
                            jn = int(day) * 9
                    else:
                        day = (time.split(",")[0]).split(" ")[0]
                        print(day)
                        h = e.split(",")[1].replace(" ", "")
                        print(h)
                        z = int(h) + 1
                        if int(z) > 15:
                            z = int(z) - 15
                            jn = int(day) * 9 + int(z)
                        elif z == "9":
                            day = int(day) + 1
                            jn = int(day) * 9
                        else:
                            jn = int(day) * 9 + int(z)
                else:
                    if int(e) < 1:
                        ans = "請假時間未滿一小時不予以計算，請重新填單"
                        d['count'] = 0
                        d['answers'] = []
                        with open(path, 'w') as f:
                            json.dump(d, f, ensure_ascii=False)
                            line_bot_api.push_message(user_id, TextSendMessage(text=ans))
                        return 0

                    elif time.split(":")[1] == "00":
                        e = time.split(":")[0]
                        jn = e
                        if int(e) >= 15:
                            e = int(e) - 15
                            jn = e
                        if int(e) == 9:
                            jn = e
                    else:
                        e = time.split(":")[0]
                        s = int(e) + 1
                        jn = s
                        if int(s) >= 15:
                            s = int(s) - 15
                            jn = s
                        if int(s) == 9:
                            jn = s

                ###################1225
                if int(jn) < 9:
                    m = "%s小時" % jn
                else:
                    if int(jn) % 9 == 0:
                        d2 = int(jn) // 9
                        m = "%s天" % d2
                    else:
                        d1 = int(jn) // 9
                        h1 = int(jn) % 9
                        m = "%s天%s小時" % (d1, h1)
                print("m是%s" %m)
                year=(datetime.strptime(start, "%Y-%m-%dT%H:%M")).year
                if answers[2]=="事假":
                    print("市價yes")
                    client = general()
                    sheet = client.open('test').get_worksheet(5)
                    col = sheet.col_values(1)
                    print(col)
                    for j in range(len(col)):
                        j = j + 1
                        if (sheet.cell(j, 1).value) == name and int((sheet.cell(j, 2).value)) == year:
                            holi = int(sheet.cell(j, 3).value)
                    rest = 126 - holi
                    if rest == 0:
                        ans = "根據勞基法，目前已達今年的事假申請上限(14日)，無法再請事假"
                        d['count'] = 0
                        d['answers'] = []
                        with open(path, 'w') as f:
                            json.dump(d, f, ensure_ascii=False)
                            line_bot_api.push_message(user_id, TextSendMessage(text=ans))
                        return 0
                    if rest < int(jn):
                        ans = "根據勞基法，你的事假申請將超過今年度的上限(14日)，請查詢請假記錄再做申請"
                        d['count'] = 0
                        d['answers'] = []
                        with open(path, 'w') as f:
                            json.dump(d, f, ensure_ascii=False)
                            line_bot_api.push_message(user_id, TextSendMessage(text=ans))
                        return 0
                    print("市價yes2")
                if answers[2] == "普通傷病假":
                    client = general()
                    sheet = client.open('test').get_worksheet(5)
                    col = sheet.col_values(1)
                    for j in range(len(col)):
                        j = j + 1
                        if (sheet.cell(j, 1).value) == name and int((sheet.cell(j, 2).value)) == year:
                            holi = int(sheet.cell(j, 5).value)
                    rest = 270 - holi
                    if rest == 0:
                        ans = "根據勞基法，目前已達今年的普通傷病假申請上限(30日)，無法再請普通傷病假"
                        d['count'] = 0
                        d['answers'] = []
                        with open(path, 'w') as f:
                            json.dump(d, f, ensure_ascii=False)
                            line_bot_api.push_message(user_id, TextSendMessage(text=ans))
                        return 0
                    if rest < int(jn):
                        ans = "根據勞基法，你的普通傷病假申請將超過今年度的上限(30日)，請查詢請假記錄再做申請"
                        d['count'] = 0
                        d['answers'] = []
                        with open(path, 'w') as f:
                            json.dump(d, f, ensure_ascii=False)
                            line_bot_api.push_message(user_id, TextSendMessage(text=ans))
                        return 0
                if answers[2] == "婚假":
                    client = general()
                    sheet = client.open('test').get_worksheet(5)
                    col = sheet.col_values(1)
                    for j in range(len(col)):
                        j = j + 1
                        if (sheet.cell(j, 1).value) == name and int((sheet.cell(j, 2).value)) == year:
                            holi = int(sheet.cell(j, 8).value)
                    rest = 72 - holi
                    if rest == 0:
                        ans = "根據勞基法，目前已達今年的婚假申請上限(8日)，無法再請婚假"
                        d['count'] = 0
                        d['answers'] = []
                        with open(path, 'w') as f:
                            json.dump(d, f, ensure_ascii=False)
                            line_bot_api.push_message(user_id, TextSendMessage(text=ans))
                        return 0
                    if rest < int(jn):
                        ans = "根據勞基法，你的婚假申請將超過今年度的上限(8日)，請查詢請假記錄再做申請"
                        d['count'] = 0
                        d['answers'] = []
                        with open(path, 'w') as f:
                            json.dump(d, f, ensure_ascii=False)
                            line_bot_api.push_message(user_id, TextSendMessage(text=ans))
                        return 0
                if answers[2] == "補休":
                    print("bubu")
                    client = general()
                    sheet = client.open('test').get_worksheet(6)
                    col = sheet.col_values(1)
                    sumt = 0
                    for kr in range(len(col)):
                        kr = kr + 1
                        if (sheet.cell(kr, 1).value) == name and int((sheet.cell(kr, 2).value)) == year:
                            sumt += int(sheet.cell(kr, 6).value)   #加班換多少補休
                    client = general()
                    sheet2 = client.open('test').get_worksheet(5)
                    col2 = sheet2.col_values(1)
                    for n in range(len(col2)):
                        n = n + 1
                        if (sheet2.cell(n, 1).value) == name and int((sheet2.cell(n, 2).value)) == year:
                            print(sheet2.cell(n, 10).value)
                            print(type(sheet2.cell(n, 10).value))
                            much = int(sheet2.cell(n, 10).value)  # 請了多少補休
                    rest = sumt - much
                    if rest == 0:
                        ans = "你的補修時數已用完，無法再做補修"
                        d['count'] = 0
                        d['answers'] = []
                        with open(path, 'w') as f:
                            json.dump(d, f, ensure_ascii=False)
                            line_bot_api.push_message(user_id, TextSendMessage(text=ans))
                        return 0
                    if rest < int(jn):
                        ans = "無法申請，你目前只能申請%s小時的補修，請重新填寫" % rest
                        d['count'] = 0
                        d['answers'] = []
                        with open(path, 'w') as f:
                            json.dump(d, f, ensure_ascii=False)
                            line_bot_api.push_message(user_id, TextSendMessage(text=ans))
                        return 0

                ######1225
                ans = "請假單\n請假人姓名:%s\n請假起始時間:%s\n請假結束時間:%s\n總計請假:%s\n請問是否確認送出?" % (
                    name, start, end, m)
                print(ans)
                confirm_template_message = TemplateSendMessage(
                    alt_text='Confirm template',
                    template=ConfirmTemplate(
                        text=ans,
                        actions=[
                            PostbackTemplateAction(
                                label='確定',
                                data='確定'
                            ),
                            MessageTemplateAction(
                                label='重新填寫',
                                text='請假申請單'
                            )
                        ]
                    )
                )
                line_bot_api.reply_message(event.reply_token, confirm_template_message)
                return 0

            if content == "從哪天開始請假?":
                buttons_template_message = TemplateSendMessage(
                    alt_text='Buttons template',
                    template=ButtonsTemplate(
                        text='從哪天開始請假?',
                        actions=[
                            DatetimePickerTemplateAction(
                                label='請選擇',
                                data='datetime',
                                mode="datetime",
                                min=(datetime.now() - timedelta(days=30)).strftime("%Y-%m-%dT09:00"),
                                max=(datetime.now() + timedelta(days=30)).strftime("%Y-%m-%dT18:00")
                            )
                        ]
                    )
                )
                line_bot_api.reply_message(event.reply_token, buttons_template_message)
                return 0

            if content == "請到哪一天結束?":
                if data=="error" or data=="off" or data=="wrong":
                    data=answers[4]
                buttons_template_message = TemplateSendMessage(
                    alt_text='Buttons template',
                    template=ButtonsTemplate(
                        text='請到哪一天結束?',
                        actions=[
                            DatetimePickerTemplateAction(
                                label='請選擇',
                                data='datetime',
                                mode="datetime",
                                min=data,
                                max=(datetime.strptime(data, "%Y-%m-%dT%H:%M") + timedelta(days=360)).strftime(
                                    "%Y-%m-%dT%H:%M")
                            )
                        ]
                    )
                )
                line_bot_api.reply_message(event.reply_token, buttons_template_message)
                return 0

            #line_bot_api.reply_message(
               # event.reply_token,
                #TextSendMessage(text=content))
            #return 0
        #################################################################################################加班申請單
        if answers[0]=='加班申請單':

            if count == 1:
                client = general()
                sheet = client.open('test').get_worksheet(2)
                col = sheet.col_values(1)
                if data not in col:
                    data = "none"
            #如果人事資料表裡沒有這個人名會出錯

            if count == 2 or count == 3:
                try:  # 如果符合時間格式，那一切沒問題
                    datetime.strptime(data, "%Y-%m-%dT%H:%M")
                    data = data
                    hour = "%s" % (datetime.strptime(data, "%Y-%m-%dT%H:%M")).hour
                    minute = "%s" % (datetime.strptime(data, "%Y-%m-%dT%H:%M")).minute
                    print(hour)
                    if 0 <= (datetime.strptime(data, "%Y-%m-%dT%H:%M")).weekday() <= 4:
                        #上班時間無法加班。下面的設定是，抓到是早上九點多的會出錯，但九點是允許範圍
                        if 9 < int(hour) < 18:
                            data = "error"
                        if int(hour) == 9 and int(minute) != 0:
                            data = "error"
                except ValueError:  # 如果不符合時間格式，直接讓data等於一個錯值，等一下再檢查刪除這個錯值
                    data = "wrong"

            if count == 4:
                if data != "加班費" and data != "補修時數":
                    data = "wrong"

            if count == 5:
                if data != "確定" and data != "加班申請單":
                    data = "wrong"

            if data == "wrong":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="格式輸入錯誤，請重新輸入"))

            if data == "error":
                count = count - 1
                line_bot_api.push_message(user_id, TextSendMessage(text="輸入的資料不正確，加班時間應為非上班時間，請重新輸入"))

            if data == "none":
                count = count - 1
                # content = over_time(count)
                line_bot_api.push_message(user_id, TextSendMessage(text="查無此人，請重新輸入"))

            answers.append(data)
            if "wrong" in answers:
                answers.remove("wrong")  # 不管data是不是對的，先寫入，如果是錯值，直接刪除
            if "error" in answers:
                answers.remove("error")
            if "none" in answers:
                answers.remove("none")
            d['answers'] = answers
            count += 1

            if count >= 6:
                print("over")  # 要結束了
                line_bot_api.push_message(user_id, TextSendMessage(text='已將申請單遞交給相關人員，請靜待審核通知，謝謝'))  #XXXXXXXXXXXXXXX改成push
                sub = str((datetime.strptime(answers[3], "%Y-%m-%dT%H:%M") - datetime.strptime(answers[2], "%Y-%m-%dT%H:%M")))
                e = sub.split(":")[0]
                answers.append(e)  # e是總時數
                unique = (datetime.now() + timedelta(hours=8)).strftime('%Y_%m_%d_%H_%M_%S')
                id = user_id + "_" + unique  # 表單id
                answers.append(id)
                y2 = (datetime.strptime(answers[2], "%Y-%m-%dT%H:%M")).year
                m2 = (datetime.strptime(answers[2], "%Y-%m-%dT%H:%M")).month
                answers.append(y2)  #年份
                answers.append(m2)   #月份
                ########找出主管id#########
                client = general()
                sheet = client.open('test').get_worksheet(2)
                col = sheet.col_values(1)
                person = answers[1]  #因為你有可能幫別人寫，所以用名字去找對方的主管id比較妥當
                if person in col:
                    value_index = col.index(person) + 1
                    boss_id = sheet.cell(value_index, 4).value


                # 將表單的暫存內容歸零，再用answers做儲存(儲存可能需要比較長的時間，但是使用者可能會馬上填第二張單子)#
                d['count'] = 0
                d['answers'] = []
                with open(path, 'w') as f:
                    json.dump(d, f, ensure_ascii=False)
                write_in(answers)  # 在把使用者的回答寫入sheet的期間，因為暫存內容已歸零，使用者已可以填下一張單

                ######d已經歸零，但answers裡的值都在，所以抓取answers裡的值來做回覆########
                name = answers[1]
                start = answers[2]
                end = answers[3]
                e = answers[6]
                typ = answers[4]
                month = "%s" % (datetime.strptime(start, "%Y-%m-%dT%H:%M")).month
                day = "%s" % (datetime.strptime(start, "%Y-%m-%dT%H:%M")).day
                ans = "%s月%s日的加班單\n加班人姓名:%s\n加班起始時間:%s\n加班結束時間:%s\n總計加班時數:%s小時\n轉換成:%s\n請問是否核准加班申請?" % (
                    month, day, name, start, end, e, typ)
                confirm_template_message = TemplateSendMessage(
                    alt_text='加班申請審核通知',
                    template=ConfirmTemplate(
                        text=ans,
                        actions=[
                            PostbackTemplateAction(
                                label='是',
                                data=id + "_" + "yes"
                            ),
                            PostbackTemplateAction(
                                label='否',
                                data=id + "_" + "no"
                            )
                        ]
                    )
                )
                line_bot_api.push_message(boss_id, confirm_template_message)  # 將表單傳給主管
                return 0


            content = over_time(count)
            d['count'] = count
            with open(path, 'w') as f:
                json.dump(d, f, ensure_ascii=False)

            if content == "轉換費用或補假":
                name=d['answers'][1]
                start = d['answers'][2]
                end = d['answers'][3]
                year=(datetime.strptime(start, "%Y-%m-%dT%H:%M")).year
                month=(datetime.strptime(start, "%Y-%m-%dT%H:%M")).month
                sub = str((datetime.strptime(end, "%Y-%m-%dT%H:%M") - datetime.strptime(start, "%Y-%m-%dT%H:%M")))
                e = sub.split(":")[0]
                ###~~~~~~
                if int(e) == 0:
                    line_bot_api.push_message(user_id, TextSendMessage(text='加班時間未滿一小時不予以計算，請重新填單'))
                    d['count'] = 0
                    d['answers'] = []
                    with open(path, 'w') as f:
                        json.dump(d, f, ensure_ascii=False)
                    return 0  #
                ####~~~~~
                ###1221
                client=general()
                sheet = client.open('test').get_worksheet(6)
                col = sheet.col_values(3)
                for j in range(len(col)):
                    j = j + 1
                    if (sheet.cell(j, 1).value) == name and int((sheet.cell(j, 2).value)) == year:
                        if int((sheet.cell(j, 3).value)) == month:
                            total = int(sheet.cell(j, 4).value)
                            rest = 46 - total
                            if rest==0:
                                ans = "這個月的總加班時數已達到勞基法的上限46小時，無法再做加班"
                                d['count'] = 0
                                d['answers'] = []
                                with open(path, 'w') as f:
                                    json.dump(d, f, ensure_ascii=False)
                                    line_bot_api.push_message(user_id, TextSendMessage(text=ans))
                                return 0
                            if rest < int(e):
                                ans= "這個月的總加班時數會超過勞基法的上限46小時，你目前總加班時數為%s，只能再加%s小時，請重新填寫表單" % (total, rest)
                                d['count']=0
                                d['answers'] = []
                                with open(path, 'w') as f:
                                    json.dump(d, f, ensure_ascii=False)
                                    line_bot_api.push_message(user_id,TextSendMessage(text=ans))
                                return 0
                ############1221
                #if int(e) == 0:  #
                    #print("dont")
                    #line_bot_api.push_message(user_id, TextSendMessage(text='加班時間未滿一小時不予以計算，請重新填單'))  #
                    #d['count'] = 0  #
                    #d['answers'] = []  #
                    #with open(path, 'w') as f:
                        #json.dump(d, f, ensure_ascii=False)
                    #return 0  #

                ans = "總計加班時數共%s小時，請問想轉換成?" %e
                confirm_template_message = TemplateSendMessage(
                    alt_text='轉換費用與時數',
                    template=ConfirmTemplate(
                        text=ans,
                        actions=[
                            PostbackTemplateAction(
                                label='加班費',
                                data='加班費'
                            ),
                            PostbackTemplateAction(
                                label='補修時數',
                                data='補修時數'
                            )
                        ]
                    )
                )
                line_bot_api.reply_message(event.reply_token, confirm_template_message)
                return 0

            if content == "確認嗎?":
                name = d['answers'][1]
                start = d['answers'][2]
                end = d['answers'][3]
                sub = str((datetime.strptime(end, "%Y-%m-%dT%H:%M") - datetime.strptime(start, "%Y-%m-%dT%H:%M")))
                e = sub.split(":")[0]
                typ = d['answers'][4]
                month = "%s" % (datetime.strptime(start, "%Y-%m-%dT%H:%M")).month
                day = "%s" % (datetime.strptime(start, "%Y-%m-%dT%H:%M")).day
                ans = "%s月%s日的加班單\n加班人姓名:%s\n加班起始時間:%s\n加班結束時間:%s\n總計加班時數:%s\n轉換成:%s\n請問是否確認送出?" % (
                month, day, name, start, end, e, typ)
                confirm_template_message = TemplateSendMessage(
                    alt_text='確認表單',
                    template=ConfirmTemplate(
                        text=ans,
                        actions=[
                            PostbackTemplateAction(
                                label='確定',
                                data='確定'
                            ),
                            MessageTemplateAction(
                                label='重新填寫',
                                text='加班申請單'
                            )
                        ]
                    )
                )
                line_bot_api.reply_message(event.reply_token, confirm_template_message)
                return 0

            if content == "加班的起始時間是?":
                buttons_template_message = TemplateSendMessage(
                    alt_text='加班的起始時間',
                    template=ButtonsTemplate(
                        text='加班的起始時間是?',
                        actions=[
                            DatetimePickerTemplateAction(
                                label='請選擇',
                                data='datetime',
                                mode="datetime",
                                min=(datetime.now() - timedelta(days=30)).strftime("%Y-%m-%dT18:00"),
                                max=(datetime.now() + timedelta(days=30)).strftime("%Y-%m-%dT09:00")
                            )
                        ]
                    )
                )
                line_bot_api.reply_message(event.reply_token, buttons_template_message)
                return 0

            if content == "加班的結束時間是?":
                if data=="wrong" or data=="error":
                    data=answers[2]
                buttons_template_message = TemplateSendMessage(
                    alt_text='加班的結束時間',
                    template=ButtonsTemplate(
                        text='加班的結束時間是?',
                        actions=[
                            DatetimePickerTemplateAction(
                                label='請選擇',
                                data='datetime',
                                mode="datetime",
                                min=data,
                                max=(datetime.strptime(data, "%Y-%m-%dT%H:%M") + timedelta(hours=4)).strftime(
                                    "%Y-%m-%dT%H:%M")
                            )
                        ]
                    )
                )
                line_bot_api.reply_message(event.reply_token, buttons_template_message)
                return 0

            #line_bot_api.reply_message(
                #event.reply_token,
                #TextSendMessage(text=content))
            #return 0

        buttons_template = TemplateSendMessage(
            alt_text='目錄 template',
            template=ButtonsTemplate(
                title='申請單',
                text='請選擇',
                thumbnail_image_url='https://i.imgur.com/DzqsWsO.jpg',
                actions=[
                    MessageTemplateAction(
                        label='我要加班',
                        text='加班申請單'
                    ),
                    MessageTemplateAction(
                        label='我要請假',
                        text='請假申請單'
                    ),
                    MessageTemplateAction(
                        label='查詢加班紀錄',
                        text='加班查詢'
                    ),
                    MessageTemplateAction(
                        label='查詢請假紀錄',
                        text='請假查詢'
                    )
                ]
            )
        )
        line_bot_api.reply_message(event.reply_token, buttons_template)


if __name__ == '__main__':
    app.run()



