# encoding: utf-8
import os
import requests
import re
import time
# import schedule  # 添加用于计划调度
import json
import chardet

import smtplib
import ssl
from email.mime.text import MIMEText
# 定时器
import multitimer


shift_to_beijing_time = 60*60*8


# 获取文件编码格式
# def file_encoding(file):
#     try:
#         if os.path.exists(file):
#             with open(file, 'rb') as f:
#                 encoding = chardet.detect(f.read())['encoding']
#                 f.close()
#                 return encoding
#         else:
#             pass
#     except Exception as e:
#         print(e)


# dict 转为对象属性访问， unbunch 将对象属性转为 json字符串
class FakeBunch(dict):
    def __init__(self, dic):
        # self.test = 'test'
        if isinstance(dic, dict):
            for item in dic.items():
                setattr(self, item[0], item[1])
        else:
            print('FakeBunch expected parameter is dict')
        self.keys = list(dic.keys())

    def unbunch(self):
        re_str = '{\n'
        for index,key in enumerate(self.keys):
            # re_str = re_str + '\"' + key + '\"' + ':' + str(getattr(self, key)) + '\n'
            value = getattr(self, key)
            if isinstance(value, str):
                re_str = re_str + '\"' + key + '\"' + ': ' + '"' + str(getattr(self, key)) + '"'
            else:
                re_str = re_str + '\"' + key + '\"' + ': ' + str(getattr(self, key))
            if index == len(self.keys)-1:
                re_str += '\n'
            else:
                re_str += ',\n'
        re_str += '}'
        #print(re_str)
        return re_str


def sendMail(text):
    with open('./setting.json', 'r', encoding='utf-8') as file:
        settingJson = json.loads(file.read())
        file.close()
    setting_obj = FakeBunch(settingJson)
    sender = setting_obj.mailSender
    receiver = setting_obj.mailReceiver
    mail_host = setting_obj.fromMailHost
    mail_password = setting_obj.mailPasswd
    mail_host_port = setting_obj.mailHostPort

    content = text
    message = MIMEText(content, 'plain', 'utf-8')
    # 主题
    message['Subject'] = '云南大学研究生院招生公告'
    # 发送方信息
    message['From'] = sender
    # 接收方
    message['To'] = receiver
    try:
        context = ssl.create_default_context()
        smtp_obj = smtplib.SMTP(mail_host, mail_host_port)
        smtp_obj.ehlo()
        smtp_obj.starttls(context=context)
        smtp_obj.ehlo()
        smtp_obj.login(sender, mail_password)
        smtp_obj.sendmail(sender, [receiver], message.as_string())
        print('Send successfully')
        smtp_obj.quit()
    except smtplib.SMTPException as e:
        print(time.strftime('%Y-%m-%d %H:%M:%S', time.gmtime(time.time()+shift_to_beijing_time)), ' error: ', e)


# def taskList():
#     schedule.every(5).seconds.do(main)
#     while True:
#         schedule.run_pending()
#         time.sleep(1)


def requestYnu(url):
    try:
        response = requests.get(url)
        response.encoding = 'utf-8'
        charsetPattern = re.compile('<head.*?charset=(.*?)">', re.S)
        charset = re.search(charsetPattern, response.text)
        response.encoding = charset[0]
        if response.status_code == 200:
            return response.text
    except requests.RequestException:
        return None


def paraseContent(html):
    patternParent = re.compile('<div id.*?招生工作.*?/h3>.*?<ul>(.*?)</ul>', re.S)
    html0 = re.search(patternParent, html)
    patternTarget = re.compile('href="(.*?)".*?title="(.*?)".*?\[(.*?)\]', re.S)
    items = re.findall(patternTarget, html0.groups()[0])
    for item in items:
        yield item[0]+' '+item[1]+' '+item[2]


def writeResultYnu(text):
    print('write start', time.strftime('%Y-%m-%d %H:%M:%S',time.gmtime(time.time()+shift_to_beijing_time)))
    with open('ynuInfo.txt', 'a', encoding='utf-8') as file:
        file.write('\n\n'+time.strftime('%Y-%m-%d %H:%M:%S',time.gmtime(time.time()+shift_to_beijing_time))+'----------------------\n')
        file.write(text)
        file.close()
    print('write end\n---------------')


def main():
    url = 'http://www.grs.ynu.edu.cn/'
    html = requestYnu(url)
    targetContent = paraseContent(html)
    news = ''
    current_month = time.gmtime(time.time()+shift_to_beijing_time).tm_mon
    current_day = time.gmtime(time.time()+shift_to_beijing_time).tm_mday
    # current_month = '04'
    # current_day = '8'
    flag = False
    for item in targetContent:
        date = item.split(' ')[2].split('-')
        if (int(date[0]) == int(current_month)) and (int(date[1]) == int(current_day)):
            news = '-----------------------------------------------------------------------------------------------\n' \
                   '|---news: ' + url + item + '\n' \
                   '------------------------------------------------------------------------------------------------\n'\
                                               '\n\n\n' + news
            flag = True
        # else:
        #     flag = False
        item = url + item + '\n'
        news += item
    if flag:
        with open('./setting.json', 'r', encoding='utf-8') as settingFile:
            settingDict = json.loads(settingFile.read())
            settingFile.close()

        settingJsonObj = FakeBunch(settingDict)
        today = time.strftime('%Y-%m-%d', time.gmtime(time.time()+shift_to_beijing_time))
        if today == settingJsonObj.sendMailDate:
            # with open('./ynuinfo.log', 'w', encoding='utf-8') as log_file:
            #     log_file.write('恭喜 今日发布了新消息' + \
            #                    time.strftime('%Y-%m-%d %H:%M:%S', time.gmtime(time.time()+shift_to_beijing_time)))
            pass
        else:
            with open('./setting.json', 'w', encoding='utf-8') as settingFile:
                # sendMail()
                settingJsonObj.sendMailDate = today
                # print(settingJsonObj.sendMainCount)
                updateSettingStr = settingJsonObj.unbunch()
                settingFile.write(updateSettingStr)
                settingFile.close()
        sendMail(news)
    elif int(time.gmtime(time.time()+shift_to_beijing_time).tm_hour) == 11 or int(time.gmtime(time.time()+shift_to_beijing_time).tm_hour) == 21:
        # writeResultYnu('今日没有更新最新招生公告' + time.ctime())
        with open('setting.json', 'r', encoding='utf-8') as settingFile:
            setting_obj = FakeBunch(json.loads(settingFile.read()))
            settingFile.close()
        if setting_obj.sendMailCount < 1:
            sendMail('今日没有更新最新招生公告' + time.strftime('%Y-%m-%d %H:%M:%S', time.gmtime(time.time()+shift_to_beijing_time)))
            setting_obj.sendMailCount = 1
            with open('setting.json', 'w', encoding='utf-8') as settingFile:
                settingFile.write(setting_obj.unbunch())
                settingFile.close()
        else:
            pass
    else:
        with open('setting.json', 'r', encoding='utf-8') as settingFile:
            setting_obj = FakeBunch(json.loads(settingFile.read()))
            settingFile.close()
        setting_obj.sendMailCount = 0
        with open('setting.json', 'w', encoding='utf-8') as settingFile:
            settingFile.write(setting_obj.unbunch())
            settingFile.close()
        print('今日暂无新消息', time.strftime('%Y-%m-%d %H:%M:%S', time.gmtime(time.time()+shift_to_beijing_time)))


timer = multitimer.MultiTimer(interval=5, function=main)
timer.start()