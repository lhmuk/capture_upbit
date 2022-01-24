#-*- coding:utf-8 -*-
import warnings
warnings.filterwarnings("ignore")

import re
import time
import pprint
import json
#import urllib.parse
#import urllib.request
import os
import copy
#import getpass

import pprint

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import chromedriver_autoinstaller
import subprocess
import shutil

import pyautogui

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

from email.mime.image     import MIMEImage

from string import Template  # 문자열 템플릿 모듈


def go_web():
    try:
        shutil.rmtree(r"c:\Windows\temp\chrometemp")  #쿠키 / 캐쉬파일 삭제
    except FileNotFoundError:
        pass
    # 크롬 드라이버 이용시 업비트 통과가 안됨
    #driver = webdriver.Chrome('chromedriver.exe')
    
    # 2021-12-01
    # 업비트에서 IE 차단
    # driver = webdriver.Ie('IEDriverServer.exe')
    
    # 2021-12-01
    # 엣지로 변경
    #driver = webdriver.Edge('msedgedriver.exe')
    
    
    # chromedriver-autoinstaller 사용을 위해 크롬으로 다시 변경
    # 봇 감지를 피하기 위해 subprocess 처리
    subprocess.Popen(r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222 --user-data-dir="c:\Windows\temp\chrometemp   "') # 디버거 크롬 구동
    
    option = Options()
    option.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    
    #크롬드라이버 버전 확인
    chrome_ver = chromedriver_autoinstaller.get_chrome_version().split('.')[0]
    

    try:
        driver = webdriver.Chrome(f'./{chrome_ver}/chromedriver.exe', options=option)
    except:
        chromedriver_autoinstaller.install(True)
        driver = webdriver.Chrome(f'./{chrome_ver}/chromedriver.exe', options=option)
    
    #브라우저 width 조정
    driver.set_window_size(1450, 1060)
    
    #driver.get('https://upbit.com/exchange?code=CRIX.UPBIT.KRW-BTC')
    
    #driver.get('https://www.google.com/search?q=%EC%97%85%EB%B9%84%ED%8A%B8')    
    driver.get('https://search.naver.com/search.naver?where=nexearch&sm=top_hty&fbm=1&ie=utf8&query=%EC%97%85%EB%B9%84%ED%8A%B8')
    driver.implicitly_wait(5)
    #time.sleep(3)
    
    #time.sleep(6000)
    driver.find_element_by_css_selector('.nsite_name>a').click()
    
    time.sleep(5)
    #driver.Close();
    
    driver.get('https://upbit.com/exchange')
    #driver.implicitly_wait(3)
    driver.implicitly_wait(10)
        
    # 업비트에 접속했는지 검증
    while 1:
        try:
            #driver.find_element_by_css_selector('input[type=email]').send_keys('sodifx')
            #driver.find_element_by_css_selector('span.RveJvd.snByac').click()
            
            driver.find_element_by_css_selector('.ty02')
        except:
            time.sleep(1)
            continue
        break
    
    ###### 브라우저가 사이트에 정상 접속했고 이후로 루틴 수행 ######
    
    
    # 업비트 상단바 제거
    driver.execute_script("document.querySelector('#UpbitLayout>article').style.display='none'")
    driver.execute_script("document.querySelector('.mainB').style.padding=0")
    driver.execute_script("document.querySelector('.mainB>.ty02').style.top=0")
    driver.execute_script("document.querySelector('.mainB>.ty02 .tabB .scrollB>div:first-child').style.height='900px'")
    
    # 대기 시간 설정(단위: 초)
    timer = 10 * 60
    
    if timer >= 60 and timer%60 == 0:
        notice = str(int(timer/60)) + '분'
    elif timer < 60:
        notice = str(timer) + '초'
    else:
        notice = str(int(timer/60)) + '분 ' + str(timer%60) + '초'
    
    
    while 1:
        try:
            driver.save_screenshot(filename)
            email_msg = email_init()
            emailObj = emailSender(smtp_info)        
            
            # 이메일 발송
            if 1:
                emailObj.send_message(email_msg)
                print('메일 발송')
            else:
                print('!!!!테스트 모드(발송 안함)!!!')
            
            print(notice + ' 대기')
        except:
            print('오류 발생. 설정 시간(' + notice + ')만큼 대기 후 진행')
        
        time.sleep(timer)
        print('')




def makepng(filename):
    #im1 = pyautogui.screenshot()
    #im2 = pyautogui.screenshot('my_screenshot.png')
    im3 = pyautogui.screenshot(filename, region=(0, 0, 300, 300))



# 사용 안함
def stmp_send(smptp_info):
    smtp = smtplib.SMTP(smtp_info['smtp_server'],smtp_info['smtp_port'])
    
    print(type(smtp))
    #print(smtp.ehlo())
    print(smtp.starttls())    
    print(smtp.login(smtp_info['user_id'], smtp_info['user_pass']))

    # 제목, 본문 작성
    msg = MIMEMultipart()
    msg['Subject'] = smtp_info['subject']
    msg['From'] = smtp_info['user_id'] + '@naver.com'
    msg['To'] = smtp_info['mail_to']

    msg.attach(MIMEText('본문', 'plain'))
    msg.attach(MIMEText('이미지 파일 전송합니다.sdfoijsdf', 'plain'))

    # 파일첨부 (파일 미첨부시 생략가능)
    attachment = open(smtp_info['filename'], 'rb')
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= " + smtp_info['filename'])
    msg.attach(part)
    print(msg.as_string())

    smtp.sendmail(msg['From'], msg['to'], msg.as_string())
    
    smtp.quit()


class EmailHTMLImageContent:
    """e메일에 담길 이미지가 포함된 컨텐츠"""
    def __init__(self, str_subject, str_image_file_name, str_cid_name, template):
        """이미지파일(str_image_file_name), 컨텐츠ID(str_cid_name)사용된 string template과 딕셔너리형 template_params받아 MIME 메시지를 만든다"""
        assert isinstance(template, Template)
        #assert isinstance(template_params, dict)
        self.msg = MIMEMultipart()
        
        # e메일 제목을 설정한다
        self.msg['Subject'] = str_subject # e메일 제목을 설정한다
        
        # e메일 본문을 설정한다
        '''
        str_msg  = template.safe_substitute(**template_params) # ${변수} 치환하며 문자열 만든다
        '''
        str_msg  = template.safe_substitute(abc='') # ${변수} 치환하며 문자열 만든다
        mime_msg = MIMEText(str_msg, 'html')                   # MIME HTML 문자열을 만든다
        self.msg.attach(mime_msg)
        # e메일 본문에 이미지를 임베딩한다
        assert template.template.find("cid:" + str_cid_name) >= 0, 'template must have cid for embedded image.'
        assert os.path.isfile(str_image_file_name), 'image file does not exist.'
        with open(str_image_file_name, 'rb') as img_file:
            mime_img = MIMEImage(img_file.read())
            mime_img.add_header('Content-ID', '<' + str_cid_name + '>')
        self.msg.attach(mime_img)
        
    def get_message(self, str_from_email_addr, str_to_eamil_addrs):
        """발신자, 수신자리스트를 이용하여 보낼메시지를 만든다 """
        mm = copy.deepcopy(self.msg)
        mm['From'] = str_from_email_addr          # 발신자 
        mm['To']   = ",".join(str_to_eamil_addrs) # 수신자리스트 
        return mm

class emailSender:
    """e메일 발송자"""
    def __init__(self, smtp_info):
        """호스트와 포트번호로 SMTP로 연결한다 """
        self.ss = smtplib.SMTP(smtp_info['smtp_server'], smtp_info['smtp_port'])
        
        # TLS(Transport Layer Security) 시작
        print(self.ss.starttls())
        print(self.ss.login(smtp_info['user_id'], smtp_info['user_pass']))
        
        self.email_from = smtp_info['mail_from']
        self.email_to = smtp_info['mail_to']
    
    def send_message(self, emailContent):
        """e메일을 발송한다 """
        cc = emailContent.get_message(self.email_from, self.email_to)
        #print(cc)
        self.ss.send_message(cc, self.email_from, self.email_to)
        del cc



def email_init():
    str_subject = smtp_info['subject']
    #str_subject = '인폼'
    template = Template("""<html>
                                <head></head>
                                <body>
                                    <img src="cid:my_image1">
                                </body>
                            </html>""")
    #template_params       = {'NAME':'Son'}
    str_cid_name          = 'my_image1'

    #emailHTMLImageContent = EmailHTMLImageContent(str_subject, filename, str_cid_name, template, template_params)
    emailHTMLImageContent = EmailHTMLImageContent(str_subject, filename, str_cid_name, template)
    return emailHTMLImageContent

#exit()
#makepng(filename)

filename = 'capture.png'

thisf = os.path.abspath( __file__ )
jsonfile = re.sub('.py', '.json', thisf)
with open(jsonfile, 'r', encoding='UTF-8') as f:
    smtp_info = json.load(f)

# json 파일 구성
# {
#   "smtp_server" : "",
#   "smtp_port" : ,
#   "user_id" : "",
#   "user_pass" : "",
#   "mail_from" : "",
#   "mail_to" : [""],
#   "subject" : ""
# } 

smtp_info['filename'] = filename

pprint.pprint(smtp_info)

# 사용 안함
#stmp_send(smtp_info)
#exit()

# 웹 접속 및 메일 발송 시작
go_web()