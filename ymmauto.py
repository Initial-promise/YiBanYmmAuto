# -*- ecoding: utf-8 -*-
# @ModuleName: topicId
# @Author: Initial-C
# @Time: 2020/8/13 13:37

import urllib.request
import re
from selenium import webdriver
import time
import base64
import xlwt
import json

options = webdriver.ChromeOptions()
options.add_argument('--headless')
print("正在配置浏览器驱动")
# drive = webdriver.Chrome(executable_path="G:\\chromedriver_win32\\chromedriver.exe",options=options)
# driver2=webdriver.Chrome(executable_path="G:\\chromedriver_win32\\chromedriver.exe",options=options)
# drive = webdriver.Chrome(executable_path="chromedriver.exe",options=options)
# driver2=webdriver.Chrome(executable_path="chromedriver.exe",options=options)
drive=webdriver.Chrome(executable_path="G:\\chromedriver_win32\\chromedriver.exe")#驱动路径，options表示无头
driver2=webdriver.Chrome(executable_path="G:\\chromedriver_win32\\chromedriver.exe")

token=''

workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('ymm', cell_overwrite_ok=True)
now=0

file=open('topicId.txt','r')
topicId=file.readline()

def getToken():
    drive.get("https://ymm.yiban.cn")
    username=input("请输入你的账号:")
    password=input("请输入你的密码:")
    drive.find_element_by_id("account-txt").send_keys(username)
    drive.find_element_by_id("password-txt").send_keys(password)
    drive.find_element_by_id("login-btn").click()
    # print(drive.current_url)
    print("正在登录...")
    time.sleep(4)
    # print(drive.current_url)
    if(drive.current_url=='https://www.yiban.cn/login?go=https://ymm.yiban.cn/#/articles/list/1'):
        # drive.find_element_by_class_name('change-captcha').click()
        time.sleep(1)
        # imgCode=drive.find_element_by_xpath("//*[@id=\"login-box\"]/div[3]/img").get_attribute("src")
        # imgCode=bytes(imgCode,encoding='utf-8')
        # with open('./code.png','wb') as fn:
        #     fn.write(imgCode)
        # req=urllib.request.Request("https://www.yiban.cn/captcha/index?Thu%20Aug%2013%202020%2011:44:03%20GMT+0800%20(%E4%B8%AD%E5%9B%BD%E6%A0%87%E5%87%86%E6%97%B6%E9%97%B4)")
        # req.headers=headers
        # imgCode=urllib.request.urlopen(req).read()
        drive.save_screenshot('screen.png')
        code=input('请输入验证码:(screen.png)')
        drive.find_element_by_id('login-captcha').send_keys(code)
        drive.find_element_by_id("login-btn").click()
        # print(imgCode)
    cookies=drive.get_cookies()

    #print(cookies)
    token=cookies[1]['value']
    # print(token)
    return token

def saveData():
    global now
    col=("易班ID","用户昵称","文字内容","图片链接","点赞数","发布时间","真实姓名","手机号")
    for i in range(0,8):
        worksheet.write(now,i,col[i])
    now+=1
    downloadData()
    workbook.save('ymm.xls')

def login():
    driver2.get("https://mp.yiban.cn")
    driver2.find_element_by_id('account').send_keys(publicAccount)#此处变量为自己学校公共平台账号
    driver2.find_element_by_id('password').send_keys(publicPassword)

    driver2.find_element_by_id('loginSubmit').click()
    print("正在登录公共平台...")
    time.sleep(2)
    if (driver2.current_url != 'http://mp.yiban.cn/notice/index'):
        driver2.find_element_by_id('imageY').click()
        time.sleep(2)  # 这句很重要 需要等待刷新完再爬取图片
        imageCode = driver2.find_element_by_id('imageY').get_attribute('src')
        # driver.get('https://'+imageCode)
        # print(imageCode)
        # print(type(imageCode))
        imageCode = re.sub('data:image/jpg;base64,', '', imageCode)
        imageCode = bytes(imageCode, encoding='utf-8')
        with open('./code.jpg', 'wb') as fn:
            fn.write(base64.b64decode(imageCode))
        # print(imageCode)
        # print(type(imageCode))
        code = input("请输入验证码：(code.png)")
        driver2.find_element_by_id('inputyzm').send_keys(code)
        driver2.find_element_by_id('loginSubmit').click()

    time.sleep(3)
    driver2.find_element_by_xpath('//*[@id="menu_manage"]/dt/i[2]').click()
    print("正在准备信息查询...")
    time.sleep(1)
    driver2.find_element_by_link_text('信息查询').click()
    time.sleep(1)
    driver2.find_element_by_link_text('账号信息查询').click()

def search(userId):
    print("正在查询...")
    time.sleep(1)
    driver2.find_element_by_id('id').clear()
    driver2.find_element_by_id('id').send_keys(userId)
    driver2.find_element_by_id('search').click()
    name=''
    phoneNumber=''
    if(driver2.find_element_by_xpath('/html/body/div[2]/div/div[2]/div[2]/ul[2]/li[12]')):
        name=driver2.find_element_by_xpath('/html/body/div[2]/div/div[2]/div[2]/ul[2]/li[12]').text
        phoneNumber=driver2.find_element_by_xpath('/html/body/div[2]/div/div[2]/div[2]/ul[2]/li[10]').text
        print('姓名：'+name)
        print('手机号码：'+phoneNumber)

    else:
        print("未查询到结果")

    return name,phoneNumber

def downloadData():
    global now
    print(token)
    while True:
        num=input("请输入要爬取的页数(每页100条数据)")
        num=int(num)
        if num>=1 or num<=100:
            break
        else:
            print("请输入1-100的整数")
    print("正在爬取数据...")
    login()
    for page in range(num):
        url=f'https://ymm.yiban.cn/news/list/news?loginToken={token}&page={page+1}&size=999&topicId={topicId}'
        #         # print(url)
        #         # headers = {
        #         #     "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.105 Safari/537.36",
        #         #     'Cookie': f'analytics=9f3135270f8872fb; UM_distinctid=17260d21561257-0d557a691dc36f-d373666-e1000-17260d215624f5; YB_SSID=d442fb32b7c95a587020c8182d1f5900; Hm_lvt_3ea23aadce9fb15e9b2f91d294efee83=1595824773,1597287714; timezone=-8; BBSCORE_SESSID=lib7e2ulfdc3794qap1lu9prd6; yiban_user_token={loginToken}; loginToken={loginToken}; Hm_lpvt_3ea23aadce9fb15e9b2f91d294efee83=1597302142'
        #         # }
        #         # req=urllib.request.Request(url=url,headers=headers)
        #         # response=urllib.request.urlopen(req).read()
        #         # print(response)
        #         # print(type(response))
        drive.get(url)
        response=drive.find_element_by_xpath('/html/body/pre').text
        response=bytes(response,encoding='utf-8')
        # print(response)
        res=json.loads(response.decode('utf-8'))
        # print(res)
        yData=res['data']['list']
        for i in yData:
            userId=i['origin']['User_id']
            userNick=i['origin']['usernick']
            text=i['title']
            imagelist=','.join(i['origin']['imageList']).replace(',','，')
            likeNum=i['likeNum']
            time_local=time.localtime(int(i['createTime']))
            date=time.strftime("%Y-%m-%d %H:%M:%S",time_local)
            nameAndPhone=search(userId)
            col=(userId,userNick,text,imagelist,likeNum,date)
            for j in range(0,6):
                worksheet.write(now,j,col[j])
            for k in range(0,2):
                worksheet.write(now,k+6,nameAndPhone[k])
            now+=1

if __name__ == '__main__':
    token=getToken()
    print("main")
    print(token)
    saveData()
    # downloadData(token)
    print("数据已保存到ymm.xls")
    print("")
    drive.close()
    driver2.close()
    input("输入任意键退出 有问题联系作者")
