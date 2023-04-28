# -*- coding: utf-8 -*-
"""
Created on Tue Apr 25 15:01:34 2023

@author: User
"""
#https://www.taiwanpscc.org/#
#https://www.taiwanpscc.org/index.php?tag=member#

from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from apscheduler.schedulers.blocking import BlockingScheduler
import os
import time
import smtplib
import pandas as pd

def check_file_existence(file_path):
    return os.path.exists(file_path)


def Dict(df1, df2):
    # 删除df1和df2的第0行
    df1 = df1.drop(0)
    df2 = df2.drop(0)

    # 将df1和df2转换为字典
    dict1 = {} 
    dict2 = {}
    for row in df1.itertuples(index=False):
        key = row[0]
        values = list(row[1:])
        dict1[key] = values
    
    for row in df2.itertuples(index=False):
        key = row[0]
        values = list(row[1:])
        dict2[key] = values

    return dict1, dict2

def compare_dicts(dict1, dict2):
    # 检查dict1和dict2中的所有key是否相等
    if set(dict1.keys()) != set(dict2.keys()):
        print("The two dictionaries have different keys:")
        print(f"dict1 keys: {set(dict1.keys())}")
        print(f"dict2 keys: {set(dict2.keys())}")
        return False

    # 检查dict1和dict2中每个key的value是否相等
    for key in dict1:
        if dict1[key] != dict2[key]:
            print(f"The values for key '{key}' are different:")
            print(f"dict1 value: {dict1[key]}")
            print(f"dict2 value: {dict2[key]}")
            return False

    return True




def Mail(to):
    filename = "exported_data.xlsx"
    attachment = open(filename, "rb")
    
    msg = MIMEMultipart()
    msg['Subject']='Record.xlsx from PSCC Web'
    msg['From']='your_email'
    msg['To']=to
    
    # 加入郵件內容
    body = "這是一封通知郵件，附加了一個 xlsx 檔案作為附件(exported_data.xlsx)表示人數有變。"
    msg.attach(MIMEText(body, "plain"))
    
    part = MIMEBase('application', "octet-stream")
    part.set_payload((attachment).read())
    # 对附件进行编码
    encoders.encode_base64(part)
    # 添加附件的头部信息
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    # 将附件添加到邮件
    msg.attach(part)
    
    # 設定 SMTP 伺服器
    smtp_server = "smtp.gmail.com"
    smtp_port = 587

    # 設定寄件者的帳號密碼
    smtp_username = "your_email"
    smtp_password = "your_email_ps"

    # 連線 SMTP 伺服器
    smtp_conn = smtplib.SMTP(smtp_server, smtp_port)
    smtp_conn.ehlo()
    smtp_conn.starttls()
    smtp_conn.login(smtp_username,'your_email_key')

    # 寄送郵件
    status = smtp_conn.send_message(msg)
    print(status)

    # 關閉 SMTP 連線
    smtp_conn.quit()






def Read_Html():
    # 設置 ChromeDriver 的路徑
    chrome_driver_path = 'C:\chromedriver.exe'

    #創建 Chrome 選項對象
    chrome_options = Options()

    # 啟用無頭模式
    chrome_options.add_argument('--headless')

    # 啟用無痕模式
    chrome_options.add_argument('--incognito')


    # 創建 WebDriver 對象，使用 Chrome
    driver = webdriver.Chrome(chrome_driver_path,options=chrome_options)
    # 訪問目標網站
    driver.get('https://www.taiwanpscc.org/#')

    # 定位 <a> 標籤並點擊
    login_link = driver.find_element(By.XPATH, '//a[@href="#" and contains(@onclick, "LoginDlg")]')
    login_link.click()

    username ='your_web_username'
    ps ='your_web_ps'
    # 定位 <input> 元素並填入 username
    username_input = driver.find_element(By.ID, '_easyui_textbox_input3')
    username_input.send_keys(username)
    ps_input = driver.find_element(By.ID, '_easyui_textbox_input5')
    ps_input.send_keys(ps)
    time.sleep(3)

    # 定位 <a> 元素並點擊
    login_button= driver.find_element(By.XPATH, '//a[contains(@class, "easyui-linkbutton") and contains(@onclick, "Login()")]')
    login_button.click()

    read_button = driver.find_element(By.XPATH, '//span[contains(@class, "l-btn-left")]//span[contains(@class, "l-btn-text") and text()="我已閱讀"]')
    read_button.click()


    # 定位 <span> 元素并点击
    signup_button = driver.find_element(By.XPATH, '//span[contains(@class, "tt-inner")]/img[contains(@src, "images/tablet.png")]/..')
    signup_button.click()
    time.sleep(3)

    # 定位所有具有特定 class 的 <td> 元素
    td_elements = driver.find_elements(By.XPATH, '//td[contains(@class, "datagrid-td-rownumber")]/div[contains(@class, "datagrid-cell-rownumber")]')

    # 選擇並點擊最後一個 <td> 元素
    last_td = td_elements[-1]
    last_td.click()


    # 定位 <span> 元素並點擊
    view_button = driver.find_element(By.XPATH, '//span[contains(@class, "l-btn-text") and text()="檢視報名人數"]')
    view_button.click()

    html ='''<div id="regAmount_dlg20230425161232" class="easyui-dialog panel-body panel-body-noborder window-body" style="padding: 5px; width: 388.4px; height: 519px;" buttons="#regAmount_dlg20230425161232-buttons" data-options="
    			closed:true,
    			resizable:true,
    			collapsible:true,
    			iconCls:'icon-ok',
    			modal:true
    		" title="">
    		<div class="panel datagrid panel-htop" style="width: 250px;"><div class="datagrid-wrap panel-body panel-body-noheader" title="" style="width: 250px;"><div class="datagrid-view" style="width: 248.4px; height: 417px;"><div class="datagrid-view1" style="width: 0px;"><div class="datagrid-header" style="width: 0px; height: 32px;"><div class="datagrid-header-inner" style="display: block;"><table class="datagrid-htable" border="0" cellspacing="0" cellpadding="0" style="height: 32px;"><tbody></tbody></table></div></div><div class="datagrid-body" style="width: 0px; margin-top: 0px; height: 352px;"><div class="datagrid-body-inner"></div></div><div class="datagrid-footer" style="width: 0px;"><div class="datagrid-footer-inner" style="display: block;"><table class="datagrid-ftable" cellspacing="0" cellpadding="0" border="0"><tbody><tr class="datagrid-row" datagrid-row-index="0"></tr></tbody></table></div></div></div><div class="datagrid-view2" style="width: 248.4px;"><div class="datagrid-header" style="width: 248px; height: 32px;"><div class="datagrid-header-inner" style="display: block;"><table class="datagrid-htable" border="0" cellspacing="0" cellpadding="0" style="height: 32px;"><tbody><tr class="datagrid-header-row"><td class="" style="" field="name"><div class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-name" style="text-align: center;"><span>運動名稱</span><span class="datagrid-sort-icon"></span></div></td><td class="" style="" field="value"><div class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-value" resizable="false" style="text-align: center;"><span>人數限制</span><span class="datagrid-sort-icon"></span></div></td><td class="" style="" field="amount"><div class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-amount" resizable="false" style="text-align: center;"><span>目前報名人數</span><span class="datagrid-sort-icon"></span></div></td></tr></tbody></table></div></div><div class="datagrid-body" style="width: 248px; margin-top: 0px; overflow-x: hidden; height: 352px;"><table class="datagrid-btable" cellspacing="0" cellpadding="0" border="0"><tbody><tr id="regAmount20230425161232_datagrid-row-r4-2-0" datagrid-row-index="0" class="datagrid-row"><td field="name"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-name">田徑</div></td><td field="value"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-value">24</div></td><td field="amount"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-amount">16</div></td></tr><tr id="regAmount20230425161232_datagrid-row-r4-2-1" datagrid-row-index="1" class="datagrid-row  "><td field="name"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-name">游泳</div></td><td field="value"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-value">14</div></td><td field="amount"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-amount">5</div></td></tr><tr id="regAmount20230425161232_datagrid-row-r4-2-2" datagrid-row-index="2" class="datagrid-row  "><td field="name"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-name">射擊</div></td><td field="value"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-value">10</div></td><td field="amount"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-amount">2</div></td></tr><tr id="regAmount20230425161232_datagrid-row-r4-2-3" datagrid-row-index="3" class="datagrid-row  "><td field="name"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-name">射箭</div></td><td field="value"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-value">14</div></td><td field="amount"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-amount">1</div></td></tr><tr id="regAmount20230425161232_datagrid-row-r4-2-4" datagrid-row-index="4" class="datagrid-row  "><td field="name"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-name">健力</div></td><td field="value"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-value">6</div></td><td field="amount"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-amount">0</div></td></tr><tr id="regAmount20230425161232_datagrid-row-r4-2-5" datagrid-row-index="5" class="datagrid-row  "><td field="name"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-name">桌球</div></td><td field="value"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-value">21</div></td><td field="amount"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-amount">5</div></td></tr><tr id="regAmount20230425161232_datagrid-row-r4-2-6" datagrid-row-index="6" class="datagrid-row  "><td field="name"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-name">地板滾球</div></td><td field="value"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-value">15</div></td><td field="amount"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-amount">15</div></td></tr><tr id="regAmount20230425161232_datagrid-row-r4-2-7" datagrid-row-index="7" class="datagrid-row  "><td field="name"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-name">羽球</div></td><td field="value"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-value">20</div></td><td field="amount"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-amount">10</div></td></tr><tr id="regAmount20230425161232_datagrid-row-r4-2-8" datagrid-row-index="8" class="datagrid-row  "><td field="name"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-name">保齡球</div></td><td field="value"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-value">14</div></td><td field="amount"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-amount">1</div></td></tr><tr id="regAmount20230425161232_datagrid-row-r4-2-9" datagrid-row-index="9" class="datagrid-row  "><td field="name"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-name">輪椅網球</div></td><td field="value"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-value">7</div></td><td field="amount"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-amount">0</div></td></tr><tr id="regAmount20230425161232_datagrid-row-r4-2-10" datagrid-row-index="10" class="datagrid-row  "><td field="name"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-name">輪椅籃球</div></td><td field="value"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-value">10</div></td><td field="amount"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-amount">0</div></td></tr></tbody></table></div><div class="datagrid-footer" style="width: 248px;"><div class="datagrid-footer-inner" style="display: block;"><table class="datagrid-ftable" cellspacing="0" cellpadding="0" border="0"><tbody><tr class="datagrid-row" datagrid-row-index="0"><td field="name"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-name">總計:</div></td><td field="value"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-value">155</div></td><td field="amount"><div style="text-align:center;height:auto;" class="datagrid-cell regAmount20230425161232_datagrid-cell-c4-amount">55</div></td></tr></tbody></table></div></div></div><table id="regAmount20230425161232" style="padding: 1px; text-align: center; display: none;" class="easyui-datagrid datagrid-f" data-options="
    				fitColumns:false,
    				method:'get',
    				width:250,
    				showGroup:false,
    				scrollbarSize:1,
                    showFooter: true
    			">
    		</table><style type="text/css" easyui="true">
    .datagrid-header-rownumber{width:29px}
    .datagrid-cell-rownumber{width:29px}
    .regAmount20230425161232_datagrid-cell-c4-name{width:79px}
    .regAmount20230425161232_datagrid-cell-c4-value{width:79px}
    .regAmount20230425161232_datagrid-cell-c4-amount{width:79px}
    </style></div></div></div>
        
    </div>
    '''
    # 解析HTML
    soup = BeautifulSoup(html, 'html.parser')

    # 查找表格行
    rows = soup.find_all('tr', {'class': 'datagrid-row'})

    # 提取表格數據
    data = []
    for row in rows:
        columns = row.find_all('td')
        data.append([column.text for column in columns])

    # 創建DataFrame
    column_names = ['運動名稱', '人數限制', '目前報名人數']
    df = pd.DataFrame(data, columns=column_names)
    print(df)
    driver.quit()
    return df

def Run():
    file_path = "exported_data.xlsx"
    if not check_file_existence(file_path):
        df = Read_Html()
        df.to_excel(file_path,index=False)
        Mail("to_email")
        print("已寄送第一封通知")
    else:
        df1 = pd.read_excel('exported_data.xlsx', engine='openpyxl')
        df2 = Read_Html()
        dict1, dict2 = Dict(df1, df2)
        dict1 = {key: [int(x) for x in value] for key, value in dict1.items()}
        dict2 = {key: [int(x) for x in value] for key, value in dict1.items()}
        # 比较dict1和dict2是否相等
        if compare_dicts(dict1, dict2):
            print("The two dictionaries are identical.")
        else:
            print("The two dictionaries have differences.")
            df2.to_excel(file_path,index=False)
            Mail("to_email")
            print("已寄送新一封通知")
        
        
    
            
           

    

if __name__ == '__main__':
    
    # 创建调度器
    scheduler = BlockingScheduler()
    # 添加任务
    scheduler.add_job(Run, 'interval', minutes=1)
    # 开始调度器
    scheduler.start()

    
    
    
    
    
    
    
    


    
   
    
