from selenium import webdriver
import requests,json
import os
import xlrd
import time
import paramiko
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver import ActionChains

#获取EXCEL路径
BASE_PATH = os.path.split(os.path.dirname(os.path.abspath(__file__)))[0]
EXCEL_PATH = os.path.join(BASE_PATH,'web_aiot','aiot账号密码.xlsx')
#获取table数据
data = xlrd.open_workbook(EXCEL_PATH)
tables = data.sheets()[0]
#获取单元格行数
rows_count = tables.nrows
#获取第一行的列数
cols_count = len(tables.row_values(0))
#行数循环
arr_data1 = []
#driver存储
L = []
for i in range(1,rows_count):
    #列数循环
    arr_data0 = []
    for j in range(0,cols_count):
        #获取某一个单元格的内容
        row_col_data = tables.cell_value(i,j)
        arr_data0.append(str(int(row_col_data)))
    arr_data1.append(arr_data0)
    try:
        #获取token
        res_token = requests.post(url='http://aiot.test.nlecloud.com:9190/nledu-cloud-sso-uaa/oauth/token',headers={'Content-Type':'application/x-www-form-urlencoded','Authorization':'Basic bmxlZHUtY2xvdWQtc3NvLXVhYToxMjM0NTY='},data={'grant_type':'password','username':arr_data1[i-1][0],'password':arr_data1[i-1][1]})
        token = json.loads(res_token.content)['data']['access_token']
        option = webdriver.ChromeOptions()
        # 打开调试模式
        option.add_argument("--auto-open-devtools-for-tabs")
        #selenium打开的浏览器在程序结束时不退出
        option.add_experimental_option("detach", True)
        option.add_argument('--user-data-dir=C:\\Users\\qitian\\AppData\\Local\\Google\\Chrome\\User Data' + str(i))
        driver = webdriver.Chrome(options=option)
        L.append(driver)
        driver.get('http://aiot.test.nlecloud.com:9090/sch_edu?token=' + token)
        driver.maximize_window()
        #点击实验中心按钮
        WebDriverWait(driver,20,0.5).until(lambda e1: driver.find_element_by_xpath('//*[@id="app"]/div/section/head/div[2]/a[2]/span/a'))
        time.sleep(1)
        driver.find_element_by_xpath('//*[@id="app"]/div/section/head/div[2]/a[2]/span/a').click()
        #点击消息动态
        WebDriverWait(driver,10,0.5).until(lambda e1: driver.find_element_by_xpath('//*[@id="app"]/div/section/main/div/div[2]/div[2]/div/div[1]/div[2]/ul/li/span[1]'))
        time.sleep(0.5)
        driver.find_element_by_xpath('//*[@id="app"]/div/section/main/div/div[2]/div[2]/div/div[1]/div[2]/ul/li/span[1]').click()
        #点击继续任务
        WebDriverWait(driver,10,0.5).until(lambda e1: driver.find_element_by_xpath('//*[@id="app"]/div/section/main/div/div[2]/div[2]/div/div/div[2]/div[2]/div[1]/button/span'))
        driver.find_element_by_xpath('//*[@id="app"]/div/section/main/div/div[2]/div[2]/div/div/div[2]/div[2]/div[1]/button/span').click()
        #点击虚拟仿真按钮,sleep防止提示任务启动中
        time.sleep(1)
        WebDriverWait(driver,10,0.5).until(lambda e1: driver.find_element_by_xpath('//*[@id="app"]/div/div/div[2]/div/div/div[2]/div[1]/div[2]/div[2]/img'))
        driver.find_element_by_xpath('//*[@id="app"]/div/div/div[2]/div/div/div[2]/div[1]/div[2]/div[2]/img').click()
        #点击仿真开启实验按钮
        window_handles = driver.window_handles
        driver.switch_to_window(window_handles[-1])
        #等待+按钮加载完成后再去切换iframe
        WebDriverWait(driver,10,0.5).until(lambda e1: driver.find_element_by_xpath('//*[@id="nleEmulator"]/iframe'))
        iframe = driver.find_element_by_xpath('//*[@id="nleEmulator"]/iframe')
        driver.switch_to.frame(iframe)
        #防止页面刷新，找不到按钮元素执行click(),强制睡一会儿,刷新完后再执行元素
        WebDriverWait(driver,20,0.5).until(lambda e1: driver.find_element_by_xpath('/html/body/div/section/section/main/div/div/div[1]/div/div[1]/div/div[2]/div/div'))
        time.sleep(3)
        driver.find_element_by_xpath('/html/body/div/section/section/main/div/div/div[1]/div/div[1]/div/div[2]/div/div').click()
        #从frame切换到主文档
        driver.switch_to.default_content()
        #鼠标移动到+按钮
        WebDriverWait(driver,10,0.5).until(lambda e1: driver.find_element_by_xpath('//*[@id="app"]/div/div/div[1]/div[1]/div[2]'))
        time.sleep(1)
        ActionChains(driver).move_to_element(driver.find_element_by_xpath('//*[@id="app"]/div/div/div[1]/div[1]/div[2]')).perform()
        #点击终端按钮
        WebDriverWait(driver,10,0.5).until(lambda e1: driver.find_element_by_xpath('//*[@id="app"]/div/div/div[1]/div[1]/div[2]/ul/li[1]'))
        driver.find_element_by_xpath('//*[@id="app"]/div/div/div[1]/div[1]/div[2]/ul/li[1]').click()
        #print(str(arr_data1[i-1][0])+'web终端已开启',time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
        #鼠标移动到+按钮
        WebDriverWait(driver,10,0.5).until(lambda e1: driver.find_element_by_xpath('//*[@id="app"]/div/div/div[1]/div[1]/div[2]'))
        time.sleep(1)
        ActionChains(driver).move_to_element(driver.find_element_by_xpath('//*[@id="app"]/div/div/div[1]/div[1]/div[2]')).perform()
        #点击thingsboard按钮
        WebDriverWait(driver,10,0.5).until(lambda e1: driver.find_element_by_xpath('//*[@id="app"]/div/div/div[1]/div[1]/div[2]/ul/li[3]'))
        driver.find_element_by_xpath('//*[@id="app"]/div/div/div[1]/div[1]/div[2]/ul/li[3]').click()
        #等待frame框架加载完成后再去切换frame
        WebDriverWait(driver,10,0.5).until(lambda e1: driver.find_element_by_xpath('//*[@id="app"]/div/div/div[2]/div[3]/iframe'))
        iframe = driver.find_element_by_xpath('//*[@id="app"]/div/div/div[2]/div[3]/iframe')
        driver.switch_to.frame(iframe)
        #等待首页中的设备加载出来
        WebDriverWait(driver,10,0.5).until(lambda e1: driver.find_element_by_xpath('/html/body/tb-root/tb-home/mat-sidenav-container/mat-sidenav-content/div/div/tb-home-links/mat-grid-list/div/mat-grid-tile[4]/figure/mat-card/mat-card-content/mat-grid-list/div/mat-grid-tile[1]/figure/a'))
        #点击设备
        driver.find_element_by_xpath('/html/body/tb-root/tb-home/mat-sidenav-container/mat-sidenav-content/div/div/tb-home-links/mat-grid-list/div/mat-grid-tile[4]/figure/mat-card/mat-card-content/mat-grid-list/div/mat-grid-tile[1]/figure/a').click()
        time.sleep(2)
        #判断网关这一行元素是否存在
        flag = True
        try:
            driver.find_element_by_xpath('/html/body/tb-root/tb-home/mat-sidenav-container/mat-sidenav-content/div/div/tb-entities-table/mat-drawer-container/mat-drawer-content/div/div/div/table/tbody/mat-row')
        except:
            flag = False
        if flag == False:
            #等待+加载出来
            WebDriverWait(driver,10,0.5).until(lambda e1: driver.find_element_by_xpath('/html/body/tb-root/tb-home/mat-sidenav-container/mat-sidenav-content/div/div/tb-entities-table/mat-drawer-container/mat-drawer-content/div/div/mat-toolbar[1]/div/div[2]/button/span[1]/mat-icon'))
            time.sleep(1)
            #点击+
            driver.find_element_by_xpath('/html/body/tb-root/tb-home/mat-sidenav-container/mat-sidenav-content/div/div/tb-entities-table/mat-drawer-container/mat-drawer-content/div/div/mat-toolbar[1]/div/div[2]/button/span[1]/mat-icon').click()
            #等待添加新设备加载出来
            WebDriverWait(driver,10,0.5).until(lambda e1: driver.find_element_by_xpath('//*[@id="mat-menu-panel-1"]/div/button[1]'))
            #点击添加新设备
            driver.find_element_by_xpath('//*[@id="mat-menu-panel-1"]/div/button[1]').click()
            #等待设备名称加载出来
            time.sleep(2)
            WebDriverWait(driver,10,1).until(lambda e1: driver.find_element_by_xpath('//*[@id="mat-input-15"]'))
            #添加设备名称
            driver.find_element_by_xpath('//*[@id="mat-input-15"]').send_keys('gateway')
            #添加标签
            driver.find_element_by_xpath('//*[@id="mat-input-16"]').send_keys('网关')
            #点击是网关
            driver.find_element_by_xpath('//*[@id="mat-checkbox-4"]/label/div').click()
            #点击添加按钮
            driver.find_element_by_xpath('//*[@id="mat-dialog-0"]/tb-device-wizard/div/div[4]/button[2]/span[1]').click()
            #点击网关这一行
            WebDriverWait(driver,10,0.5).until(lambda e1: driver.find_element_by_xpath('/html/body/tb-root/tb-home/mat-sidenav-container/mat-sidenav-content/div/div/tb-entities-table/mat-drawer-container/mat-drawer-content/div/div/div/table/tbody/mat-row'))
            driver.find_element_by_xpath('/html/body/tb-root/tb-home/mat-sidenav-container/mat-sidenav-content/div/div/tb-entities-table/mat-drawer-container/mat-drawer-content/div/div/div/table/tbody/mat-row').click()
            #点击管理凭证
            WebDriverWait(driver,10,0.5).until(lambda e1: driver.find_element_by_xpath('//*[@id="mat-tab-content-0-0"]/div/tb-device/div[1]/button[4]/span[1]'))
            #防止页面刷新点击不到管理凭证按钮
            time.sleep(2)
            driver.find_element_by_xpath('//*[@id="mat-tab-content-0-0"]/div/tb-device/div[1]/button[4]/span[1]').click()
            #输入访问令牌
            WebDriverWait(driver,10,0.5).until(lambda e1: driver.find_element_by_xpath('//*[@id="mat-input-30"]'))
            driver.find_element_by_xpath('//*[@id="mat-input-30"]').clear()
            driver.find_element_by_xpath('//*[@id="mat-input-30"]').send_keys(arr_data1[i-1][0])
            #点击保存
            driver.find_element_by_xpath('//*[@id="mat-dialog-1"]/tb-device-credentials-dialog/form/div[3]/button[2]/span[1]').click()
            #点击x按钮,返回到设备页面
            WebDriverWait(driver,10,0.5).until(lambda e1: driver.find_element_by_xpath('/html/body/tb-root/tb-home/mat-sidenav-container/mat-sidenav-content/div/div/tb-entities-table/mat-drawer-container/mat-drawer/div/tb-entity-details-panel/tb-details-panel/header/mat-toolbar/div/button/span[1]/mat-icon'))
            driver.find_element_by_xpath('/html/body/tb-root/tb-home/mat-sidenav-container/mat-sidenav-content/div/div/tb-entities-table/mat-drawer-container/mat-drawer/div/tb-entity-details-panel/tb-details-panel/header/mat-toolbar/div/button/span[1]/mat-icon').click()
        print(str(arr_data1[i-1][0])+'ThingsBoard已开启',time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    except:
        print(str(arr_data1[i-1][0])+'web终端开启失败！！！',time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    finally:
        #从frame切换到主文档,专门为aiot.test.nlecloud.com:9090的record-timestamp使用，而不是TB。目的是防止35分钟后鼠标无操作导致任务实验退出
        driver.switch_to.default_content()
        for k in L:
            k.execute_script("localStorage.setItem('record-timestamp', new Date().getTime().toString())")
while True:
    for k in L:
        k.execute_script("localStorage.setItem('record-timestamp', new Date().getTime().toString())")
        time.sleep(1)
        
