import requests,json
import os
import xlrd
import time
import paramiko

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
while True:
    #行数循环
    arr_data1 = []
    for i in range(1,rows_count):
        #列数循环
        arr_data0 = []
        for j in range(0,cols_count):
            #获取某一个单元格的内容
            row_col_data = tables.cell_value(i,j)
            arr_data0.append(str(int(row_col_data)))
        arr_data1.append(arr_data0)
        #获取token
        try:
            res_token = requests.post(url='http://aiot.test.nlecloud.com:9190/nledu-cloud-sso-uaa/oauth/token',headers={'Content-Type':'application/x-www-form-urlencoded','Authorization':'Basic bmxlZHUtY2xvdWQtc3NvLXVhYToxMjM0NTY='},data={'grant_type':'password','username':arr_data1[i-1][0],'password':'123456'})
            token = json.loads(res_token.content)['data']['access_token']
            #获取最新的学生实验任务id
            task_id = requests.get(url='http://52.130.248.0:9191/api-teaching/student/tasks/own/list?isReport=false',headers={'Content-Type':'application/x-www-form-urlencoded','Authorization':'bearer' + ' ' + token})
            id = json.loads(task_id.content)['data'][0]['id']
            #开启容器
            res_container = requests.put(url='http://52.130.248.0:9191/api-teaching/student/tasks/do/start/' + id + '?ports=:22,:20805,:20905&type=2',headers={'Authorization':'bearer' + ' ' + token})
            #发送心跳
            res_heart = requests.put(('http://52.130.248.0:9191/api-teaching/student/tasks/do/heartbeat/' + str(id) + '/master'),headers={'Authorization':'bearer' + ' ' + token})
            heart = json.loads(res_heart.content)
            heart_mess = heart['message']
            print(arr_data1[i-1][0],time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),heart_mess)
        except:
            print(arr_data1[i-1][0],time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),"心跳失败！！！")
    time.sleep(30)
