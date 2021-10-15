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
    try:
        #获取token
        res_token = requests.post(url='http://aiot.test.nlecloud.com:9190/nledu-cloud-sso-uaa/oauth/token',headers={'Content-Type':'application/x-www-form-urlencoded','Authorization':'Basic bmxlZHUtY2xvdWQtc3NvLXVhYToxMjM0NTY='},data={'grant_type':'password','username':arr_data1[i-1][0],'password':'123456'})
        token = json.loads(res_token.content)['data']['access_token']
        #print(token)
        #获取最新的学生实验任务id
        task_id = requests.get(url='http://52.130.248.0:9191/api-teaching/student/tasks/own/list?isReport=false',headers={'Content-Type':'application/x-www-form-urlencoded','Authorization':'bearer' + ' ' + token})
        id = json.loads(task_id.content)['data'][0]['id']
        #print(json.loads(task_id.content))
        #开启容器，并获取容器id
        res_container = requests.put(url='http://52.130.248.0:9191/api-teaching/student/tasks/do/start/' + id + '?ports=:22,:20805,:20905&type=2',headers={'Authorization':'bearer' + ' ' + token})
        container_id = json.loads(res_container.content)['data']
        # 获取容器的状态
        container_status = 'Creating'
        while container_status != 'Running':
            time.sleep(5)
            res_container = requests.get(url='http://52.130.248.0:9191/api-teaching/containers/detail?id=' + str(container_id),headers={'Authorization':'bearer' + ' ' + token})
            container = json.loads(res_container.content)
            container_status = container['data']['container']['state']
        if i == 1:
            ip1 = container['data']['virtualMachine']['ip']
            print('主虚拟机已开启')
        else:
            publicIp = container['data']['virtualMachine']['publicIp']
            #连接虚拟机
            ssh = paramiko.SSHClient()
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            ssh.connect(publicIp,'10001','账号','密码')
            stdin, stdout, stderr = ssh.exec_command('ls')
            content = stdout.read().decode()
            if content == '':
                stdin, stdout, stderr = ssh.exec_command('cmd=`cat /etc/profile|grep docEnv|head -1|sed \"s/\'//g\"`;$cmd;echo $docEnv;wget https://newlandblob.blob.core.chinacloudapi.cn/test/chirpstack-docker-cn.tgz;tar -zxf chirpstack-docker-cn.tgz;sed -i "s/d1-cdn.alpinelinux.org/mirrors.aliyun.com/g" /etc/apk/repositories;apk add sshpass;sshpass -p 密码 scp -r -P 10001 -o StrictHostKeyChecking=no root@' + ip1 + ':/root/chirpstack-docker-cn/data /root/chirpstack-docker-cn;cd chirpstack-docker-cn;docker-compose up -d')
                stdout.channel.set_combine_stderr(True)
                output = stdout.readlines()
            else:
                stdin, stdout, stderr = ssh.exec_command('cmd=`cat /etc/profile|grep docEnv|head -1|sed \"s/\'//g\"`;$cmd;echo $docEnv;cd chirpstack-docker-cn;docker-compose down;docker-compose up -d')        
                stdout.channel.set_combine_stderr(True)
                output = stdout.readlines()
            ssh.close()
            #登录获取chirpstack token,需要公网ip
            res_token = requests.post(url='http://' + publicIp + ':10002/api/internal/login',headers={'Content-Type':'application/json'},data=json.dumps({'email':'admin','password':'admin'}))
            token = json.loads(res_token.content)['jwt']
            #更新thingsboard网关设备Token
            tb = requests.put(url='http://' + publicIp + ':10002/api/applications/1/integrations/thingsboard',headers={'Content-Type':'application/json','Grpc-Metadata-Authorization':'Bearer' + ' ' + token},data=json.dumps({"integration": {"applicationID": "1","server": "{\"Server\":\"52.130.177.111:1883\",\"Token\":\"" + str(arr_data1[i-1][0]) + "\"}","thingsboard_token": str(arr_data1[i-1][0])}}))
            print(str(arr_data1[i-1][0])+ '已完成容器启动',time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    except:
        print(str(arr_data1[i-1][0])+ '容器启动失败！！！',time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))

