print('环境初始化中，请稍后......(苏纳预约工具&&powered by@徐法师)')


import requests                                                                                 #加载需要的模块
import re
import pandas as pd
import time
import random


print('初始化完成！！！')


#T_Week = int(input('预约本周请输入0，预约下周请输入1：'))【未开发】
T_Day = int(input('请输入日期(1-6)：')) - 1
Length_Try = int(input('请输入理想的时段个数：'))
#Length_Min = int(input('请输入可接受的最少时段个数(小于等于29且不大于理想时段数的正整数)：'))【未开发】
    

_FacilityID_ = 231                                                                              #输入设备的id，在该设备预约时刻表网址的最后

'''（三个点之内的内容为注释，勿cue）【未开发】
T_Earliest = (input('请输入最早时段(缺省为0)：'))
if T_Earliest == '':
    T_Earliest = 0
else:
    T_Earliest = int(T_Earliest)
    
T_Latest = (input('请输入最晚时段(缺省为28)：'))
if T_Latest == '':
    T_Latest = 28
else:
    T_Latest = int(T_Latest)
'''

    
print('开始获取预约时刻表，请稍后......')


session = requests.Session()


headers = {
	'User-Agent':"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.90 Safari/537.36"           #【模拟用户UA，让服务器认为你是个浏览器用户】
}


login_url = 'http://58.210.56.163:8085/LoginNew.aspx'
page_text_login = session.get(url=login_url,headers=headers).text


ex = 'VIEWSTATE" value="(.*?)" />'                                                              #利用re模块的遍历功能，从GET命令获取到的页面html中获取参量的值，并添加到后面的字典中，这三个参量每次都不一样
VIEWSTATE_login = re.findall(ex,page_text_login,re.S)
ex = 'VIEWSTATEGENERATOR" value="(.*?)" />'
VIEWSTATEGENERATOR_login = re.findall(ex,page_text_login,re.S)
ex = 'EVENTVALIDATION" value="(.*?)" />'
EVENTVALIDATION_login = re.findall(ex,page_text_login,re.S)


data_login = {                                                                                  #POST提交命令中的参量，存在{字典}中，前面遍历得来的三个参量的值携带有前一个网页的信息
	'__VIEWSTATE': VIEWSTATE_login, 
	'__VIEWSTATEGENERATOR': VIEWSTATEGENERATOR_login,
	'__EVENTVALIDATION': EVENTVALIDATION_login,
	'UserName': 'xilufan',                                                                   #这一行的单引号内填写账户
	'Pwd': '111111',                                                                        #这一行的单引号内填写密码
	'btnLogin.x': '1',
	'btnLogin.y': '1',
	}                                                                                       #如果脚本不能用了，检查是不是网站后台针对POST命令增加了新的参量需求 


session.post(url=login_url,headers=headers,data=data_login)                                     #POST指令提交，网址，headers,和字典


UserDefault_url = 'http://58.210.56.163:8085/UserDefault.aspx'                                  
page_text_UserDefault = session.get(url=UserDefault_url,headers=headers).text

#with open('./登录.html','w',encoding='utf-8') as fp:                                           #可以保存数据，用于调试
#	fp.write(page_text_UserDefault)

	
ex = 'VIEWSTATE" value="(.*?)" />'
VIEWSTATE_UserDefault = re.findall(ex,page_text_UserDefault,re.S)
ex = 'VIEWSTATEGENERATOR" value="(.*?)" />'
VIEWSTATEGENERATOR_UserDefault = re.findall(ex,page_text_UserDefault,re.S)
ex = 'EVENTVALIDATION" value="(.*?)" />'
EVENTVALIDATION_UserDefault = re.findall(ex,page_text_UserDefault,re.S)



data_UserDefault = {                                                                            #选课题组的参量字典
	'ToolkitScriptManager1_HiddenField': ';;AjaxControlToolkit, Version=3.5.40412.0, Culture=neutral, PublicKeyToken=28f01b0e84b6d53e:zh-CN:1547e793-5b7e-48fe-8490-03a375b13a33:475a4ef5:effe2a26:1d3ed089:5546a2b:497ef277:a43b07eb:751cdd15:dfad98a5:3cf12cf1',
        '__EVENTTARGET': 'ctl00$ContentPlaceHolder1$rptSubject$ctl00$btnSubject',
        '__EVENTARGUMENT': '', 
	'__VIEWSTATE': VIEWSTATE_UserDefault,
	'__VIEWSTATEGENERATOR': VIEWSTATEGENERATOR_UserDefault,
        '__VIEWSTATEENCRYPTED': '',
	'__EVENTVALIDATION': EVENTVALIDATION_UserDefault,
        'ctl00$ContentPlaceHolder1$txtId': '',
        'ctl00$ContentPlaceHolder1$txtNo': '' ,
        'ctl00$ContentPlaceHolder1$ddlStatus': 1,
        'ctl00$ContentPlaceHolder1$ddlSubjectId': 45,
        'ctl00$ContentPlaceHolder1$txtAmount': '',
        'ctl00$ContentPlaceHolder1$txtDueDate': '',
        'ctl00$ContentPlaceHolder1$txtRemark': '',
        'ctl00$ContentPlaceHolder1$ddlMaterial': 1,
        'ctl00$ContentPlaceHolder1$txtMaterialQty': '',
	'ctl00$ContentPlaceHolder1$txtCMachineFee': '0',
	'ctl00$ContentPlaceHolder1$txtCMaterialFee': '0',
	'ctl00$ContentPlaceHolder1$txtCOtherFee': '0',
	'ctl00$ContentPlaceHolder1$txtCOtherFeeRemark': '',
	'ctl00$ContentPlaceHolder1$txtAdminLoginId': '',
	'ctl00$ContentPlaceHolder1$txtAdminPwd': '',
        'ctl00$ContentPlaceHolder1$txtOperator': '',
        'ctl00$ContentPlaceHolder1$hid_operatorId': '',
        'ctl00$ContentPlaceHolder1$txtSupervisor': '',
        'ctl00$ContentPlaceHolder1$hid_SupervisorId': '',
	'ctl00$ContentPlaceHolder1$HiddenField1': '',
	}
	

session.post(url=UserDefault_url,headers=headers,data=data_UserDefault)                         #这个POST的目的是选择课题组


Schedule_url = 'http://58.210.56.163:8085/Schedule.aspx?id='+str(_FacilityID_)                  #这一行的网址是预约设备的时刻表网址，实际上只有最后的id会变
page_text_Schedule1 = session.get(url=Schedule_url,headers=headers).text


ex = 'VIEWSTATE" value="(.*?)" />'
VIEWSTATE_Schedule = re.findall(ex,page_text_Schedule1,re.S)
ex = 'VIEWSTATEGENERATOR" value="(.*?)" />'
VIEWSTATEGENERATOR_Schedule = re.findall(ex,page_text_Schedule1,re.S)
ex = 'EVENTVALIDATION" value="(.*?)" />'
EVENTVALIDATION_Schedule = re.findall(ex,page_text_Schedule1,re.S)


data_Schedule = {
	'ToolkitScriptManager1_HiddenField': '',
        '__EVENTTARGET': 'ctl00$ContentPlaceHolder1$AspNetPager1',
        '__EVENTARGUMENT': '2',                                                        
	'__VIEWSTATE': VIEWSTATE_Schedule,
	'__VIEWSTATEGENERATOR': VIEWSTATEGENERATOR_Schedule,
	'__EVENTVALIDATION': EVENTVALIDATION_Schedule,
	}


page_text_Schedule2 = session.post(url=Schedule_url,headers=headers,data=data_Schedule).text    #相当于点设备的预约按键，进入时刻表页面


pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

#df1 = pd.DataFrame()
#df1 = pd.concat([df1, pd.read_html(page_text_Schedule1)[0].iloc[::,1:]])


df2 = pd.DataFrame()                                                                            #用pd模块将html字符格式的表格提取成矩阵(行列格式)
df2 = pd.concat([df2, pd.read_html(page_text_Schedule2)[0].iloc[::,1:]])

print(df2)

Length_Temp = 0
T_Start = 0
T_Start_Option = []


'''-----------------------------------------------------------------------------------------------------------------------
------------------------------------------------判断可预约时刻的模块------------------------------------------------------
这一部分实现：在从选中的日期(T_Day)对应的列中筛选满足时刻长度(Length_Try)的连续空闲时段；
从第一个时刻开始遍历所选日期这一列，nan(NotaNumber)代表空，若为空则计数Length_Temp加1，若不为空则Length_Temp归零，不断循环遍历这一列；
当Length_Temp等于Length_Try时，说明有连续Length_Try个时段是nan了，代表有这么多个连续时段可选，也就筛选出来了；
将符合要求的时段的第一个时间戳，存在T_Start_Option里'''

for i in range(len(df2)):                                                                       
    if str(df2.iat[i,T_Day]) == 'nan':
        Length_Temp += 1
        if Length_Temp == Length_Try:
            T_Start_Option.append(T_Start)
            T_Start += 1
            Length_Temp -= 1
    else:
        T_Start = i+1
        Length_Temp = 0
print('可选的预约开始时段包括'+str(T_Start_Option))


'''------------------------------------------------------完----------------------------------------------------------------
---------------------------------------------------------------------------------------------------------------------------'''




T_Select = int(input('请输入想要的开始时段：'))

print('开始检测预约状态并尝试预约，请稍后...')


while str(df2.iat[10,T_Day]) != 'nan':                                                       #触发预约的条件，因为纳米所的预约系统，在12点整点的会把时刻表清空几秒钟，利用这一特点，让预约程序不断循环获取时刻表，当一个本身已经被预约的时刻，例如[?,T_Day]变成void（也就是nan/not a number）后，预约程序跳出循环，进行后面的预约操作
    ex = 'VIEWSTATE" value="(.*?)" />'
    VIEWSTATE_Schedule = re.findall(ex,page_text_Schedule2,re.S)
    ex = 'VIEWSTATEGENERATOR" value="(.*?)" />'
    VIEWSTATEGENERATOR_Schedule = re.findall(ex,page_text_Schedule2,re.S)
    ex = 'EVENTVALIDATION" value="(.*?)" />'
    EVENTVALIDATION_Schedule = re.findall(ex,page_text_Schedule2,re.S)
    data_Schedule = {
            'ToolkitScriptManager1_HiddenField': '',
            '__EVENTTARGET': 'ctl00$ContentPlaceHolder1$AspNetPager1',
            '__EVENTARGUMENT': '2',
            '__VIEWSTATE': VIEWSTATE_Schedule,
            '__VIEWSTATEGENERATOR': VIEWSTATEGENERATOR_Schedule,
            '__EVENTVALIDATION': EVENTVALIDATION_Schedule,
            }
    page_text_Schedule2 = session.post(url=Schedule_url,headers=headers,data=data_Schedule).text
    df2 = pd.DataFrame()
    df2 = pd.concat([df2, pd.read_html(page_text_Schedule2)[0].iloc[::,1:]])
    t = random.uniform(0.7,1.2)                                                             #循环体每次休眠随机时长，定义上下限
    time.sleep(t)                                                                           #休眠命令，避免拥塞服务器



ex = 'VIEWSTATE" value="(.*?)" />'
VIEWSTATE_Schedule2 = re.findall(ex,page_text_Schedule2,re.S)
ex = 'VIEWSTATEGENERATOR" value="(.*?)" />'
VIEWSTATEGENERATOR_Schedule2 = re.findall(ex,page_text_Schedule2,re.S)
ex = 'EVENTVALIDATION" value="(.*?)" />'
EVENTVALIDATION_Schedule2 = re.findall(ex,page_text_Schedule2,re.S)


data_Schedule2 = {
	'ToolkitScriptManager1_HiddenField': '',
        '__EVENTTARGET': '',
        '__EVENTARGUMENT': '',
	'__VIEWSTATE': VIEWSTATE_Schedule2,
	'__VIEWSTATEGENERATOR': VIEWSTATEGENERATOR_Schedule2,
	'__EVENTVALIDATION': EVENTVALIDATION_Schedule2,
#        'ctl00$ContentPlaceHolder1$GridView1$ctl02$chk_0_0': 'on',                         #这一行与预约的时段相关，ctl和chk后的数字都需要根据时段来变化，n个时段对应n行由于后面用了扩展语句实现时段的添加，所以此处注释掉，留在这里是为了展示post请求里需要提交的字典的完整格式
        'ctl00$ContentPlaceHolder1$btnNext2': '下 一 步' ,
	}


for j in range(Length_Try):
    data_Schedule2['ctl00$ContentPlaceHolder1$GridView1$ctl'+str(T_Select+j+2).zfill(2)+'$chk_'+str(T_Day + 1)+'_'+str(T_Select + j)] = 'on'          #循环进行每次一行的扩展字典，实现时段的添加
    

session.post(url=Schedule_url,headers=headers,data=data_Schedule2)


Submit_url = 'http://58.210.56.163:8085/Submit.aspx?id='+str(_FacilityID_)                  #提交的网页，只有最后的id会根据设备不同而改变
page_text_Submit = session.get(url=Submit_url,headers=headers).text



#with open('./Submit.html','w',encoding='utf-8') as fp:
#	fp.write(page_text_Submit)



ex = 'VIEWSTATE" value="(.*?)" />'
VIEWSTATE_Submit = re.findall(ex,page_text_Submit,re.S)
ex = 'VIEWSTATEGENERATOR" value="(.*?)" />'
VIEWSTATEGENERATOR_Submit = re.findall(ex,page_text_Submit,re.S)
ex = 'EVENTVALIDATION" value="(.*?)" />'
EVENTVALIDATION_Submit = re.findall(ex,page_text_Submit,re.S)


data_Submit = {
        'ToolkitScriptManager1_HiddenField': ';;AjaxControlToolkit, Version=3.5.40412.0, Culture=neutral, PublicKeyToken=28f01b0e84b6d53e:zh-CN:1547e793-5b7e-48fe-8490-03a375b13a33:475a4ef5:effe2a26:1d3ed089:5546a2b:497ef277:a43b07eb:751cdd15:dfad98a5:3cf12cf1',
        '__EVENTTARGET': '',
        '__EVENTARGUMENT': '',
	'__VIEWSTATE': VIEWSTATE_Submit,
	'__VIEWSTATEGENERATOR': VIEWSTATEGENERATOR_Submit,
	'__EVENTVALIDATION': EVENTVALIDATION_Submit,
        'ctl00$ContentPlaceHolder1$ctl00$txtValue': 'AZ5214',                               #这部分信息是针对MA6光刻填写的，用于其他设备工艺的预约时请修改，格式在网页的post请求里找
        'ctl00$ContentPlaceHolder1$ctl01$txtValue': '小片',
        'ctl00$ContentPlaceHolder1$ctl02$txtValue': '4inch',
        'ctl00$ContentPlaceHolder1$txtResRemark': '',
        'ctl00$ContentPlaceHolder1$btnSubmit': '提交',
        'ctl00$ContentPlaceHolder1$txtTechnicName': '',
        'ctl00$ContentPlaceHolder1$txtTechnicRemark':'', 
	}



session.post(url=Submit_url,headers=headers,data=data_Submit)                               #最终的预约提交指令，调试时请注释掉，避免误操作提交预约


print('预约已完成,预约时段为'+' 周'+str(T_Day + 1)+' 的 '+str(T_Select + 1)+'-'+str(T_Select + Length_Try))



#with open('./Schedule.html','w',encoding='utf-8') as fp:
#	fp.write(page_text_Schedule)


