import pandas as pd
from time import sleep
import pymysql
import datetime
import pyttsx3
data=pd.read_excel('/Users/v_yuanliangliang/Desktop/time_campare.xlsx', usecols=["Name","Clock_in_time","Clock_in_time_after_work"] )
Name = data["Name"]
Clock_in_time = data["Clock_in_time"]
Clock_in_time_after_work = data["Clock_in_time_after_work"]
conn = pymysql.connect(host="43.142.123.148", user="abc_mysql", password="12345678a", database="abc_mysql")
cur = conn.cursor()
for (Name,Clock_in_time,Clock_in_time_after_work) in zip(Name,Clock_in_time,Clock_in_time_after_work):
        cur.execute("insert into time_campare (Name,Clock_in_time,Clock_in_time_after_work) values ('%s','%s','%s')"%(Name,Clock_in_time,Clock_in_time_after_work))
# sql = 'create table time_campare(Name VARCHAR(255) not null,Clock_in_time VARCHAR(255) not null ,Clock_in_time_after_work VARCHAR(255) not null );'
        conn.commit()
cur.execute("select Name from  time_campare ;")
s1=cur.fetchall()
print("**********************************")
for i in s1:
    cur.execute("select Name ,Clock_in_time , Clock_in_time_after_work  from  time_campare where Name ='%s';"%(i))

    a = cur.fetchall()

    b = (a[0][0])   # 姓名

    b0 = (a[0][1])  # 早上打卡时间

    b1 = (a[0][2])   # 晚上打卡时间
    sleep(1)
    # print(b,b0,b1)

    Clock_in_time = (b0)
    Clock_in_time_after_work = (b1)

    morning_time_stamp = pd.Timestamp(Clock_in_time)   # 早上实际打卡时间
    morning_starnd_time= pd.Timestamp("2021/1/1 10:00")    # 早上标准打卡时间

    envening_time_stamp = pd.Timestamp(Clock_in_time_after_work)  # 晚上实际打卡时间
    enevning_starnd_time = pd.Timestamp("2021/1/1 19:00")    # 晚上标准打卡时间


    result1 = datetime.datetime.fromtimestamp(morning_time_stamp.timestamp()) #早上时间戳转换
    result11 = datetime.datetime.fromtimestamp(morning_starnd_time.timestamp()) #早上时间戳转换

    result2 = datetime.datetime.fromtimestamp(envening_time_stamp.timestamp()) #晚上时间戳转换
    result22 = datetime.datetime.fromtimestamp(enevning_starnd_time.timestamp()) #晚上时间戳转换

    # 创建  初始化
    engine = pyttsx3.init()
    # 说话

    if    result1  < result11:   # 早上打卡时间比对结果
        print(b +  "上班卡OK")
        engine.runAndWait()
        say1 = b +  "上班卡OK"
        engine.say(say1)

    else:
        print( b + "(先生/女士)，您已迟到！！")
        engine.runAndWait()
        say2 = b + "(先生/女士)，您已迟到！！"
        engine.say(say2)


    if    result2 > result22:   # 晚上打卡时间比对结果

        print(b +  "下班卡OK")
        engine.runAndWait()
        say3 = b +  "下班卡OK"
        engine.say(say3)

    else:

        print(b +   "(先生/女士)，您已早退！！！")
        engine.runAndWait()
        say4 = b +   "(先生/女士)，您已早退！！！"
        engine.say(say4)


    if  result1  < result11  and   result2 > result22:



        print("恭喜您，今日打卡正常 🐂。")


        say5 =  "恭喜您，今日打卡正常 🐂。"
        engine.say(say5)

        print("**********************************")
    else:
        print( "今日您已旷工！！！")
        engine.runAndWait()
        say6 = "今日您已旷工！！！"
        engine.say(say6)

        print("**********************************")

        sleep(2)
cur.execute("delete FROM time_campare ")
conn.commit()


engine.runAndWait()