# -*- coding: utf-8 -*-
"""
Created on Tue Jun 13 09:25:19 2017

@author: admin
"""
#串口收数
import time
import serial

ser = serial.Serial(
    port = '',
    baudrate = 115200)


# 数据传输一次传12800bytes信息，每8位数据位后有一位停止位
# 传输一次时间为250ms
# 小端字节





# 按each byte transform
# 此时data首先按照顺序依次排列
def Dataconvert(data,datalen):
    result = [data & 0xFF]
    for i in range(datalen-1):
        tmp = data >> 8
        result.append(tmp & 0xFF)
        data = tmp
    return result[::-1]

# 校验帧头帧尾
def Finddata(data):
    # 寻找帧头帧尾，存入缓存序列中
# 问题：data格式。设定查找范围为 波特率/9 = 12800
    for i in range(12800):
        if (data[i] == 0xFF):
            if data[i+1] == 0xFB:
                datahead = i
                break
        else: 
            return "No head found"
    for j in range(datahead+15,12800):
        if data[j] == 0xFF:
            if data[j+1] == 0xFE:
                dataend = j+1
                break
        else:
            dataend = data.index(data[-1]) 
            
    datapack = data[datahead:dataend+1]
    maindata = data[datahead+11:dataend-3]
    
    return datapack, maindata
    


# 合并相同包号的数据帧并排序
def Combine(datapack,packbase): # packbase为dict，key是包号
    packnum = datapack[6]
    packorder = datapack[8]
    if packnum in packbase.keys():
        packbase[packnum][packorder-1] = datapack
        
    else:
        packtotal = datapack[7]
        packbase[packnum] = [0]*packtotal
        packbase[packnum][packorder-1] = datapack


