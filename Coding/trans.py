# -*- coding: utf-8 -*-
"""
Created on Mon Jun 12 15:36:33 2017

@author: admin
"""

# 移位
def Dataconvert(data,datalen):
    result = [data & 0xFF]
    for i in range(datalen-1):
        tmp = data >> 8
        result.append(tmp & 0xFF)
        data = tmp
    return result[::-1]
    
    
maindata = 0b1100010101000101
length = 2
s = Dataconvert(maindata,length)
for i in range(length):
    print(hex(s[i]))