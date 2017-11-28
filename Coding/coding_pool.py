# -*- coding: utf-8 -*-
"""
Created on Wed Jun 21 09:59:30 2017

@author: admin
"""

# coding pool
'''Timer'''
import time
import os
def timer(interv,a):
    while True:
        try:
            time_remaining = interv - time.time()%interv
            time.sleep(time_remaining)
            b = run(a)
            return b
            
        except:
            return 'error'
        
def run(a):
    a = a*2
    return a
    
b = timer(1,2)
print(b)

'''CD-Class buffer试验'''
import threading
class data():
    def __init__(self,num):
        self.num = num

def Getdata(tmpdata):
    tmpdata.num +=1
    
testdata = data(2)

Getdata(testdata)
print(testdata.num)

'''串口持续收数，只要有数据输入就能存入data，始终循环'''
try:
    ser = serial.Serial( #下面这些参数根据情况修改
        port='COM2',
        baudrate=115200,
        parity=serial.PARITY_ODD,
        stopbits=serial.STOPBITS_ONE,
        bytesize=serial.EIGHTBITS,
        timeout = 2    # 若有timeout，则在timeout时间内没有输入的话就返回no。若无timeout，则一直等待直到有数据发过来为止
        )
   
    if (ser.isOpen() == False):
        ser.open()
        print('port is open from false')
    print('port is open')
    data = ''
    data = [hex(ord(ser.read(1)))]
    print('data is ',data)
    n = ser.inWaiting()
    print(n)
    #data = [hex(ord(ser.read(1)))]
    flag = 0
    while True:
        if ser.inWaiting() != 0 :
            data.append(hex(ord(ser.read(1))))
            flag = 1
        if flag == 1 and ser.inWaiting() == 0:
            print(data)
            flag = 0
    #if data != '':
    #print (data)
    #else:
    #    print('no data')
    print('yes')
    ser.close()
except:
    print(data)
    print('no')
    ser.close()
    
    
'''成功版本！！！！！！！！！！！'''
'''两个线程，一个counter不停从串口读数，存入缓存区Buffer中
另一个主线程'''
import serial, threading, time

class Buffer():
    def __init__(self,data):
        self.data = data
       
class Counter(threading.Thread):
    def __init__(self, lock, threadName, ser, buffer):
        super(Counter, self).__init__(name = threadName)  #注意：一定要显式的调用父类的初始化函数。
        self.lock = lock
        self.buffer = buffer
        self.ser = ser   
       
    def run(self):
        try:
            #self.buffer.data = [hex(ord(self.ser.read(1)))]
            while True:
                if self.ser.inWaiting() != 0:
                    self.buffer.data.append(hex(ord(self.ser.read(1))))
        except:
            return 'error'
  
lock = threading.Lock()
ser = serial.Serial( #下面这些参数根据情况修改
            port='COM4',
            baudrate=115200,
            parity=serial.PARITY_ODD,
            stopbits=serial.STOPBITS_ONE,
            bytesize=serial.EIGHTBITS,
            timeout = 2    # 若有timeout，则在timeout时间内没有输入的话就返回no。若无timeout，则一直等待直到有数据发过来为止
            )
tmpdata = Buffer([])
Counter(lock, 'thread',ser,tmpdata).start()

time.sleep(2)
print(tmpdata.data)
ser.close()