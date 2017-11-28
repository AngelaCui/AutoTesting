# -*- coding: utf-8 -*-
"""
Created on Fri Jun 16 10:21:09 2017

@author: admin
"""

import serial
import time
import threading, time

class DataBuffer():
    def __init__(self,data):
        self.num = data

class Counter(threading.Thread):
    def __init__(self,lock, threadName):
        super(Counter, self).__init__(name = threadName)
        self.lock = lock
        
    def run(self):
        global tmpdata
        try:
            ser.close()
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
            #data = ''
            tmpdata.data = [hex(ord(ser.read(1)))]
            print('first data is ',tmpdata.data)
            n = ser.inWaiting()
            print(n)
            #data = [hex(ord(ser.read(1)))]
            flag = 0
            while True:
                if ser.inWaiting() != 0 :
                    tmpdata.data.append(hex(ord(ser.read(1))))
                    flag = 1
                    
                #time.sleep(0.25)
                if flag == 1 and ser.inWaiting() == 0:
                    #print(data)
                    flag = 0
            #if data != '':
            #print (data)
            #else:
            #    print('no data')
            print('yes')
            ser.close()
        except:
            print(tmpdata.data)
            print('no')
            ser.close()
        
tmpdata = DataBuffer([])

for i in range(10):
    Counter(lock,'thread').start()
    print(tmpdata.num)
    time.sleep(1)