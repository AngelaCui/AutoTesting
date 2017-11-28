# -*- coding: utf-8 -*-
"""
Created on Mon Jun 19 09:12:32 2017

@author: admin
"""
import serial

class Get_Data():
    def __init__(self,port,baud):
        self.port = serial.Serial(port,
                                  baud,
                                  parity = serial.PARITY_ODD, 
                                  stopbits = serial.STOPBITS_ONE,
                                  bytesize = serial.EIGHTBITS)
        if (self.port.isOpen() == False):
            self.port.open()
        
    def closeport(self):
        self.port.close()
        
    def trigger(self):
        if self.port.inWaiting() != 0:
            return 1
        else:
            return 0
        
    def oridata(self,data):
        while self.trigger(self) == 1:
            data += [hex(ord(self.port.read(1)))]
        return data