import os
import xlwings as xw
import sys
from pymodbus.client.sync import ModbusSerialClient as ModbusClient
import time
import json
import xlsxwriter 


def main():
    
    print("Written by Furkan Ciylan \nContact me: furkanciylan@gmail.com\nGithub: /HawkPhantom")
    
    a = input("Do you want to configure the software? Y/N:")
    
    if (a == "n" or a == "N" or a == "no"):
        conffile_test = input("Do you have a custom configuration file? Y/N:")
    else:
        conffile_test = ""
    
    return(modbus(a,conffile_test))   
        
def modbus(a,conffile_test):
    name = input("Name this session:")
    
    if (conffile_test == "Y" or conffile_test == "y" or conffile_test =="yes"):
        conffile = input("Please write the path of configuration file:")
    else:
        conffile = ("configuration.json")
    
    
    try:
        with open("{}".format(conffile), 'r') as database:      
            raw_data = json.load(database)
            method = raw_data["method"]
            port = raw_data["port"]
            timeout = int(raw_data["timeout"])
            baudrate = int(raw_data["baudrate"])
            stopbits = int(raw_data["stopbits"])
            parity = raw_data["parity"]
            bytesize = int(raw_data["bytesize"])
            number_of_registers = int(raw_data["number_of_registers"])
            database.close()
            
    except:
        print("Configuration File Cannot be Found")
        try_again = input("Do you want to specify another location? Y/N")
        if(try_again == "y" or try_again == "Y" or try_again == "yes"):
            modbus(a,conffile_test)
        
    if (a =="y" or a=="Y" or a=="yes"):
            
        print("Default Values Are: \n method:rtu \n port:COM6 \n timeout:10ds \n baudrate:19200 \n stopbits:1 \n parity:Even \n bytesize:8 \n number of input registers:7 \n")
            
        method_1 = input("Please enter the method (ascii or rtu):")
        if (method_1 != ""):
            method = method_1
            raw_data["method"] = method
            
        port_1 = input("Please specify the COM port (COM7):")
        if (port_1 != ""):
            port=port_1
            raw_data["port"] = port

        try:
            timeout_1 = int(input("Please enter timeout value in decisecond:"))
            timeout = timeout_1
            raw_data["timeout"] = str(timeout_1)
            
        except:
            pass
                    
        try:
            baudrate_1 = int(input("Please enter the baudrate (19200):"))
            baudrate = baudrate_1
            raw_data["baudrate"] = str(baudrate_1)
        except:
            pass
            
        try:
            stopbits_1 = int(input("Please enter the number of stopbits (1/2):"))
            stopbits = stopbits_1
            raw_data["stopbits"] = stopbits_1
        except:
            pass
            
            
        parity_1 = input("Please enter the parity (E for even/O for odd/N for none):")
        parity_1 = parity_1.lower()
        if (parity_1 != ""):
            if(parity_1 == "even"):
                 parity_1 = "E"
            elif(parity_1 == "odd"):
                parity_1 = "O"
            elif(parity_1=="none"):
                parity_1 = "N"
                    
            parity=parity_1
            raw_data["parity"] = parity
                    
            
        try:
            bytesize_1 = int(input("Please enter the bytesize:"))
            raw_data["bytesize"] = bytesize_1
            bytesize = bytesize_1
        except:
            pass
            
            
        try:
            number_of_registers_1 = int(input("Please enter the number of input registers:"))
            raw_data["number_of_registers"] = number_of_registers_1
            number_of_registers =  str(number_of_registers_1)    
        except:
            pass
            
                    
                
        save = input("Do you want to save this configuration? Y/N:")
        if (save=="Y" or save=="y" or "yes"):
            with open("{}.json".format(name), 'w') as file:
                json.dump(raw_data, file, indent=2)
                file.close()
                
        
    workbook = xlsxwriter.Workbook('{}.xlsx'.format(name)) 
    worksheet = workbook.add_worksheet() 
    workbook.close()
    while True:
        try:
            client = ModbusClient(method=method, port=port, timeout=timeout, baudrate=baudrate, stopbits=stopbits, parity=parity, bytesize=bytesize)
            client.connect()
            rr = client.read_input_registers(0x00, number_of_registers, unit=1)
            rr = rr.registers
            
        except:
            print("Unable to Connect!")
            return 0
            
        try:    
            xw.Book('{}.xlsx'.format(name)).set_mock_caller()
            my_macro(number_of_registers,rr)
        
            client.close()
        
            time.sleep(1)
            
        except:
            return 0
            
        
            
def my_macro(number_of_registers,rr):
    
    sht = xw.Book.caller().sheets[0]
    
    for i in range (number_of_registers):
        sht.range('A{}'.format(i+1)).value = rr[i]
        
main()
