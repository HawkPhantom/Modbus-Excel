import os
import xlwings as xw
import sys
from pymodbus.client.sync import ModbusSerialClient as ModbusClient
import time
import json
import xlsxwriter 


def main():
    
    print("Written by Furkan Ciylan \nContact me: furkanciylan@gmail.com\nGithub: /HawkPhantom")
    
    name = input("\nName this session:")
    a = input("\nDo you want to configure the software? \n Default Values Are: \n method:rtu \n port:COM7 \n timeout:10ds \n baudrate:19200 \n stopbits:1 \n parity:Even \n bytesize:8 \n number of input registers:8 \n number of holding registers:8 \n number of coils:8 \n address:0x00 \nY/N:")
    
    if (a == "n" or a == "N" or a == "no"):
        conffile_test = input("\nDo you have a custom configuration file? Y/N:")
    else:
        conffile_test = ""
    
    return(modbus(a,conffile_test,name))   
        
def modbus(a,conffile_test,name):
    
    
    if (conffile_test == "Y" or conffile_test == "y" or conffile_test =="yes"):
        conffile = input("Please write the path of configuration file:")
    else:
        conffile = ("configuration.json")
        
    read_selector = input("\nChoose what to read\n1)Holding Registers\n2)Input Registers\n3)Read Coils(Experimental)\nPlease press 1-2-3:")
    
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
            number_of_input_registers = int(raw_data["number_of_input_registers"])
            
            number_of_holding_registers = int(raw_data["number_of_holding_registers"])
            
            number_of_coils = int(raw_data["number_of_coils"])
            
            address = raw_data["address"]
            
            database.close()
            
    except:
        print("Configuration File Cannot be Found")
        try_again = input("Do you want to specify another location? Y/N:")
        if(try_again == "y" or try_again == "Y" or try_again == "yes"):
            modbus(a,conffile_test,name)
        else:
            return 0
        
    if (a =="y" or a=="Y" or a=="yes"):
        
        print("\nConfiguration process is starting\nNote that:If you left it empty it will stay as default")
            
        method_1 = input("\nPlease enter the method (ascii or rtu):")
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
            raw_data["stopbits"] = str(stopbits_1)
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
            bytesize = str(bytesize_1)
        except:
            pass
            
        if(read_selector == "1"):
            try:
                number_of_holding_registers_1 = int(input("Please enter the number of holding registers to read:"))
                raw_data["number_of_holding_registers"] = str(number_of_holding_registers_1)
                number_of_holding_registers =  number_of_holding_registers_1
            except:
                pass
            
        elif(read_selector == "2"):
            try:
                number_of_input_registers_1 = int(input("Please enter the number of input registers to read:"))
                raw_data["number_of_input_registers"] = str(number_of_input_registers_1)
                number_of_input_registers =  number_of_input_registers_1
            except:
                pass
            
        elif(read_selector == "3"):
            try:
                number_of_coils_1 = int(input("Please enter the number of coils to read:"))
                raw_data["number_of_coils"] = str(number_of_coils_1)
                number_of_coils =  number_of_coils_1
            except:
                pass
            
            
        address_1 = input("Please specify the starting address:")
        if (address_1 != ""):
            address = address_1
            raw_data["address"] = address_1
            
        save = input("\nDo you want to save this configuration? Y/N:")
        if (save=="Y" or save=="y" or "yes"):
            with open("{}.json".format(name), 'w') as file:
                json.dump(raw_data, file, indent=2)
                file.close()
                
    print("\nCreating the Excel Document")
    workbook = xlsxwriter.Workbook('{}.xlsx'.format(name)) 
    worksheet = workbook.add_worksheet() 
    workbook.close()
    
    address=(int(address,16))
    
    if(read_selector == "1"):
        holding_registers(method,port,timeout,baudrate,stopbits,parity,bytesize,number_of_holding_registers,name,address)
        
    elif(read_selector == "2"):
        input_registers(method,port,timeout,baudrate,stopbits,parity,bytesize,number_of_input_registers,name,address)
   
    elif(read_selector == "3"):
        coils(method,port,timeout,baudrate,stopbits,parity,bytesize,number_of_coils,name,address) 
    else:
        print("\nWrong Function Selection")       
        
            
def my_macro(counter,rr):
    
    sht = xw.Book.caller().sheets[0]
    
    for i in range (counter):
        sht.range('A{}'.format(i+1)).value = rr[i]
        
        
        
def holding_registers(method,port,timeout,baudrate,stopbits,parity,bytesize,number_of_holding_registers,name,address):
    i = 0
    while True:
        
        try:
            if(i==0):
                print("\nTrying to Connect")
            client = ModbusClient(method=method, port=port, timeout=timeout, baudrate=baudrate, stopbits=stopbits, parity=parity, bytesize=bytesize)
            client.connect()
            rr = client.read_holding_registers(address, number_of_holding_registers, unit=1)
            if(i==0):
                print("\nConnection Successful")
                print(rr)
            rr = rr.registers
            
        except:
            print("\nUnable to Connect!")
            return 0
            
        try:    
            if(i==0):
                print("\nWriting to Excel")
            xw.Book('{}.xlsx'.format(name)).set_mock_caller()
            my_macro(number_of_holding_registers,rr)
        
            client.close()
        
            time.sleep(1)
            
        except:
            client.close()
            return 0   
         
        i = i + 1
        if(i==1 or i==2):
            i=1
    
        
        
def input_registers(method,port,timeout,baudrate,stopbits,parity,bytesize,number_of_input_registers,name,address):
    i = 0
    while True:
        try:
            if(i==0):
                print("\nTrying to Connect")
            client = ModbusClient(method=method, port=port, timeout=timeout, baudrate=baudrate, stopbits=stopbits, parity=parity, bytesize=bytesize)
            client.connect()
            rr = client.read_input_registers(address, number_of_input_registers, unit=1)
            if(i==0):
                print("\nConnection Successful")
                print(rr)
            rr = rr.registers
            
        except:
            print("\nUnable to Connect!")
            return 0
            
        try:    
            if(i==0):
                print("\nWriting to Excel")
            xw.Book('{}.xlsx'.format(name)).set_mock_caller()
            my_macro(number_of_input_registers,rr)
        
            client.close()
        
            time.sleep(1)
            
        except:
            client.close()
            return 0
        i = i + 1
        if(i==1 or i==2):
            i=1
        
        
        
def coils(method,port,timeout,baudrate,stopbits,parity,bytesize,number_of_coils,name,address):
    i=0
    while True:
        try:
            if(i==0):
                print("\nTrying to Connect")
            client = ModbusClient(method=method, port=port, timeout=timeout, baudrate=baudrate, stopbits=stopbits, parity=parity, bytesize=bytesize)
            client.connect()
            rr = client.read_coils(address, number_of_coils, unit=1)
            if(i==0):
                print("\nConnection Successful")
                print(rr)
            rr=rr.registers()
            
        except:
            print("\nUnable to Connect!")
            return 0
            
        try:    
            if(i==0):
                print("\nWriting to Excel")
            xw.Book('{}.xlsx'.format(name)).set_mock_caller()
            my_macro(number_of_coils,rr)
        
            client.close()
        
            time.sleep(1)
            
        except:
            client.close()
            return 0    
        
        i = i + 1
        if(i==1 or i==2):
            i=1

        

    
main()
