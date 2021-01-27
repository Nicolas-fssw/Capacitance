import numpy as np
import time
import os
import tqdm as _tqdm
import visa 
import serial 

#Hioki backend 
def get_impedance_data(inst,voltage,frequency,num):
    inst.write(':HEAD OFF')
    inst.write(':AVER 256')
    inst.write(':MEAS:ITEM 45,64')
    inst.write(':LEV:V')
    inst.write(':LEV:VOLT %0.2f'%(voltage)) #You may want to change the voltage level 
    inst.write(':FREQ %0.2f'%(frequency)) #you may want to change this 
    Z = []
    Ph = []
    C = []
    D = []
    Rdc = []

    for i in range(num):
        pars = inst.query(':MEAS?')
        pars = pars.split(',')
        Z.append(float(pars[0]))
        Ph.append(float(pars[1]))
        C.append(float(pars[2]))
        D.append(float(pars[3]))
        Rdc.append(float(pars[4]))
        time.sleep(0.1)

    #Z_av = np.average(Z)
    #Ph_av = np.average(Ph)
    C_av = np.average(C)*1e9
    D_av = np.average(D)*100
    #Rdc_av = np.average(Rdc)
    print(C_av)

    return (C_av,D_av)

 
#switching matrix backend 
def reset(inst):
    inst.write('*CLS')
    inst.write('*RST')
    
def make_contact(inst,row,col):
    inst.write('ROUTe:CLOSe (@%d0%d)'%(row,col))
    
def break_contact(inst,row,col):
    inst.write('ROUTe:OPEN (@%d0%d)'%(row,col))

def reset_matrix(inst_ID):
    rm = visa.ResourceManager()
    switch_matrix = rm.open_resource(inst_ID[1])
    reset(switch_matrix)

def main(arg, inst_ID):
    rm = visa.ResourceManager()
    hioki =  rm.open_resource(inst_ID[0]) 
    switch_matrix = rm.open_resource(inst_ID[1])
    voltage = int(arg[0])
    frequency = int(arg[1])
    num = int(arg[2])
    switch_number = int(arg[3])
    make_contact(switch_matrix,1,switch_number)
    time.sleep(0.5)
    (Cap, Loss) = get_impedance_data(hioki,voltage,frequency,num)
    break_contact(switch_matrix,1,switch_number)
    return (Cap, Loss)
