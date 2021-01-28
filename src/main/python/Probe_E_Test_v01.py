from PyQt5 import QtWidgets, uic, QtGui, QtCore
from PyQt5.QtWidgets import QMainWindow, QFileDialog
from PyQt5.QtCore import QCoreApplication
from pyqtgraph import PlotWidget, plot
import numpy as np
import pyqtgraph as pg
import sys  # We need sys so that we can pass argv to QApplication
import os
import cap_backend
import pandas as pd
from openpyxl import load_workbook
from os import path
import xlrd


class SAVEWINDOW(QtWidgets.QMainWindow):                           # <===
    def __init__(self):
        super(SAVEWINDOW, self).__init__()
        uic.loadUi('SaveWindow.ui', self) 
        self.setWindowTitle("Save Window")

class Cap_Test(QtWidgets.QMainWindow):

    def __init__(self, ui, connections, config, parent = None):
        super().__init__(parent)     
        self.colors = [(255,0,0),(255,187,51),(25,64,255),(179,0,149),(26, 196, 20)]
        self.Instrument_ID = [0]*2
        self.ImpArray_arg = [0]*7
        self.NUM_COLORS = {
            1: '#a6a5b0',  #grey 
            2: '#f2050c',  #red
            3: '#0cc92c'   #green
        }
        
        #Load the UI Page
        uic.loadUi(ui, self)
        Connections = pd.read_csv(connections)
        print(Connections)
        self.Instrument_ID[0] = Connections.iloc[0,1] #Hioki
        self.Instrument_ID[1] = Connections.iloc[1,1] #Switching Matrix 
        self.save_data_folder = Connections.iloc[2,1]#Data

        Config = pd.read_csv(config)
        print(Config)
        self.CapLo = Config.iloc[0,1]
        self.CapHi = Config.iloc[1,1]
        self.LossLo = Config.iloc[2,1]
        self.LossHi = Config.iloc[3,1]
        self.default_voltage = Config.iloc[4,1]
        self.default_frequency = Config.iloc[5,1]
        self.default_num_meas = Config.iloc[6,1]

        self.graphWidget_CapHistogram.setBackground('w')
        self.graphWidget_LossHistogram.setBackground('w')

        ##########################################

        self.tableWidget_CapArray.setRowCount(8)
        self.tableWidget_CapArray.setColumnCount(1)   
        self.tableWidget_CapArray.horizontalHeader().setResizeMode(QtWidgets.QHeaderView.Stretch)  
        self.tableWidget_CapArray.verticalHeader().setResizeMode(QtWidgets.QHeaderView.Stretch)  

        for i in range(8):
            for j in range(1):
                self.tableWidget_CapArray.setItem(i,j, QtWidgets.QTableWidgetItem("N/A"))
                self.tableWidget_CapArray.item(i, j).setBackground(QtGui.QColor(self.NUM_COLORS[1]))
                self.tableWidget_CapArray.item(i, j).setTextAlignment(QtCore.Qt.AlignHCenter)
                font = QtGui.QFont()
                font.setPointSize(14)
                font.setBold(True)
                self.tableWidget_CapArray.item(i, j).setFont(font)

        self.tableWidget_LossArray.setRowCount(8)
        self.tableWidget_LossArray.setColumnCount(1)   
        self.tableWidget_LossArray.horizontalHeader().setResizeMode(QtWidgets.QHeaderView.Stretch)  
        self.tableWidget_LossArray.verticalHeader().setResizeMode(QtWidgets.QHeaderView.Stretch)  

        for i in range(8):
            for j in range(1):
                self.tableWidget_LossArray.setItem(i,j, QtWidgets.QTableWidgetItem("N/A"))
                self.tableWidget_LossArray.item(i, j).setBackground(QtGui.QColor(self.NUM_COLORS[1]))
                self.tableWidget_LossArray.item(i, j).setTextAlignment(QtCore.Qt.AlignHCenter)
                font = QtGui.QFont()
                font.setPointSize(14)
                font.setBold(True)
                self.tableWidget_LossArray.item(i, j).setFont(font)
        
        self.state.setText('Not Saved') #save indicator
        self.state.setStyleSheet("background-color: yellow;  border: 1px solid black;")         
         
        self.Run.clicked.connect(self.measureImpArray)
        self.save_ImpArray.clicked.connect(self.name_construction)


    def check_valuesCap(self,Cap):
        global colorCap
        global textCap

        if Cap>self.CapHi:
            textCap = 'Hi'
            colorCap = self.NUM_COLORS[2]

        if Cap<self.CapLo:
            textCap = 'Lo'
            colorCap = self.NUM_COLORS[2]

        if ((Cap>self.CapLo) and (Cap<self.CapHi)):
            textCap = 'In'
            colorCap = self.NUM_COLORS[3]

        return (textCap,colorCap)


    def check_valuesLoss(self,Loss):
        global colorLoss
        global textLoss

        if Loss>self.LossHi:
            textLoss = 'Hi'
            colorLoss = self.NUM_COLORS[2]

        if Loss<self.LossLo:
            textLoss = 'Lo'
            colorLoss = self.NUM_COLORS[2]

        if ((Loss>self.LossLo) and (Loss<self.LossHi)):
            textLoss = 'In'
            colorLoss = self.NUM_COLORS[3]

        return (textLoss, colorLoss)

    
    def measureImpArray(self): 
        
        self.state.setText('Not Saved') #save indicator
        self.state.setStyleSheet("background-color: yellow;  border: 1px solid black;")
        
        self.Cap_Array = []
        self.Loss_Array = []
        self.graphWidget_CapHistogram.clear() 
        self.graphWidget_LossHistogram.clear() 
        cap_backend.reset_matrix(self.Instrument_ID)

        num_X = int(self.NumProbeLocs.text())
        num_Y = 1

        self.tableWidget_CapArray.setRowCount(num_X)
        self.tableWidget_CapArray.setColumnCount(num_Y)   
        self.tableWidget_CapArray.horizontalHeader().setResizeMode(QtWidgets.QHeaderView.Stretch)  
        self.tableWidget_CapArray.verticalHeader().setResizeMode(QtWidgets.QHeaderView.Stretch)

        self.tableWidget_LossArray.setRowCount(num_X)
        self.tableWidget_LossArray.setColumnCount(num_Y)   
        self.tableWidget_LossArray.horizontalHeader().setResizeMode(QtWidgets.QHeaderView.Stretch)  
        self.tableWidget_LossArray.verticalHeader().setResizeMode(QtWidgets.QHeaderView.Stretch)


        #initialize display matrices
        for i in range(num_X):
            for j in range(num_Y):
                self.tableWidget_CapArray.setItem(i,j, QtWidgets.QTableWidgetItem("N/A"))
                self.tableWidget_CapArray.item(i, j).setBackground(QtGui.QColor('#a6a5b0'))
                self.tableWidget_CapArray.item(i, j).setTextAlignment(QtCore.Qt.AlignHCenter)
                font = QtGui.QFont()
                font.setPointSize(14)
                font.setBold(True)
                self.tableWidget_CapArray.item(i, j).setFont(font)

        for i in range(num_X):
            for j in range(num_Y):
                self.tableWidget_LossArray.setItem(i,j, QtWidgets.QTableWidgetItem("N/A"))
                self.tableWidget_LossArray.item(i, j).setBackground(QtGui.QColor('#a6a5b0'))
                self.tableWidget_LossArray.item(i, j).setTextAlignment(QtCore.Qt.AlignHCenter)
                font = QtGui.QFont()
                font.setPointSize(14)
                font.setBold(True)
                self.tableWidget_LossArray.item(i, j).setFont(font)

        for i in range(num_X):
                self.ImpArray_arg[0] = self.default_voltage
                self.ImpArray_arg[1] = self.default_frequency
                self.ImpArray_arg[2] = self.default_num_meas
                self.ImpArray_arg[3] = i+1 

                (Cap,Loss) = cap_backend.main(self.ImpArray_arg, self.Instrument_ID)

                self.Cap_Array.append(Cap)
                self.Loss_Array.append(Loss)
                (textCap, colorCap) = self.check_valuesCap(Cap)
                (textLoss, colorLoss) = self.check_valuesLoss(Loss)

                self.tableWidget_CapArray.setItem(i,0, QtWidgets.QTableWidgetItem('Cap = %0.2f' %(Cap) +'  '+'('+textCap+')'))
                self.tableWidget_CapArray.item(i, 0).setBackground(QtGui.QColor(colorCap))
                self.tableWidget_CapArray.item(i,0).setTextAlignment(QtCore.Qt.AlignHCenter)
                font = QtGui.QFont()
                font.setPointSize(14)
                font.setBold(True)
                self.tableWidget_CapArray.item(i,0).setFont(font)

                self.tableWidget_LossArray.setItem(i,0, QtWidgets.QTableWidgetItem('Loss = %0.2f' %(Loss) +'  '+'('+textLoss+')'))
                self.tableWidget_LossArray.item(i, 0).setBackground(QtGui.QColor(colorLoss))
                self.tableWidget_LossArray.item(i,0).setTextAlignment(QtCore.Qt.AlignHCenter)
                font = QtGui.QFont()
                font.setPointSize(14)
                font.setBold(True)
                self.tableWidget_LossArray.item(i,0).setFont(font)


        y,x = np.histogram(self.Cap_Array, bins=np.linspace( np.min(self.Cap_Array), np.max(self.Cap_Array), 20))
        self.graphWidget_CapHistogram.plot(x, y, stepMode=True, fillLevel=0, brush= self.NUM_COLORS[1])
        self.max_Cap.setText(str(round(np.max(self.Cap_Array),2)  ))
        self.min_Cap.setText(str(round(np.min(self.Cap_Array),2)  ))
        self.avg_Cap.setText(str(round(np.average(self.Cap_Array),2)  ))
        self.stdev_Cap.setText(str(round(np.std(self.Cap_Array),2)  ))
        self.Cap_LSL.setText(str(round((self.CapLo),2)  ))
        self.Cap_USL.setText(str(round((self.CapHi),2)  ))

        y,x = np.histogram(self.Loss_Array, bins=np.linspace( np.min(self.Loss_Array), np.max(self.Loss_Array), 20))
        self.graphWidget_LossHistogram.plot(x, y, stepMode=True, fillLevel=0, brush= self.NUM_COLORS[1])
        self.max_Loss.setText(str(round(np.max(self.Loss_Array),2)  ))
        self.min_Loss.setText(str(round(np.min(self.Loss_Array),2)  ))
        self.avg_Loss.setText(str(round(np.average(self.Loss_Array),2)  ))
        self.stdev_Loss.setText(str(round(np.std(self.Loss_Array),2)  ))
        self.Loss_LSL.setText(str(round((self.LossLo),2)  ))
        self.Loss_USL.setText(str(round((self.LossHi),2)  ))


        self.ImpArray_Detailed = np.zeros(shape=(len(self.Cap_Array),2))
        self.ImpArray_Detailed[:,0] = self.Cap_Array
        self.ImpArray_Detailed[:,1] = self.Loss_Array
            
        self.ImpArray_Summary = np.zeros(shape=(1,8))
        self.ImpArray_Summary[:,0] = round(np.max(self.Cap_Array),2)
        self.ImpArray_Summary[:,1] = round(np.min(self.Cap_Array),2)
        self.ImpArray_Summary[:,2] = round(np.average(self.Cap_Array),2)
        self.ImpArray_Summary[:,3] = round(np.std(self.Cap_Array),2)
        
        self.ImpArray_Summary[:,4] = round(np.max(self.Loss_Array),2)
        self.ImpArray_Summary[:,5] = round(np.min(self.Loss_Array),2)
        self.ImpArray_Summary[:,6] = round(np.average(self.Loss_Array),2)
        self.ImpArray_Summary[:,7] = round(np.std(self.Loss_Array),2)

    def name_construction(self):
        
        self.w = SAVEWINDOW()
        self.w.show()
        self.w.ID.setText(self.SerialNum.text())
        
        self.w.save_b.clicked.connect(self.export_1)
        
    def export_1(self):
        
        saveName = '\Capacitance_' + self.w.ID.text() + '.xlsx'
        tag = 'Other'
        
        filePath = r'C:\Users\Nicolas Madhavapeddy\OneDrive - Frore Systems\Desktop\Capacitance\data' + r''
        
        if self.w.bottom_stack.isChecked():
            tag = 'Bottom Stack'
        if self.w.top_plate.isChecked():
            tag = 'Top Plate'
        if self.w.heat_spreader.isChecked():
            tag = 'Heat Spreader'
        if self.w.after_tack.isChecked():
            tag = 'After Tack'
        if self.w.after_bond.isChecked():
            tag = 'After Bond'
        if self.w.after_cure.isChecked():
            tag = 'After Cure'
            
        
        while True:
            try:
                dir_path = QFileDialog.getExistingDirectory(self, 'open a folder', r'C:\Users\Nicolas Madhavapeddy\OneDrive - Frore Systems\Desktop\Capacitance\data')
                break
            except PermissionError:
                return None
        if path.exists(dir_path + saveName): #testing if sheet already exists
            xls = xlrd.open_workbook(dir_path + saveName, on_demand=True)
            if tag in xls.sheet_names():
                self.state.setText('Warning: File already exists')
                self.state.setStyleSheet("background-color: red;  border: 1px solid black;") 
                return None
            
          
        print('Saving: ' + str(saveName) + ' ' + str(tag) + ' Capacitance Test')
        
        writer = pd.ExcelWriter(dir_path + saveName, engine='openpyxl')

        try:
            writer.book = load_workbook(dir_path + saveName)
        except:
            print('Writting new file')
        
        summary = pd.DataFrame({saveName : ['Max Capacitance', 'Min Capacitance', 'Average Capacitance', 'STD Capacitance',
                                           'Max Loss', 'Min Loss', 'Average Loss', 'STD Loss'],
                                'Statistics': self.ImpArray_Summary[0] })
        
        summary.to_excel(writer, sheet_name=tag, index=False, startcol=0, startrow=0)
        
        detailed = pd.DataFrame({'Side' : ['1L', '1R', '2L', '2R', '3L', '3R', '4L', '4R'],
                                'Capacitance': [sub[0] for sub in self.ImpArray_Detailed],
                                'Loss': [sub[1] for sub in self.ImpArray_Detailed] })
        
        detailed.to_excel(writer, sheet_name=tag, index=False, startcol=2, startrow=0)
        
        writer.save()
        print('**************************')
        print('Finished Saving')
        self.state.setText('Saved File')
        self.state.setStyleSheet("background-color: green;  border: 1px solid black;")
        
        self.w.hide()
             
