# Demetrios Doumas 2/26/2020
# A User interface to control electronics of a Hybrid Rocket Engine- Aries Project
# This is an upgraded from the original senior design project in 2018
# This software reads telemetry data as well as send commands to the engine.
# Using the telemetry data, it is displayed as text and visually plotted in real-time.
# The operator has access to the controls during this process.
# The data is saved in excel during the process. PLease make sure to rename the excel file before runnning the software twice.
import sys
from PyQt4 import QtGui, QtCore
from PyQt4.QtCore import *
from PyQt4.QtGui import *
from PyQt4.QtCore import pyqtSlot, QObject
from PyQt4.QtCore import QThread
from PyQt4.QtCore import QMutex
import csv
import serial
import threading
import time
import os
import Queue
import pyqtgraph as pg
import pyqtgraph.exporters
import numpy as np
from PyQt4.QtCore import QTimer
from PyQt4.QtGui import QApplication



def create_excel_file():
  RESULT = ['Mission_time','Temp1','Temp2','Pressure1', 'Pressure2','Temp3','Temp4','Pressure3', 'Pressure4','Temp5','Temp6','Pressure5', 'Pressure6', 'Pressurant_Fill_Indicator','Pressurant_Oxidizer_Indicator','Oxidizer_Fill_Indicator','Oxidizer_Combustion_Indicator']
  with open("output.csv",'wb') as resultFile:
      wr = csv.writer(resultFile, dialect='excel')
      wr.writerow(RESULT)

t1 = threading.Thread(target = create_excel_file)
t1.daemon = True
t1.start()



Lock_One = QMutex() # control ordering threads, Read data then display data
Lock_Two = QMutex() # control ordering threads, Display data then save data

Data_Q=Queue.Queue() # share data object instance
Data_Saved_Q=Queue.Queue() # Shared data being saved to excel file
Ser_Q=Queue.Queue() #  share serial object instance 

Data_plot_Q=Queue.Queue() # Save a instance of the Design class, pass it through a timed function for plotting

global u        # used as alock to only display the legend once
u = False 

# This class creates a serial connection between the PC and microcontroller.
class Com_open:
  def __init__(self, serialport, baud):
    self.baud=baud
    self.serialport=serialport
    print self.serialport
    print self.baud
    try:
      self.serial_object=serial.Serial('COM' + str(serialport),baud,timeout=None)
        
    except ValueError:
      print "baud"
      print "serialport"
      print("Enter Baud and Port")
        

# This class creates the main window GUI, which includes text, plots, and buttons.
class design:

  def __init__(self):

    self.w=QtGui.QWidget()
    self.w.setFixedSize(1640, 800)
    self.w.move(300, 300)
    self.w.setWindowTitle('Aries Project')

    # Arrays for plotting
    # Creating 10 consecutive memories that hold a zero values.
    # The 10 signifies, the number of data points to display on the graphs.
    self.Mission_time=[0]*10
    self.Pressurant_temp=[0]*10
    self.Pressurant_pressure=[0]*10
    self.Oxidizer_temp=[0]*10
    self.Oxidizer_pressure=[0]*10
    self.Combustion_pressure=[0]*10
    
    # placing plots
    pg.setConfigOption('background', 'w')
    self.plot_layout=pg.LayoutWidget(self.w)

    self.my_plot_One=pg.PlotWidget()
    self.my_plot_One.showGrid(x=True,y=True)
    self.my_plot_One.setTitle("Pressurant Tank")
    self.my_plot_One.setFixedSize(533,400)
    self.plot_layout.addWidget(self.my_plot_One)
    self.my_plot_One.setLabel('left', 'Temperature (C) & Pressure (psi)')
    self.my_plot_One.setLabel('bottom', 'Time', units='s')
    self.l=pg.LegendItem((100,60), offset=(70,30))
    self.l.setParentItem(self.my_plot_One.graphicsItem()) 
    
    self.my_plot_Two=pg.PlotWidget()
    self.my_plot_Two.showGrid(x=True,y=True)
    self.my_plot_Two.setTitle("Oxidizer Tank")
    self.my_plot_Two.setFixedSize(533,400)
    self.plot_layout.addWidget(self.my_plot_Two)
    self.my_plot_Two.setLabel('left', 'Temperature (C) & Pressure (psi)')
    self.my_plot_Two.setLabel('bottom', 'Time', units='s')
    self.l_two=pg.LegendItem((100,60), offset=(70,30))
    self.l_two.setParentItem(self.my_plot_Two.graphicsItem()) 
    
    self.my_plot_Three=pg.PlotWidget()
    self.my_plot_Three.showGrid(x=True,y=True)
    self.my_plot_Three.setTitle("Combustion Tank")
    self.my_plot_Three.setFixedSize(533,400)
    self.plot_layout.addWidget(self.my_plot_Three)
    self.my_plot_Three.setLabel('left', 'Pressure (psi)')
    self.my_plot_Three.setLabel('bottom', 'Time', units='s')

    #The lines below create the buttons and QLabels you see on the windoself.
    #Creating Widgets 
    self.QLabelOne= QtGui.QLabel(self.w)#,text= "Tanks:",font = "Times 14 bold")
    self.QLabelOne.setText("Tanks:")
    self.QLabelTwo=QtGui.QLabel(self.w)#(root,text= "Pressurant:",font = "Times 14 bold")
    self.QLabelTwo.setText("Pressurant:")
    self.QLabelThree=QtGui.QLabel(self.w)#(root,text= "Oxidizer:",font = "Times 14 bold")
    self.QLabelThree.setText("Oxidizer:")
    self.QLabelFour=QtGui.QLabel(self.w)#(root,text = "Combustion:",font = "Times 14 bold")
    self.QLabelFour.setText("Combustion:")
    self.QLabelFive=QtGui.QLabel(self.w)#(root,text= "Temperature  (C):")
    self.QLabelFive.setText("Temperature  (C):")
    self.QLabelSix=QtGui.QLabel(self.w)#(root,text= "Temperature  (C):")	
    self.QLabelSix.setText("Temperature  (C):")
    self.QLabelSeven=QtGui.QLabel(self.w)
    self.QLabelSeven.setText("Pressure        (psi):")
    self.QLabelEight=QtGui.QLabel(self.w)
    self.QLabelEight.setText("Pressure        (psi):")
    self.QLabelNine=QtGui.QLabel(self.w)
    self.QLabelNine.setText("Baud Rate:")
    self.QLabelTen=QtGui.QLabel(self.w)
    self.QLabelTen.setText("Port:")	
    self.QLabelTwelve=QtGui.QLabel(self.w)
    self.QLabelTwelve.setText("Valves: ")
    self.QLabelThirteen=QtGui.QLabel(self.w)
    self.QLabelThirteen.setText("High/Low (1/0):")
    self.QLabel_14=QtGui.QLabel(self.w)
    self.QLabel_14.setText("Pressurant_Fill:")
    self.QLabel_16=QtGui.QLabel(self.w)
    self.QLabel_16.setText("Pressurant_Oxidizer:")
    self.QLabel_17=QtGui.QLabel(self.w)
    self.QLabel_17.setText("Oxidizer_fill:")
    self.QLabel_19=QtGui.QLabel(self.w)
    self.QLabel_19.setText("Oxi_Combustion:")

    #Placing Widgets
    self.QLabelOne.move(40, 602)
    self.QLabelTwo.move(175, 602)
    self.QLabelThree.move(275, 602)
    self.QLabelFour.move(370, 602)
    self.QLabelFive.move(40, 640)
    self.QLabelSix.move(40, 660)
    self.QLabelSeven.move(40, 680)
    self.QLabelEight.move(40, 700)
    self.QLabelNine.move(1100, 500)
    self.QLabelTen.move(1100, 540)
    
    self.QLabelTwelve.move(600, 602)
    self.QLabelThirteen.move(800, 602)
    self.QLabel_14.move(600, 640)
    self.QLabel_16.move(600, 660)
    self.QLabel_17.move(600, 680)
    self.QLabel_19.move(600, 700)

    #Data QLabels
    self.QLabel_data1=QtGui.QLabel(self.w)
    self.QLabel_data1.setText("X")
    self.QLabel_data2=QtGui.QLabel(self.w)
    self.QLabel_data2.setText("X")
    self.QLabel_data3=QtGui.QLabel(self.w)
    self.QLabel_data3.setText("X")
    self.QLabel_data4=QtGui.QLabel(self.w)
    self.QLabel_data4.setText("X")
    self.QLabel_data5=QtGui.QLabel(self.w)
    self.QLabel_data5.setText("X")
    self.QLabel_data6=QtGui.QLabel(self.w)
    self.QLabel_data6.setText("X")
    self.QLabel_data7=QtGui.QLabel(self.w)
    self.QLabel_data7.setText("X")
    self.QLabel_data8=QtGui.QLabel(self.w)
    self.QLabel_data8.setText("X")
    self.QLabel_data9=QtGui.QLabel(self.w)
    self.QLabel_data9.setText("X")
    self.QLabel_data10=QtGui.QLabel(self.w)
    self.QLabel_data10.setText("X")
    self.QLabel_data11=QtGui.QLabel(self.w)
    self.QLabel_data11.setText("X")
    self.QLabel_data12=QtGui.QLabel(self.w)
    self.QLabel_data12.setText("X")

    self.QLabel_Pressurant_Fill=QtGui.QLabel(self.w)
    self.QLabel_Pressurant_Fill.setText("X")
    self.QLabel_Pressurant_Oxidizer=QtGui.QLabel(self.w)
    self.QLabel_Pressurant_Oxidizer.setText("X")
    self.QLabel_Oxidizer_fill=QtGui.QLabel(self.w)
    self.QLabel_Oxidizer_fill.setText("X")
    self.QLabel_Oxi_Combustion=QtGui.QLabel(self.w)
    self.QLabel_Oxi_Combustion.setText("X")

    #move Data QLabels
    self.QLabel_data1.move(175,640) #  Temp of Pressurant
    self.QLabel_data2.move(175,660) #  Temp Redundant of Pressurant
    self.QLabel_data3.move(175,680) #  Pressure of Pressurant 
    self.QLabel_data4.move(175,700) #  Pressure Redundant of Pressurant
    self.QLabel_data5.move(275,640) #  Temp of Oxidizer
    self.QLabel_data6.move(275,660) #  Temp Redundant of Oxidizer
    self.QLabel_data7.move(275,680) #  Pressure of Oxidizer
    self.QLabel_data8.move(275,700) #  Pressure Redundant of Oxidizer
    self.QLabel_data9.move(375,640) #  Temp of Combustion Chamber
    self.QLabel_data10.move(375,660)#  Temp Redundant of Combustion Chamber
    self.QLabel_data11.move(375,680)#  Pressure of Combustion Chamber
    self.QLabel_data12.move(375,700)#  Pressure Redundant of Combustion Chamber
        
    self.QLabel_Pressurant_Fill.move(815, 640)
    self.QLabel_Pressurant_Oxidizer.move(815, 660)
    self.QLabel_Oxidizer_fill.move(815, 680)
    self.QLabel_Oxi_Combustion.move(815, 700)
    
    # Entry Box
    self.baud_rate = QtGui.QLineEdit(self.w)
    self.baud_rate.move(1180, 500)
    self.port_entry = QtGui.QLineEdit(self.w)
    self.port_entry.move(1180, 540)
    self.port = str(self.port_entry.text())
    self.baud = str(self.baud_rate.text())

    # Buttons
    self.CONNECT = QtGui.QPushButton("CONNECT", self.w)
    self.CONNECT.move(1020, 500)
    self.DISCONNECT = QtGui.QPushButton("Disconnect", self.w)
    self.DISCONNECT.move(1020, 540)
    self.button1 = QtGui.QPushButton('Open P-Fill', self.w)
    self.button1.move(1020, 600) 
    self.button2 = QtGui.QPushButton('Close P-Fill', self.w)
    self.button2.move(1020, 640)  
    self.button3 = QtGui.QPushButton('Open POIV', self.w)
    self.button3.move(1020, 680)  
    self.button4 = QtGui.QPushButton('Close POIV', self.w)
    self.button4.move(1020, 720) 
    self.button5 = QtGui.QPushButton('Open Oxi-Fill', self.w)
    self.button5.move(1180, 600)
    self.button6 = QtGui.QPushButton('Close Oxi-Fill', self.w)
    self.button6.move(1180, 640)
    self.button7 = QtGui.QPushButton('Launch', self.w)
    self.button7.move(1180, 680) 
    self.button8 = QtGui.QPushButton('Abort', self.w)
    self.button8.move(1180, 720)  

   
    QtCore.QObject.connect(self.CONNECT, QtCore.SIGNAL("clicked()"),self.my_connect)
    QtCore.QObject.connect(self.DISCONNECT, QtCore.SIGNAL("clicked()"), self.Disconnect)
    QtCore.QObject.connect(self.button1, QtCore.SIGNAL("clicked()"),self.open_Pressurant_fill)
    QtCore.QObject.connect(self.button2, QtCore.SIGNAL("clicked()"),self.close_Pressurant_fill)
    QtCore.QObject.connect(self.button3, QtCore.SIGNAL("clicked()"),self.open_POIV)
    QtCore.QObject.connect(self.button4, QtCore.SIGNAL("clicked()"),self.close_POIV)
    QtCore.QObject.connect(self.button5, QtCore.SIGNAL("clicked()"),self.open_Oxidizer_fill)
    QtCore.QObject.connect(self.button6, QtCore.SIGNAL("clicked()"),self.close_Oxidizer_fill)  
    QtCore.QObject.connect(self.button7, QtCore.SIGNAL("clicked()"), self.launch)
    QtCore.QObject.connect(self.button8, QtCore.SIGNAL("clicked()"), self.Abort)


    
  # This function is placed in a thread to record data continuously to a saved excel file
  def Save_Data_Excel_file(self):
    while(1):
      try:
        if(Data_Saved_Q.empty()== False): # empty return true if empty
          Lock_Two.lock()
          self.save_data=Data_Saved_Q.get(block=False)
          Lock_Two.unlock()
          with open("output.csv","a+") as resultFile:   # Save to the excel file
            wr = csv.writer(resultFile, dialect='excel')
            wr.writerow(self.save_data)
      except:
          pass

  # This function creates the serial com communication once the user enters the approprate information
  # and clicks connect. Serial reading will begin in a new thread.
  @pyqtSlot()
  def my_connect(self):
    self.port = str(self.port_entry.text())
    self.baud = str(self.baud_rate.text())
    c=Com_open(self.port, self.baud)
    print c.serial_object
    '''Open a thread'''

    self.myThread=Read_Data(c.serial_object)
    self.myThread.start()
 

  # Sending numerical commands to actuate the engine
  @pyqtSlot()
  def open_Pressurant_fill(self):
    serial_object=Ser_Q.get(block=False)
    serial_object.write(b'1')
    Ser_Q.put(serial_object)
    
  @pyqtSlot()
  def close_Pressurant_fill(self):
    serial_object=Ser_Q.get(block=False)
    serial_object.write(b'2')
    Ser_Q.put(serial_object)
    

  @pyqtSlot()
  def open_POIV(self):
    serial_object=Ser_Q.get(block=False)
    serial_object.write(b'3')
    Ser_Q.put(serial_object)
    

  @pyqtSlot()
  def close_POIV(self):
    serial_object=Ser_Q.get(block=False)
    serial_object.write(b'4')
    Ser_Q.put(serial_object)
      

  @pyqtSlot() 
  def open_Oxidizer_fill(self):
    serial_object=Ser_Q.get(block=False)
    serial_object.write(b'5')
    Ser_Q.put(serial_object)
    

  @pyqtSlot()
  def close_Oxidizer_fill(self):
    serial_object=Ser_Q.get(block=False)
    serial_object.write(b'6')
    Ser_Q.put(serial_object)
      

  @pyqtSlot()
  def launch(self):
    serial_object=Ser_Q.get(block=False)
    serial_object.write(b'7')
    Ser_Q.put(serial_object)

  @pyqtSlot()
  def Abort(self):
    serial_object=Ser_Q.get(block=False)
    serial_object.write(b'8')
    Ser_Q.put(serial_object)



  @pyqtSlot()
  def Disconnect(self):
    serial_object=Ser_Q.get(block=False)
    try:
        serial_object.close()  # Close serial port connection
        
    except AttributeError:
        print("Closed without opening it")
    Ser_Q.put(serial_object)
  
  

    
    
# This class is responsible to read data place it in a queue, where another thread reads the data and displays it to the screen.
class Read_Data(QThread):

  def __init__(self, Ser_obj):
    self.Ser_obj=Ser_obj
    Ser_Q.put(Ser_obj)
    QThread.__init__(self)

  def __del__(self):
    #self.wait()
    self.exit()
  
  def run(self):
  
    while(1):
      
      while(self.Ser_obj.inWaiting()==0):
        pass
      try:
        self.serial_data=self.Ser_obj.readline().strip('\n').strip('\r')
        print self.serial_data
        self.filter_data= self.serial_data.split(',')
        Lock_One.lock()
        Data_Q.put(self.filter_data)
        Lock_One.unlock()
      except:
        pass


# This function is called every 1 sec to update the info on the screen. Info such as text data, and plot data.
def update():
  global u
  
  temp=Data_plot_Q.get(block=False)
  Data_plot_Q.put(temp)

  if(Data_Q.empty()== False): # empty return true if empty
      Lock_One.lock()
      temp.fill_data=Data_Q.get(block=False)
      Lock_One.unlock()
      
      if (temp.fill_data):
        # Display variables on the screen in text
        # Data Label 1                               temp1    
        temp.QLabel_data1.setFixedWidth(80)             
        temp.QLabel_data1.setText(temp.fill_data[1])


        # Data Label 2                               temp2
        temp.QLabel_data2.setFixedWidth(80) 
        temp.QLabel_data2.setText(temp.fill_data[2])

        # Data Label 3                            pressure1

        temp.QLabel_data3.setFixedWidth(80) 
        temp.QLabel_data3.setText(temp.fill_data[3])

        # Data Label 4                          pressure2                            
        temp.QLabel_data4.setFixedWidth(80) 
        temp.QLabel_data4.setText(temp.fill_data[4])

        #Column 2 from the left
        ######################################

        # Data Label 5                          temp3
        temp.QLabel_data5.setFixedWidth(80) 
        temp.QLabel_data5.setText(temp.fill_data[5])

        # Data Label 6                          temp4
        temp.QLabel_data6.setFixedWidth(80) 
        temp.QLabel_data6.setText(temp.fill_data[6])

        # Data Label 7                          pressure3
        temp.QLabel_data7.setFixedWidth(80) 
        temp.QLabel_data7.setText(temp.fill_data[7])

        # Data Label 8                         pressure4
        temp.QLabel_data8.setFixedWidth(80) 
        temp.QLabel_data8.setText(temp.fill_data[8])

        #Column 3 from the left
        #######################################
        # Data Label 9                          temp5
        temp.QLabel_data9.setFixedWidth(80) 
        temp.QLabel_data9.setText(temp.fill_data[9])

        # Data Label 10                         temp6
        temp.QLabel_data10.setFixedWidth(80) 
        temp.QLabel_data10.setText(temp.fill_data[10])

        # Data Label 11                         pressure5
        temp.QLabel_data11.setFixedWidth(80) 
        temp.QLabel_data11.setText(temp.fill_data[11])

        # Data Label 12                         pressure6
        temp.QLabel_data12.setFixedWidth(80) 
        temp.QLabel_data12.setText(temp.fill_data[12])

        #Column 4 Valves   Indicators
        ######################################

        # Data Pressurant_Fill Valve Indicator           Pressurant_Fill_Indicator
        temp.QLabel_Pressurant_Fill.setFixedWidth(80) 
        temp.QLabel_Pressurant_Fill.setText(temp.fill_data[13])

        # Data Pressurant_Oxidizer Valve Indicator       Pressurant_Oxidizer_Indicator
        temp.QLabel_Pressurant_Oxidizer.setFixedWidth(80)
        temp.QLabel_Pressurant_Oxidizer.setText(temp.fill_data[14])

        # Data Oxidizer_fill Valve Indicator             Oxidizer_Fill_Indicator
        temp.QLabel_Oxidizer_fill.setFixedWidth(80)
        temp.QLabel_Oxidizer_fill.setText(temp.fill_data[15])

        # Data Oxi_Combustion Valve Indicator            Oxidizer_Combustion_Indicator
        temp.QLabel_Oxi_Combustion.setFixedWidth(80)
        temp.QLabel_Oxi_Combustion.setText(temp.fill_data[16])

        

        temp.QLabel_data1.setAutoFillBackground(True)
        temp.QLabel_data3.setAutoFillBackground(True)
        temp.QLabel_data5.setAutoFillBackground(True)
        temp.QLabel_data7.setAutoFillBackground(True)
      

       
        PressurantTemp=(float(temp.fill_data[1]))
        PressurantPressure=(float(temp.fill_data[3]))
        OxidizerTemp=(float(temp.fill_data[5]))
        OxidizerPressure=(float(temp.fill_data[7]))
        
        # Warning colors
        if(PressurantPressure >= 50.0):  
          temp.QLabel_data3.setStyleSheet('color: white; background-color: red')
        elif ((PressurantPressure >=30.0 and PressurantPressure <= 49.99) ):
          temp.QLabel_data3.setStyleSheet('color: black; background-color: yellow')
        else:
          temp.QLabel_data3.setStyleSheet('color: black; background-color: None')
          
        
        if(PressurantTemp >= 40.0):
          temp.QLabel_data1.setStyleSheet('color: white; background-color: red')
        elif((PressurantTemp >=30) and (PressurantTemp < 40) ):
          temp.QLabel_data1.setStyleSheet('color: black; background-color: yellow')
        else:
          temp.QLabel_data1.setStyleSheet('color: black; background-color: None')     

          
        if (OxidizerPressure >= 50.0):
          temp.QLabel_data7.setStyleSheet('color: white; background-color: red')
        elif ((OxidizerPressure >=30) and (OxidizerPressure < 50)):
          temp.QLabel_data7.setStyleSheet('color: black; background-color: yellow')
        else:
          temp.QLabel_data7.setStyleSheet('color: black; background-color: None')

      
        if (OxidizerTemp >= 40.0):
          temp.QLabel_data5.setStyleSheet('color: white; background-color: red')
        elif ((OxidizerTemp >=30) and (OxidizerTemp < 40) ):
          temp.QLabel_data5.setStyleSheet('color: black; background-color: yellow')
        else:
          temp.QLabel_data5.setStyleSheet('color: black; background-color: None')

        
        # Push the data to the arrays for plotting
        temp.Mission_time.append(float(temp.fill_data[0]))
        temp.Pressurant_temp.append(float(temp.fill_data[1]))
        temp.Pressurant_pressure.append(float(temp.fill_data[3]))
        temp.Oxidizer_temp.append(float(temp.fill_data[5]))
        temp.Oxidizer_pressure.append(float(temp.fill_data[7]))
        temp.Combustion_pressure.append(float(temp.fill_data[11]))
  
        temp.my_plot_One.clear()
        temp.my_plot_Two.clear()
        temp.my_plot_Three.clear() 

        C1=temp.my_plot_One.plot(temp.Mission_time, temp.Pressurant_temp, pen=pg.mkPen('r'), style=QtCore.Qt.DashLine)
        C2=temp.my_plot_One.plot(temp.Mission_time,temp.Pressurant_pressure,pen=pg.mkPen('k'), style=QtCore.Qt.DashLine)
        C3=temp.my_plot_Two.plot(temp.Mission_time, temp.Oxidizer_temp, pen=pg.mkPen('b'), style=QtCore.Qt.DashLine)
        C4=temp.my_plot_Two.plot(temp.Mission_time, temp.Oxidizer_pressure, pen=pg.mkPen('m'), style=QtCore.Qt.DashLine)
        

        if(u== False):
            temp.l.addItem(C1,'Temperature (C)')
            temp.l.addItem(C2,'Pressure (Psi)')
            temp.l_two.addItem(C3,'Temperature (C)')
            temp.l_two.addItem(C4,'Pressure (Psi)')
            u=True
        
        temp.my_plot_Three.plot(temp.Mission_time, temp.Combustion_pressure, pen=pg.mkPen('g'), style=QtCore.Qt.DashLine) 
        QApplication.processEvents()

        temp.Mission_time.pop(0)
        temp.Pressurant_temp.pop(0)
        temp.Pressurant_pressure.pop(0)
        temp.Oxidizer_temp.pop(0)
        temp.Oxidizer_pressure.pop(0)
        temp.Combustion_pressure.pop(0)
        
        Lock_Two.lock()
        Data_Saved_Q.put(temp.fill_data)
        Lock_Two.unlock()
 


def main():
  
  app = QtGui.QApplication(sys.argv) 
  
  x=design()
  
  Data_plot_Q.put(x)

 
  
  timer = QTimer()
  timer.timeout.connect(update)
  timer.start(1000)  #5000 #3000 #1500

  
  
  
  t3 = threading.Thread(target = x.Save_Data_Excel_file)
  t3.daemon = True
  t3.start()
  
  x.w.show()
  
  
  
  app.exec_()

if __name__ == '__main__':
    main()