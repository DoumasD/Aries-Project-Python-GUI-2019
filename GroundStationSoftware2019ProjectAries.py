import sys
from PyQt4 import QtGui, QtCore
from PyQt4.QtCore import *
from PyQt4.QtGui import *
from PyQt4.QtCore import pyqtSlot, QObject
import csv
import serial
import threading
import time
import os


def create_excel_file():
  RESULT = ['Mission_time','Temp1','Temp2','Pressure1', 'Pressure2','Temp3','Temp4','Pressure3', 'Pressure4','Temp5','Temp6','Pressure5', 'Pressure6', 'Pressurant_Fill_Indicator','Pressurant_Oxidizer_Indicator','Oxidizer_Fill_Indicator','Oxidizer_Combustion_Indicator']
  with open("output.csv",'wb') as resultFile:
      wr = csv.writer(resultFile, dialect='excel')
      wr.writerow(RESULT)

t1 = threading.Thread(target = create_excel_file)
t1.daemon = True
t1.start()

global x
    
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
        
    
class design:


  def __init__(self):

    self.w=QtGui.QWidget()
    self.w.setFixedSize(1600, 800)
    self.w.move(300, 300)
    self.w.setWindowTitle('Aries Project')

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

    #print self.port_entry.text()
    #print self.baud_rate.text()
    self.port = str(self.port_entry.text())
  
    self.baud = str(self.baud_rate.text())
  
    #self.com_info=[]

    #self.com_info.append(self.port)
    #self.com_info.append(self.baud)
    #print self.com_info

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


        # Functionsserial_objectserial_objectseserial_object
  @pyqtSlot()
  def my_connect(self):
    self.port = str(self.port_entry.text())
    self.baud = str(self.baud_rate.text())
    Com_open(self.port, self.baud)

  # Sending numerical commands to actuate the engine
  @pyqtSlot()
  def open_Pressurant_fill(self):
    Com_open(self.port, self.baud).serial_object.write(b'1')
    
  @pyqtSlot()
  def close_Pressurant_fill(self):
    Com_open(self.port, self.baud).serial_object.write(b'2')
    

  @pyqtSlot()
  def open_POIV(self):
    Com_open(self.port, self.baud).serial_object.write(b'3')
    

  @pyqtSlot()
  def close_POIV(self):
    Com_open(self.port, self.baud).serial_object.write(b'4')
      

  @pyqtSlot() 
  def open_Oxidizer_fill(self):
    Com_open(self.port, self.baud).serial_object.write(b'5')
    

  @pyqtSlot()
  def close_Oxidizer_fill(self):
    Com_open(self.port, self.baud).serial_object.write(b'6')
      

  @pyqtSlot()
  def launch(self):
    Com_open(self.port, self.baud).serial_object.write(b'7')

  @pyqtSlot()
  def Abort(self):
    Com_open(self.port, self.baud).serial_object.write(b'8')



  @pyqtSlot()
  def Disconnect(self):
    try:
        Com_open(self.port, self.baud).serial_object.close()  # Close serial port connection
        
    except AttributeError:
        print("Closed without opening it")




  


    

    
    


  
'''
#@pyqtSlot()
def get_data():
  """  Update all the labels based on telemetry string
      Mission_time, temp1, temp2, pressure1, pressure2, temp3, temp4, pressure3, pressure4, temp5,
      temp6, pressure5, pressure6, Pressurant_Fill_Indicator, Pressurant_Oxidizer_Indicator,
      Oxidizer_Fill_Indicator, and Oxidizer_Combustion_Indicator.
  """
  
  #port = str(design.port_entry.text())
  #print port
  #baud = str(design.baud_rate.text())
  #print baud

  #Com_open(port, baud)
  

  while(1):  ## Read data infinitley until the disconnect button is pressed.

    print "helloworld"
      

    while (Com_open(port, baud).serial_object.inWaiting()==0): #Wait here until there is data.
        print Com_open(port, baud).serial_object.inWaiting()
        print "stuck"
        time.sleep(.1) #1
        
        pass #do nothing
    

    try:       
        
        serial_data=ser.serial_object.readline().strip('\n').strip('\r')
      
        print serial_data
                                
    
    
        filter_data = serial_data.split(',')
        
    
    
        print(filter_data)

        
        if (filter_data):
          
            # Expand the label box to cover old repeated text in x and y directions
            # Storing data in variables
            # Display variables on the screen in text
            # Data Label 1                               temp1
                                        
            design.QLabel_data1.setText(filter_data[1])
            
            # Data Label 2                               temp2
            design.QLabel_data2.setText(filter_data[2])

            # Data Label 3                            pressure1
            
            design.QLabel_data3.setText(filter_data[3])

            # Data Label 4                          pressure2                            
        
            design.Qlabel_data4.setText(filter_data[4])

            #Column 2 from the left
            ######################################

            # Data Label 5                          temp3
            
            design.Qlabel_data5.setText(filter_data[5])

            # Data Label 6                          temp4
            
            design.Qlabel_data6.setText(filter_data[6])

            # Data Label 7                          pressure3
          
            design.Qlabel_data7.setText(filter_data[7])

            # Data Label 8                         pressure4
            
            design.Qlabel_data8.setText(filter_data[8])

            #Column 3 from the left
            #######################################
            # Data Label 9                          temp5
            
            design.Qlabel_data9.setText(filter_data[9])

            # Data Label 10                         temp6
            
            design.Qlabel_data10.setText(filter_data[10])

            # Data Label 11                         pressure5
            
            design.Qlabel_data11.setText(filter_data[11])

            # Data Label 12                         pressure6
            
            design.Qlabel_data4.setText(filter_data[12])

            #Column 4 Valves   Indicators
            ######################################

            # Data Pressurant_Fill Valve Indicator           Pressurant_Fill_Indicator
            design.Qlabel_data13.setText(filter_data[13])

            # Data Pressurant_Oxidizer Valve Indicator       Pressurant_Oxidizer_Indicator
            
            design.Qlabel_data14.setText(filter_data[14])

            # Data Oxidizer_fill Valve Indicator             Oxidizer_Fill_Indicator
            
            design.Qlabel_data15.setText(filter_data[15])

            # Data Oxi_Combustion Valve Indicator            Oxidizer_Combustion_Indicator
            
            design.Qlabel_data16.setText(filter_data[16])

            #######################################




            
            ############# Convert data into numbers for warning messages
            MissionTime=float(filter_data[0])
            PressurantTemp= float(filter_data[1])
            PressurantPressure = float(filter_data[3])
            OxidizerTemp= float(filter_data[5])
            OxidizerPressure = float(filter_data[7])
            CombustionPressure=float(filter_data[11])
            

            try:
                with open("output.csv","a+") as resultFile:   # Save to the excel file
                    wr = csv.writer(resultFile, dialect='excel')
                    wr.writerow(filter_data)
            except:
                pass
            
    except:
        pass

'''



def main():
  
  app = QtGui.QApplication(sys.argv)

 
  
  x=design()
  
  



  x.w.show()
  sys.exit(app.exec_())

  


 






if __name__ == '__main__':
    main()
