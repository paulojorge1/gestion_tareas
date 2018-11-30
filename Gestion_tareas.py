#!/usr/bin/python
# -*- coding: utf-8 -*-


from PyQt4 import QtGui, QtCore
import csv, sys, os
import pandas as pd
import sqlite3
#from pandas.io import sql
#try:
#    import StringIO
#except ImportError:
#    from io import StringIO
import datetime
#import pythoncom
import win32com.client
#from pandasql import sqldf
#pysqldf = lambda q: sqldf(q, globals()) 

#print sys.path

fichero="C:\\Data\\automatizar_report_horas\\" 
conn = sqlite3.connect("%s%s"%(fichero,"report_tareas.sl3"))
conn.text_factory = str
curs = conn.cursor() 

class Gestion(QtGui.QMainWindow):
    
    def __init__(self):
        super(Gestion, self).__init__()
        os.chdir('C:\\Data\\automatizar_report_horas\\') # Set working directory    
        self.initUI()
       
    def initUI(self):      

        openFile = QtGui.QAction(QtGui.QIcon('open.png'), 'Importar', self)
        openFile.setShortcut('Ctrl+O')
        openFile.setStatusTip('Importar tareas a outlook')
        openFile.triggered.connect(self.showDialog)
        openFile.triggered.connect(self.buttonClicked)          
        
        saveFile = QtGui.QAction(QtGui.QIcon('save.png'), 'Exportar', self)
        saveFile.setShortcut('Ctrl+S')
        saveFile.setStatusTip(u'Exportar histórico de tareas')
        saveFile.triggered.connect(self.showDialog)
        saveFile.triggered.connect(self.buttonClicked)   
        
        exitAction = QtGui.QAction("Salir",self) 
        exitAction.setShortcut('Ctrl+Q')
        exitAction.setStatusTip('Salir de la app')       
        exitAction.triggered.connect(self.close)
        
        menubar = self.menuBar()
        fileMenu = menubar.addMenu('&Menu')
        fileMenu.addAction(openFile)  
        fileMenu.addAction(saveFile)  
        fileMenu.addAction(exitAction) 
        
        btn1 = QtGui.QPushButton("Importar", self)
        btn1.setStatusTip('Importar tareas a outlook')    
        btn1.move(30, 50)

        btn2 = QtGui.QPushButton("Exportar", self)
        btn2.setStatusTip(u'Exportar histórico de tareas')    
        btn2.move(150, 50)
        
        btn1.clicked.connect(self.showDialog) 
        btn1.clicked.connect(self.buttonClicked)  
                
        btn2.clicked.connect(self.showDialog1)  
        btn2.clicked.connect(self.buttonClicked)

        self.statusBar()
       
        self.setGeometry(300, 300, 290, 150)
        self.setWindowTitle(u'Gestión de tareas')
#        self.show()
        
    def buttonClicked(self):
      
        sender = self.sender()
        self.statusBar().showMessage(sender.text() + ' : hecho', 50000)
    
    def showDialog(self):
        
#        fname = QtGui.QFileDialog.getOpenFileName(self, 'Importar fichero', 
#                'C:/Data/automatizar_report_horas/ScoreCard.csv')
        
#        fichero = QtGui.QFileDialog.getOpenFileName(self, 'Importar fichero', 
#                'C:/Data/automatizar_report_horas/ScoreCard.xlsx')
        
        tareas=pd.io.excel.read_excel("%s%s"%(fichero,"ScoreCard.xlsx"),sheetname="subir outlook",parse_cols=(0,1,2,5,8,9))
        curs.execute("""DROP TABLE IF EXISTS tareas;""") 
        pd.DataFrame.to_sql(tareas, name='tareas', con=conn)  
        
        for index, row in tareas.iterrows():
            
            start = row["Start1"]
#            start = start.encode('utf-8')
            
#            subject = str(row["Subject"])
            subject = row["Subject"]
#            subject = subject.encode('utf-8')
#            subject = subject.decode('ascii', 'ignore')
            
            categories = row["Categories"]
#            categories = categories.encode('utf-8')
#            categories = categories.decode('ascii', 'ignore')
            
            duration = row["Duration"]
#            duration = duration.encode('utf-8')
            
            self.addEvent(start, subject, categories, duration)  
        
    def keyPressEvent(self, e):
        
        if e.key() == QtCore.Qt.Key_Escape:
            self.close()        
           
    # importar eventos para outlook
    def addEvent(self, start, subject, categories, duration):

        oOutlook = win32com.client.Dispatch("Outlook.Application")
        appointment = oOutlook.CreateItem(1) # 1=outlook appointment item
        appointment.Start = start
        appointment.Subject = subject
        appointment.Categories = categories        
        appointment.Duration = duration
        appointment.Location = 'Mi sitio'
        appointment.ReminderSet = True
        appointment.ReminderMinutesBeforeStart = 15
        appointment.Save()
        return

    def showDialog1(self):
        
        Outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = Outlook.Session 
        myCalendar = namespace.GetDefaultFolder(9)
        self.makeTable(myCalendar)

#    def getItemProperty(self, propertyName):  
#
#        try:
#            result = getattr(self, propertyName)
#        except pythoncom.com_error as ce: 
#            result = ce.excepinfo[2]
#        return result
   
    def makeTable(self, myCalendar):
    
        csvfile = "C:\\data\\automatizar_report_horas\\datos.csv" 
    
        items = myCalendar.Items
        
        begin = datetime.date.today() - datetime.timedelta(days = 90);
        end = datetime.date.today() + datetime.timedelta(days = 30);
    #    restriction = "[Start] >= '" + begin.strftime("%m/%d/%Y") + "' AND [End] <= '" + end.strftime("%m/%d/%Y") + "'"
        restriction = "[Start]  >='" + begin.strftime("%d/%m/%Y") + "' AND [End] <= '" + end.strftime("%d/%m/%Y") + "'"
        restrictedItems = items.Restrict(restriction)
        
        with open(csvfile, "w") as output:
            writer = csv.writer(output, lineterminator='\n')
            for appointmentItem in restrictedItems:
    
                startDate = getattr(appointmentItem, "Start")
#                startDate = startDate.encode('utf-8')
                
                endDate = getattr(appointmentItem, "End")
#                endDate = endDate.encode('utf-8')            
                
                subject = getattr(appointmentItem, "Subject")
                subject = subject.encode('utf-8')
                                
                organizer = getattr(appointmentItem, "Organizer")
                organizer = organizer.encode('utf-8')
                
                categories = getattr(appointmentItem, "Categories")
                categories = categories.encode('utf-8')
                
                writer.writerow([startDate]+[endDate]+[subject]+[organizer]+[categories])

        curs.execute("""DROP TABLE IF EXISTS datos;""")
        curs.execute("""CREATE TABLE datos (
                    startDate REAL,
                    endDate datetime,
                    subject text,
                    organizer text,
                    categories text
                    );""")
    
        datos = csv.reader(open("%s%s"%(fichero,"datos.csv"), 'r'), delimiter=',')
#        next(datos, None)
#        e=1
        for row in datos:
#            e=e+1
            datos =[row[0], row[1], row[2], row[3], row[4]]
            curs.execute("""INSERT INTO datos (startDate, endDate, subject, organizer, categories) VALUES (?, ?, ?, ?, ?);""", datos)
            conn.commit()
        
        dateval=pd.io.excel.read_excel("%s%s"%(fichero,"calendar.xlsm"),sheetname="DATEVAL")
        curs.execute("""DROP TABLE IF EXISTS dateval;""") 
        pd.DataFrame.to_sql(dateval, name='dateval', con=conn)
#        dateval=pd.read_sql_query("select * from dateval;",conn)
    
        calendario_vacas_gestion=pd.io.excel.read_excel("%s%s"%(fichero,"calendar.xlsm"),sheetname="calendario_vacas_gestion")
        curs.execute("""DROP TABLE IF EXISTS calendario_vacas_gestion;""") 
        pd.DataFrame.to_sql(calendario_vacas_gestion, name='calendario_vacas_gestion', con=conn)    
#        calendario_vacas_gestion=pd.read_sql_query("select * from calendario_vacas_gestion;",conn)
    
def main():
    
    app = QtGui.QApplication(sys.argv)
    ex = Gestion()
    ex.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
    

