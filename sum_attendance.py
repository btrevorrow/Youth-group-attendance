#! python3
# -*- coding: utf-8 -*-
"""
Created on Thu Aug 30 14:11:11 2018

@author: btrev

A GUI which accepts churchbuilder .ods attendance spreadsheets and creates a
new spreadsheet which displays the attendance totals for each person.
"""

import sys, re, os
import sheet_fns
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtWidgets import (QWidget, QMainWindow, QPushButton, QAction, 
                             QLabel, QVBoxLayout, QGridLayout, QApplication, 
                             QFileDialog, QInputDialog, QMessageBox)


class FileDisplay(QWidget):
    
    def __init__(self):
        super().__init__()
        
        self.initUI()
        
    def initUI(self):
        
        self.grid = QGridLayout()
        #store the widgets in the grid
        self.entries = []
        #store filenames and group names
        self.registers = {}
        
        #headings
        self.grid.addWidget(QLabel('<b>File<b>', self), 0, 0)
        self.grid.addWidget(QLabel('<b>Group<b>', self), 0,2)
        self.grid.addWidget(QLabel('<b>Del<b>', self), 0,3)
        
        vbox = QVBoxLayout()
        vbox.addLayout(self.grid)
        vbox.addStretch(1)
        self.setLayout(vbox)
        
    def add_file(self, file, group):
        
        #show a truncated path
        pathRegex = re.compile(r'/[^/]+/[^/]+\.ods')
        shortPath = '...' + pathRegex.search(file[0]).group()
        
        imgLbl = QLabel(self)
        imgLbl.setPixmap(QPixmap('file.png'))
        fileLbl = QLabel(shortPath, self)
        grpLbl = QLabel(group, self)
        delbutton = QPushButton('',self)
        delbutton.setIcon(QIcon('delbox.png'))
        delbutton.clicked.connect(self.delClicked)
        
        self.entries.append([imgLbl, fileLbl, grpLbl, delbutton])
        rows = len(self.entries)
        
        for i, label in enumerate(self.entries[-1]):
            self.grid.addWidget(label, rows, i)
        
        
        #save file path and group name for use with sheet_fns
        self.registers[group] = file[0]
#        registerLbl = QLabel(str(self.registers),self)
#        self.grid.addWidget(registerLbl, rows+1, 1)
        
#        entiresLbl = QLabel(str(self.entries),self)
#        self.grid.addWidget(entiresLbl, rows+1, 1)
        
    def delClicked(self):
        
        sender = self.sender()
        for i, innerL in enumerate(self.entries):
            if sender in innerL:
                group = innerL[2].text()
                for wid in innerL:
                    self.grid.removeWidget(wid)
                    wid.deleteLater()
                del self.entries[i]
                del self.registers[group]
                

class MainWindow(QMainWindow):
    
    def __init__(self):
        super().__init__()
        
        self.initUI()
        
    def initUI(self):
        
        self.fileDisp = FileDisplay()
        self.setCentralWidget(self.fileDisp)
        
        #Set up the menubar and toolbar with actions
        addAct = QAction(QIcon('addbox.png'), '&Add', self)
        addAct.setStatusTip('Add a spreadsheet to the list.')
        addAct.triggered.connect(self.openDialog)
        
        genAct = QAction(QIcon('sheet.png'), '&Generate', self)
        genAct.setStatusTip('Generate a totals spreadsheet.')
        genAct.triggered.connect(self.generate_totals)
        
        self.statusBar = self.statusBar()
        
        menubar = self.menuBar()
        fileMenu = menubar.addMenu('&File')
        fileMenu.addAction(addAct)
        fileMenu.addAction(genAct)
        
        self.toolbar = self.addToolBar('Add file')
        self.toolbar.addAction(addAct)
        self.toolbar.addAction(genAct)
        
        self.setGeometry(1000, 200, 500, 350)
        self.setWindowTitle('Youth Group Attendance')
        self.setWindowIcon(QIcon('tickbox.png'))
        self.show()
        
    def openDialog(self):

        while True:
            fname = QFileDialog.getOpenFileName(self,'Open File', '.', '*.ods')
            if fname[0] not in self.fileDisp.registers.values():
                break
            else:
                QMessageBox.warning(self, 'Warning', 'File already '
                                            'chosen, choose a different file.',
                                            QMessageBox.Ok)
            
        
        if fname[0]:
            while True:
                group, ok = QInputDialog.getText(self, 'Input Dialog', 'Enter ' 
                                          'the group name for this register: ')
                if group not in self.fileDisp.registers:
                    break
                else:
                    QMessageBox.warning(self, 'Warning', 'Group name already '
                      'exists, choose a different group name.', QMessageBox.Ok)
            if ok:
                self.fileDisp.add_file(fname, group)
            else:
                self.fileDisp.add_file(fname, '')
                
    def generate_totals(self):
        
        registers = self.fileDisp.registers
        if registers == {}:
            QMessageBox.warning(self, 'Warning', 'No data available. Please '
                                'add a spreadsheet.', QMessageBox.Ok)
        else:
            try:
                sheet_fns.convert_to_xlsx(registers)
                sheet_fns.check_dates(registers)
                sheet_fns.check_names(registers)
                sheet_fns.check_Ys(registers)
               
            except Exception as err:
                QMessageBox.warning(self, 'Warning', err, QMessageBox.Ok)
            else:
                attendance = sheet_fns.get_names(registers)
                sheet_fns.sum_attendance_data(attendance, registers)
                #return a list of group names for the write function
                groups = list(registers.keys())
                saveName = QFileDialog.getSaveFileName(self, 'Save file', '.',
                                                       '*.xlsx')
                sheet_fns.write_totals_sheet(attendance, groups, saveName[0])
                
                ans = QMessageBox.question(self, 'Spreadsheet generated','What'
                                     ' would you like to do with {0}?'
                                     .format(os.path.basename(saveName[0])), 
                                     QMessageBox.Open | QMessageBox.Cancel, 
                                     QMessageBox.Cancel)
                if ans == QMessageBox.Open:
                    os.startfile(saveName[0])
            finally:
                for filename in registers.values():
                    if os.path.isfile(filename) and filename[-5:] == '.xlsx':
                        os.remove(filename)
                
        
if __name__ == '__main__':
    
    app = QApplication(sys.argv)
    mw = MainWindow()
    sys.exit(app.exec_())