#!/usr/bin/python3
# -*- coding: utf-8 -*-

import pandas as pd 
from PyQt5.QtWidgets import (QWidget, QLabel, 
    QComboBox, QApplication, QPushButton)
import sys
import os

grad_file = 'Graduation data.xls'
reg_file = 'Reg data (1140).xls'
grad_data = pd.read_excel(grad_file, sheet_name = 0)
reg_data = pd.read_excel(reg_file, sheet_name = 0)

class Example(QWidget):
    selected_sem = 0
    
    def __init__(self):
        super().__init__()
        
        self.initUI()
        
        
    def initUI(self):      

        self.lbl_col1 = QLabel("Registered", self)
        self.lbl_col2 = QLabel("Graduated", self)
        self.lbl_row1 = QLabel("Total", self)
        self.lbl_row2 = QLabel("311", self)
        self.lbl_row3 = QLabel("312", self)
        self.lbl_row4 = QLabel("313", self)
        self.lbl_grad_count = QLabel("-", self)
        self.lbl_reg_count = QLabel("-", self)
        self.lbl_grad_311_count = QLabel("-", self)
        self.lbl_grad_312_count = QLabel("-", self)
        self.lbl_grad_313_count = QLabel("-", self)
        self.lbl_reg_311_count = QLabel("-", self)
        self.lbl_reg_312_count = QLabel("-", self)
        self.lbl_reg_313_count = QLabel("-", self)


        combo = QComboBox(self)
        combo.addItem("Select Semester")
        combo.addItem("1139")
        combo.addItem("1140")
        combo.addItem("1141")
        combo.addItem("1142")
        combo.addItem("1143")
        combo.addItem("1144")
        combo.addItem("1145")

        combo.move(50, 50)
        self.lbl_col1.move(150, 100)
        self.lbl_col2.move(300, 100)
        self.lbl_row1.move(50, 150)
        self.lbl_row2.move(50, 200)
        self.lbl_row3.move(50, 250)
        self.lbl_row4.move(50, 300)
        self.lbl_reg_count.move(150, 150)
        self.lbl_reg_311_count.move(150, 200)
        self.lbl_reg_312_count.move(150, 250)
        self.lbl_reg_313_count.move(150, 300)
        self.lbl_grad_count.move(300, 150)
        self.lbl_grad_311_count.move(300, 200)
        self.lbl_grad_312_count.move(300, 250)
        self.lbl_grad_313_count.move(300, 300)

        combo.activated[str].connect(self.onActivated)        
         
           
        button = QPushButton('Show data', self)
        button.setToolTip('This button shows data')
        button.move(50,350)
        button.clicked.connect(self.on_click)
        
        
        self.setGeometry(300, 300, 450, 450)
        self.setWindowTitle('Python GUI')
        self.show()
        
        
    def onActivated(self, text):
        if text == "Select Semester":
            self.selected_sem = 0
            text = "-"
            self.lbl_grad_count.setText(text)
            self.lbl_grad_count.adjustSize() 
            self.lbl_reg_count.setText(text)
            self.lbl_reg_count.adjustSize() 
            self.lbl_reg_311_count.setText(text)
            self.lbl_reg_311_count.adjustSize() 
            self.lbl_grad_311_count.setText(text)
            self.lbl_grad_311_count.adjustSize() 
            self.lbl_reg_312_count.setText(text)
            self.lbl_reg_312_count.adjustSize() 
            self.lbl_grad_312_count.setText(text)
            self.lbl_grad_312_count.adjustSize() 
            self.lbl_reg_313_count.setText(text)
            self.lbl_reg_313_count.adjustSize() 
            self.lbl_grad_313_count.setText(text)
            self.lbl_grad_313_count.adjustSize() 
        
        else:
        
            sem = int(text)
            self.selected_sem = sem
            
            grad_in_sem = grad_data[grad_data["semester.1"]==sem]
            count_grad = grad_in_sem.shape[0]
            text = str(count_grad)
            self.lbl_grad_count.setText(text)
            self.lbl_grad_count.adjustSize() 
            
            reg_data['ID'] = reg_data['ID'].astype(str)
            
            reg_in_sem = reg_data[reg_data["Semester"]==self.selected_sem]
            reg_in_sem_unique = reg_in_sem.drop_duplicates(subset=['ID'], keep='last')
            reg_in_sem_unique = reg_in_sem_unique.loc[reg_in_sem_unique['ID'].str.match('311') | reg_in_sem_unique['ID'].str.match('312') | reg_in_sem_unique['ID'].str.match('313')]
            count_reg = reg_in_sem_unique.shape[0]
            text = str(count_reg)
            self.lbl_reg_count.setText(text)
            self.lbl_reg_count.adjustSize() 
            
            count_grad_311 = 0
            count_grad_312 = 0
            count_grad_313 = 0
            
            reg_311 = reg_in_sem_unique.loc[reg_in_sem_unique['ID'].str.match('311')]
            count_reg_311 = reg_311.shape[0]
            text = str(count_reg_311)
            self.lbl_reg_311_count.setText(text)
            self.lbl_reg_311_count.adjustSize() 
            bits_id_list = reg_311["Campus ID"].to_numpy()
            for ID in bits_id_list:
                for grad_id in grad_in_sem["idno"]:
                    if ID == grad_id:
                        count_grad_311 = count_grad_311 + 1
            self.lbl_grad_311_count.setText(str(count_grad_311))
            self.lbl_grad_311_count.adjustSize() 
            
            reg_312 = reg_in_sem_unique.loc[reg_in_sem_unique['ID'].str.match('312')]
            count_reg_312 = reg_312.shape[0]
            text = str(count_reg_312)
            self.lbl_reg_312_count.setText(text)
            self.lbl_reg_312_count.adjustSize() 
            bits_id_list = reg_312["Campus ID"].to_numpy()
            for ID in bits_id_list:
                for grad_id in grad_in_sem["idno"]:
                    if ID == grad_id:
                        count_grad_312 = count_grad_312 + 1
            self.lbl_grad_312_count.setText(str(count_grad_312))
            self.lbl_grad_312_count.adjustSize() 
            
            reg_313 = reg_in_sem_unique.loc[reg_in_sem_unique['ID'].str.match('313')]
            count_reg_313 = reg_313.shape[0]
            text = str(count_reg_313)
            self.lbl_reg_313_count.setText(text)
            self.lbl_reg_313_count.adjustSize() 
            bits_id_list = reg_313["Campus ID"].to_numpy()
            for ID in bits_id_list:
                for grad_id in grad_in_sem["idno"]:
                    if ID == grad_id:
                        count_grad_313 = count_grad_313 + 1
            self.lbl_grad_313_count.setText(str(count_grad_313))
            self.lbl_grad_313_count.adjustSize() 
        
    
    def on_click(self):
        if self.selected_sem != 0:
            grad_in_sem = grad_data[grad_data["semester.1"]==self.selected_sem]
            grad_in_sem.to_excel("output.xlsx", index=False)
            os.system('libreoffice --calc output.xlsx')
                
if __name__ == '__main__':
    
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())