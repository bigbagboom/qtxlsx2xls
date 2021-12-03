# -*- coding: utf-8 -*-

import sys
import os
from xlsx2xls import xlsx2xls, xls2xlsx

from PyQt5 import QtCore  #, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QDialog, QFileDialog
from PyQt5.QtCore import pyqtSlot, QThread, QBasicTimer
from PyQt5.QtGui import QTextCursor

from Ui_mainwin import Ui_Dialog

class custom_mainwin(QDialog, Ui_Dialog):
    
    def __init__(self, parent=None):
        super(custom_mainwin, self).__init__(parent)
        self.timer = QBasicTimer()
        self.setupUi(self)
        self.custom_setupUi()
    
    def custom_setupUi(self):

        self.f = FakeOut()
        self.old = sys.stdout
        self.olderr = sys.stderr
        
        sys.stdout = self.f
        sys.stderr = self.f
        
        self.pushButton_3.setEnabled(False)
        
        self.timer.start(10, self)
        
        self._want_to_close = True
        #self._want_to_close = False
    
    def timerEvent(self, e):
        pass
        if  self.f.str != self.plainTextEdit.toPlainText():
            self.plainTextEdit.setPlainText(self.f.str)      
            self.plainTextEdit.moveCursor(QTextCursor.End)        
        
    def closeEvent(self, evnt):
        if self._want_to_close:
            sys.stdout = self.old
            sys.stderr = self.olderr
            self.timer.stop()
            super(custom_mainwin, self).closeEvent(evnt)
        else:
            evnt.ignore()
            
    def finished(self):

        self._want_to_close = True

        self.plainTextEdit.setPlainText(self.f.str)
        self.plainTextEdit.moveCursor(QTextCursor.End)

    def disable_controls(self):
        self._want_to_close = False
        self.pushButton.setEnabled(False)
        self.pushButton_2.setEnabled(False)
        self.pushButton_3.setEnabled(False)
        self.pushButton_4.setEnabled(False)

    def reenable_controls(self):
        self._want_to_close = True
        self.pushButton.setEnabled(True)
        self.pushButton_2.setEnabled(True)
        self.pushButton_3.setEnabled(True)
        self.pushButton_4.setEnabled(True)
        
    @pyqtSlot()
    def on_pushButton_2_clicked(self):
        """
        Slot documentation goes here.
        """

        self.disable_controls()
        
        dir2 = "."

        dir= QFileDialog.getExistingDirectory (self,"选择文件存放路径", dir2, QFileDialog.ShowDirsOnly)
        if dir:
            dir2 = dir.replace('/', '\\')
            self.lineEdit_3.setText(dir2)
        
        self.reenable_controls()


    @pyqtSlot()
    def on_pushButton_3_clicked(self):
        """
        Slot documentation goes here.
        """
        
        self.workpath = self.lineEdit_3.text()
        
        if os.path.isdir(self.workpath):
            self.disable_controls()
            self.LongRunJob = LongRunJob(self)
            self.LongRunJob.start()

    @pyqtSlot()
    def on_pushButton_4_clicked(self):
        """
        Slot documentation goes here.
        """
        #self.disable_controls()
        
        self.f.str = ''        
        
        #self.reenable_controls()


class FakeOut:
    
    def __init__(self):
        self.str=''
        self.n = 0
    
    def write(self,s):
        self.str += s
        self.n+=1
    
    def show(self): #显示函数，非必须
        print(self.str) 
    
    def clear(self): #清空函数，非必须
        self.str = ''
        self.n = 0

class LongRunJob(QThread):

    def __init__(self, parent = None):
        super(QThread, self).__init__()
        self.parent = parent 
        
    def xlsx2xlsProcess(self):

        xlsxlist=[]
        xlslist=[]

        if self.parent !=None:
            self.workpath = self.parent.lineEdit_3.text()
            self.parent.f.str = ''  
            self.parent.f.str += '开始批量转换 xlsx 到 xls\n'
            
            for i in os.listdir(self.workpath):
                if i[-4:]=='xlsx':
                    xlsxlist.append(i)
                if i[-3:]=='xls':
                    xlslist.append(i)
            for i in xlsxlist:
                i2 = i.replace('.xlsx', '.xls')
                if i2 not in xlslist:
                    self.parent.f.str += '转换 ' + i + ' 为 ' + i2 + '\n'
                    filepath = os.path.join(self.workpath, i)
                    xlsx2xls(filepath)

            self.parent.f.str += '批量转换完成\n'

        self.reenable_controls()
    
    def xls2xlsxProcess(self):

        xlsxlist=[]
        xlslist=[]

        if self.parent !=None:
            self.workpath = self.parent.lineEdit_3.text()
            self.parent.f.str = ''  
            self.parent.f.str += '开始批量转换 xls 到 xlsx\n'
            
            for i in os.listdir(self.workpath):
                if i[-4:]=='xlsx':
                    xlsxlist.append(i)
                if i[-3:]=='xls':
                    xlslist.append(i)
            for i in xlslist:
                i2 = i.replace('.xls', '.xlsx')
                if i2 not in xlsxlist:
                    self.parent.f.str += '转换 ' + i + ' 为 ' + i2 + '\n'
                    filepath = os.path.join(self.workpath, i)
                    xls2xlsx(filepath)

            self.parent.f.str += '批量转换完成\n'

        self.reenable_controls()
    
    def run(self):

        if self.parent.radioButton.isChecked():
            self.xlsx2xlsProcess()
        else:
            self.xls2xlsxProcess()
        #print('工作完成!')
        #sys.stderr.write('Job done!')
        
    def reenable_controls(self):
        if self.parent !=None:
            self.parent._want_to_close = True
            self.parent.pushButton.setEnabled(True)
            self.parent.pushButton_2.setEnabled(True)
            self.parent.pushButton_3.setEnabled(True)
            self.parent.pushButton_4.setEnabled(True)

if __name__ == "__main__":

    prgpath=os.path.dirname(os.path.realpath(__file__))
        
    app = QApplication(sys.argv)
    win=custom_mainwin()
    win.show()
    sys.exit(app.exec_())
