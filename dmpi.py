
import sys
import win32clipboard
import win32con
import subprocess
import os
import configparser

from PyQt5 import QtCore, QtWidgets
from mpi import Ui_MainWindow
from form import Ui_formDialog
from pywinauto.application import Application
from pywinauto.findwindows import find_windows

class MyForm(QtWidgets.QDialog):
    def __init__(self, parent, file):
        QtWidgets.QDialog.__init__(self, parent)
        self.form = Ui_formDialog()
        self.form.setupUi(self)
        self.form.file = file


        self.form.plainTextEdit.textChanged.connect(self.textchange)

        try:
            f = open(file, 'r', encoding='utf-8-sig')
        except UnicodeDecodeError:
            f = open(file, 'r', encoding='cp950')
        else:
            MainText = os.linesep.join(f.read().split('\n'))
            f.close()
        self.form.plainTextEdit.setPlainText(MainText)
        self.form.TextModified = False

        return

    def textchange(self):
        self.form.TextModified = True
        return

    def isTextchange(self):
        return self.form.TextModified

    def getInputs(self):
        return self.form.plainTextEdit.toPlainText()




class MyWin(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        QtWidgets.QWidget.__init__(self, parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
          
        # set status change and color change
        self.ui.Ant.currentIndexChanged.connect(lambda index: self.ui.label_ant.setStyleSheet("QLabel {color:%s}" % ('red' if index else 'black')))
        self.ui.Sep.currentIndexChanged.connect(lambda index: self.ui.label_sep.setStyleSheet("QLabel {color:%s}" % ('red' if index else 'black')))
        self.ui.Inf.currentIndexChanged.connect(lambda index: self.ui.label_inf.setStyleSheet("QLabel {color:%s}" % ('red' if index else 'black')))
        self.ui.Lat.currentIndexChanged.connect(lambda index: self.ui.label_lat.setStyleSheet("QLabel {color:%s}" % ('red' if index else 'black')))
        self.ui.AP.currentIndexChanged.connect(lambda index: self.ui.label_ap.setStyleSheet("QLabel {color:%s}" % ('red' if index else 'black')))
        
        self.ui.SEF.returnPressed.connect(lambda:self.PresseEnter('SEF'))
        self.ui.REF.returnPressed.connect(lambda:self.PresseEnter('REF'))
        self.ui.SLH.returnPressed.connect(lambda:self.PresseEnter('SLH'))
        self.ui.RLH.returnPressed.connect(lambda:self.PresseEnter('RLH'))
        self.ui.Ext.returnPressed.connect(lambda:self.PresseEnter('Ext')) 
        
        self.ui.Addlesion.clicked.connect(self.Addlesion)
        self.ui.CleanL.clicked.connect(self.CleanL)
        self.ui.AddF.clicked.connect(self.AddFunction)
        self.ui.Send.clicked.connect(self.SendF)
        self.ui.CleanA.clicked.connect(self.Cleana)
        self.ui.MBFQSend.clicked.connect(self.mbfqSend)
        self.ui.Load.clicked.connect(self.LoadF)
        self.ui.Sign.clicked.connect(self.SignF)
        self.ui.actionMPI_2.triggered.connect(self.MPImenu)
        self.ui.actionMBFQ.triggered.connect(self.MBFQmenu)


        return

    def MPImenu(self):
        file = './setting/MainText.txt'
        myfrom = MyForm(myapp, file)
        if myfrom.exec_() and myfrom.isTextchange():
            Text = myfrom.getInputs()
            try:
                f = open(file, 'w', encoding='utf-8-sig')
            except UnicodeDecodeError:
                f = open(file, 'w', encoding='cp950')
            else:
                f.write(Text)
                f.close()
                global MainText
                MainText = Text

        return

    def MBFQmenu(self):
        file = './setting/MBFQText.txt'
        myfrom = MyForm(myapp, file)
        if myfrom.exec_() and myfrom.isTextchange():
            Text = myfrom.getInputs()
            try:
                f = open(file, 'w', encoding='utf-8-sig')
            except UnicodeDecodeError:
                f = open(file, 'w', encoding='cp950')
            else:
                f.write(Text)
                f.close()
                global  MBFQ_Text
                MBFQ_Text = Text

        return
        
    def GetClipboard(self):
        win32clipboard.OpenClipboard()
        temp = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()
        return temp
    
    def SetClipboard(self, text):
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
        win32clipboard.CloseClipboard()
        return

    def SignF(self):
        w_handle = find_windows(title='EBM UniReport') # find the EBM if exist or not
        if not w_handle:
            self.SetStatusBar('red', 'EMB report is not found')
        else:
            # print("I find the EBM Report!!!")
            app = Application().connect(title_re="EBM UniReport")
            title = str(app.TMainForm.element_info)
            Title = title.split()
            if 'ID' in Title:
                r = ""
                v = ""
                if self.ui.RBox.currentIndex():
                    r = self.ui.RBox.currentText()
                if self.ui.VBox.currentIndex():
                    v = self.ui.VBox.currentText()
                ctext = 'IamSep' + r + 'IamSep' + v
                temp = self.GetClipboard()
                self.SetClipboard(ctext)
                os.system("sender.ahk")
                self.SetClipboard(temp)
            else:
                self.SetStatusBar('red', 'EBM report is not in EIDT window')

        return

    def AddComm(self, n, nline):
        Text = self.ui.Comm.toPlainText().splitlines()
        #print(Text)
        if not n:
            Text.append(str(len(Text) + 1) + ". " + nline)
        elif n > len(Text):
            Text.append(str(len(Text) + 1) + ". " + nline)
        else:
            Text[n-1] = str(n) + ". " + nline

        self.ui.Comm.setText(os.linesep.join(Text))
        return
    
    def SetStatusBar(self, color, text):
        self.ui.StatusBar.setStyleSheet("color:" + color)
        self.ui.StatusBar.setText(text)
        #self.ui.StatusBar.setStyleSheet('color: black')
        
        return

    def LoadF(self):
        entryList = {}
        entryStatus = {}
        #exchange = {'partial':'partial reversible'}
        w_handle = find_windows(title='EBM UniReport') # find the EBM if exist or not
        if not w_handle:
            self.SetStatusBar('red', 'EMB report is not found, try in clipboard')
            Text = self.GetClipboard()
        else:
            # print("I find the EBM Report!!!")
            app = Application().connect(title_re="EBM UniReport")
            title = str(app.TMainForm.element_info)
            Title = title.split()
            if 'ID' in Title:
                #ahk_res = subprocess.check_output(["loader.exe", "*"], shell=False)
                #Text = ahk_res.decode("utf8")
                os.system("loader.exe")
                Text = self.GetClipboard()
            else:
                self.SetStatusBar('red', 'EBM report is not in EIDT window')

        if Text.find("NUCLEAR MEDICINE REPORT") != -1:
            Text = Text.split('NUCLEAR CARDIOLOGY DIAGNOSIS :'+os.linesep)
            self.ui.Comm.setText(Text[1])
            textList = Text[0].splitlines()

            for text in textList:
                #print(text)
                if text.find("myocardial perfusion scan") != -1:
                    tracer = text.split('myocardial')[0].strip()
                    if tracer == "Tl-210":
                        self.ui.Tl201.setChecked(1)
                    elif tracer == "Tc-99m sestamibi":
                        self.ui.MiBi.setChecked(1)
                    else:
                        self.ui.Tl201.setChecked(1)
                if text.find("Abnormal uptake,") != -1:
                    tWall = text.split()[1][:3]
                    if tWall == "Api":
                        tWall = "AP"
                    tStatus = text.split()[6].capitalize().strip('.')
                    if tStatus == "Partial":
                        tStatus = "Partial reversible"
                    elif tStatus == "Reverse":
                        tStatus = "Reverse redistribution"
                    entryStatus[tWall] = tStatus
                if text.find("(stress)") != -1:
                    self.ui.SLH.setText(text.split()[8])
                if text.find("(rest)") != -1:
                    self.ui.RLH.setText(text.split()[3])

            for combo in self.ui.groupStatus_2.findChildren(QtWidgets.QComboBox):
                    if combo.objectName() in entryStatus:
                        combo.setCurrentText(entryStatus[combo.objectName()])
        else:
            self.SetStatusBar('red', '不符報告格式')

            #print(entryStatus)

        return

    def Addlesion(self):  # new lesion added
        Deep = ""
        Wallstatus = {'Ant':0,'AL':0,'Lat':0,'IL':0,'Inf':0,'IS':0,'Sep':0,'AS':0,'AP':0}
        
        for radiobox in self.ui.groupStatus.findChildren(QtWidgets.QRadioButton):
            if radiobox.isChecked():
                Status = radiobox.text()
                
        if not (Status == 'Attenuation' or Status == 'Hypokinesia'):
            # get the deep
            for checkbox in self.ui.groupDeep.findChildren(QtWidgets.QCheckBox): 
                if checkbox.isChecked():
                    checkbox.setChecked(0)
                    if Deep == "":
                        Deep = checkbox.text()
                    else:
                        Deep = Deep + "/" + checkbox.text()
            # print('%s: %s' % (checkbox.text(), checkbox.isChecked()))
            
            if Deep == "apical/mid/basal" or Deep == "":
                Deep = ""
            else:
                Deep = Deep + '-'
            
        for checkbox in self.ui.groupWall.findChildren(QtWidgets.QCheckBox):
            if checkbox.isChecked():
                Wallstatus[checkbox.text()] = 1
                checkbox.setChecked(0)
                if not (Status == 'Attenuation' or Status == 'Hypokinesia'):
                    if checkbox.text() == 'Ant':
                        self.ui.Ant.setCurrentText(Status)
                    if checkbox.text() == 'AL' or checkbox.text() == 'Lat' or checkbox.text() == 'IL':
                        self.ui.Lat.setCurrentText(Status)
                    if checkbox.text() == 'Inf':
                        self.ui.Inf.setCurrentText(Status)
                    if checkbox.text() == 'AS' or checkbox.text() == 'Sep' or checkbox.text() == 'IS':
                        self.ui.Sep.setCurrentText(Status)
                    if checkbox.text() == 'AP':
                        self.ui.AP.setCurrentText(Status)

        if not (Status == 'Attenuation' or Status == 'Hypokinesia'):
            for key in Wallstatus:
                if Wallstatus[key]:
                    if key == "AP":
                        TLesion.append(WallDic[key])
                    else:
                        TLesion.append(Deep + WallDic[key])
            if len(TLesion) > 2:
                Wischemia = ', '.join(TLesion[:len(TLesion) - 1])
                Wischemia = Wischemia + ' and ' + TLesion[len(TLesion) - 1]
                self.AddComm(1, Wischemia.capitalize() + " wall ischemia.")
            elif len(TLesion):
                Wischemia = ' and '.join(TLesion)
                self.AddComm(1, Wischemia.capitalize() + " wall ischemia.")
            else:
                Wischemia = "Normal myocardial perfusion."
                self.AddComm(1, Wischemia)
        #print(len(TLesion))
        elif Status == 'Hypokinesia':
            tLesion = []
            for key in Wallstatus:
                if Wallstatus[key]:
                    tLesion.append(WallDic[key])
            if len(tLesion) > 2:
                Wischemia = ', '.join(tLesion[:len(tLesion) - 1])
                Wischemia = Wischemia + ' and ' + tLesion[len(tLesion) - 1]
                self.AddComm(0, Wischemia.capitalize() + " wall hypokinesia.")
            elif len(tLesion):
                Wischemia = ' and '.join(tLesion)
                self.AddComm(0, Wischemia.capitalize() + " wall hypokinesia.")
            else:
                self.AddComm(0, "Global LV wall hypokinesia.")
        else:
            tLesion = []
            for key in Wallstatus:
                if Wallstatus[key]:
                    tLesion.append(WallDic[key])
            if len(tLesion) > 2:
                Wischemia = ', '.join(tLesion[:len(tLesion) - 1])
                Wischemia = Wischemia + ' and ' + tLesion[len(tLesion) - 1]
                self.AddComm(0, "Decreased tracer uptake in the " + Wischemia.capitalize() + " wall is likely due to attenuation artifact.")
            elif len(tLesion):
                Wischemia = ' and '.join(tLesion)
                self.AddComm(0, "Decreased tracer uptake in the " + Wischemia.capitalize() + " wall is likely due to attenuation artifact.")

        self.ui.Partial.setChecked(1)
        return
            
        #print(TLesion[len(TLesion)-1])
        
    def CleanL(self):
        TLesion.clear()
        #print("should clean")
        for combo in self.ui.groupStatus_2.findChildren(QtWidgets.QComboBox):
            combo.setCurrentIndex(0)
        self.ui.Comm.setText("")
        return
    
    def Cleana(self):
        TLesion.clear()
        for combo in self.ui.groupStatus_2.findChildren(QtWidgets.QComboBox):
            combo.setCurrentIndex(0)
        for line in self.ui.groupFun.findChildren(QtWidgets.QLineEdit):
            line.clear()
        self.ui.LVD.setCheckState(0)
        self.ui.Comm.setText("")
        
        return


    def PresseEnter(self,i):
        IndexLine = 0
        for line in self.ui.groupFun.findChildren(QtWidgets.QLineEdit):
            if len(line.text()) > 3 and not line.text().count("."):
                if line.objectName() == 'SEF' or line.objectName() == 'REF':
                    t = line.text()
                    self.ui.SEF.setText(t[:2])
                    self.ui.REF.setText(t[2:])
                    self.ui.SLH.setFocus()
                elif line.objectName() == 'SLH' or line.objectName() == 'RLH':
                    t = line.text()
                    self.ui.SLH.setText("0." + t[:2])
                    self.ui.RLH.setText("0." + t[2:])
                    self.ui.Ext.setFocus()
            elif not (len(line.text()) + IndexLine):
                IndexLine = 1
                line.setFocus()
        return

    def Sender(self, text):
        r = ""
        v = ""
        if self.ui.RBox.currentIndex():
            r = self.ui.RBox.currentText()
        if self.ui.VBox.currentIndex():
            v = self.ui.VBox.currentText()
        w_handle = find_windows(title='EBM UniReport') # find the EBM if exist or not
        if w_handle:
            app = Application().connect(title_re="EBM UniReport")
            title = str(app.TMainForm.element_info)
            text = text.replace('\r','')
            Title = title.split()
            if 'ID' in Title:
                ctext = text + 'IamSep' + r + 'IamSep' + v
                authotkey_process = subprocess.Popen(["sender.exe", "*"], shell=False,\
                                                     stdin=subprocess.PIPE, stdout=subprocess.PIPE,\
                                                     stderr=subprocess.STDOUT, encoding='cp950')
                authotkey_process.communicate(ctext)
                authotkey_process.kill()
                #Self.SetClipboard(ctext)
                #os.system("sender.exe")
            else:
                self.SetClipboard(text)
                self.SetStatusBar('red', 'EBM report is not in EIDT window')
        else:
            self.SetClipboard(text)
            self.SetStatusBar('red', 'The report is in the clipboard')
        return
                
    def AddFunction(self):
        FunDic = {}
        for line in self.ui.groupFun.findChildren(QtWidgets.QLineEdit):
            FunDic[line.objectName()] = line.text()

        if FunDic['SEF']:
            if FunDic['REF']:
                if int(FunDic['SEF']) > 40:
                    self.AddComm(0, "Stress LVEF:{}%, rest LVEF:{}%.".format(FunDic['SEF'],FunDic['REF']))
                else:
                    self.AddComm(0, "Impaired LV contractility, Stress LVEF:{}%, rest LVEF:{}%.".format(FunDic['SEF'], FunDic['REF']))
            elif int(FunDic['SEF']) > 40:
                self.AddComm(0, "Stress LVEF:{}%.".format(FunDic['SEF']))
            else:
                self.AddComm(0, "Impaired LV contractility, Stress LVEF:{}%.".format(FunDic['SEF']))
        if FunDic['SLH'] and float(FunDic['SLH']) > 0.5:
            self.AddComm(0, "Lung congestion.")
        #elif FunDic['RLH'] and float(FunDic['RLH']) > 0.5:
        #    self.AddComm(0, "Lung congestion.")

        if FunDic['Ext'] and float(FunDic['Ext']) >= 10:
            self.AddComm(0, "The total ischemia extent: {}%.".format(FunDic['Ext']))

        if FunDic["Tid"]:
            self.AddComm(0, "Transient ischemic dilatation of LV after stress(TID={}).".format(FunDic['Tid']))

        if self.ui.LVD.isChecked():
            self.AddComm(0, "Cardiomegaly with LV dilatation.")
        return

    def SendF(self):
        PatInfo =[]
        PatCc = []
        lStatus = {'Ant':'Normal uptake', 'Sep':'Normal uptake', 'Inf':'Normal uptake', \
                   'Lat':'Normal uptake', 'AP':'Normal uptake'}
        TracerT = "Tl-201"
        Tracer = "Tl201"
        patInCc = ""
        lungC = 'normal'
        for checkbox in self.ui.patInfo.findChildren(QtWidgets.QCheckBox):
            if checkbox.isChecked():
                if checkbox.text() == "CAD s/p PTCA" and self.ui.PTCA_year.text():
                    PatInfo.append("CAD s/p PTCA {} years ago".format(self.ui.PTCA_year.text()))
                else:
                    PatInfo.append(checkbox.text())
                checkbox.setCheckState(0)
        for checkbox in self.ui.patCc.findChildren(QtWidgets.QCheckBox):
            if checkbox.isChecked():
                PatCc.append(checkbox.text())
                checkbox.setCheckState(0)
            
        #加入病人的info
        if PatInfo:
            if PatCc:
                patInCc = "The patient has history of {} and suffer from {} recently.\r\n\r\n".format(', '.join(PatInfo), ', '.join(PatCc))
            else:
                patInCc = "The patient has history of {}.\r\n\r\n".format(', '.join(PatInfo))
        elif PatCc:
            patInCc = "The patient suffer from {} recently.\r\n\r\n".format(', '.join(PatCc))
        else:
            patInCc = ""
        #for radiobox in self.ui.groupFun.findChildren(QtWidgets.QRadioButton):
        #    if radiobox.isChecked():
        #        Status = radiobox.text()

        if self.ui.Tl201.isChecked():
            TracerT="Tl-201"
            Tracer="Tl201"
            self.ui.Tl201.setAutoExclusive(0)
            self.ui.Tl201.setChecked(0)
            self.ui.Tl201.setAutoExclusive(1)
        elif self.ui.MiBi.isChecked():
            TracerT="Tc99m-MiBi"
            Tracer="MiBi"
            self.ui.MiBi.setAutoExclusive(0)
            self.ui.MiBi.setChecked(0)
            self.ui.MiBi.setAutoExclusive(1)
        else:
            self.SetStatusBar('red', '未指定tracer，預設為 Tl-201')
            TracerT="Tl-201"
            Tracer="Tl201"

        if self.ui.SLH.text():
            if float(self.ui.SLH.text()) > 0.5:
                lungC = "abnormal"
            else:
                lungC = "  normal"

        for combo in self.ui.groupStatus_2.findChildren(QtWidgets.QComboBox):
            if combo.currentIndex():
                lStatus[combo.objectName()] = "Abnormal uptake, {}.".format(combo.currentText().lower())
            else:
                lStatus[combo.objectName()] = "Normal uptake."

        temp = self.ui.Comm.toPlainText()
        if temp:
            Comm = temp.splitlines()
            #prepare for the report
            Text = MainText.format(tracerT=TracerT, Patient=patInCc, TracerProc=procd[Tracer],  Antstatus=lStatus['Ant'], \
                              Sepstatus=lStatus['Sep'], Infstatus=lStatus['Inf'], \
                              Latstatus=lStatus['Lat'], Apstatus=lStatus['AP'], \
                              LungC=lungC, SLH=self.ui.SLH.text(), RLH=self.ui.RLH.text(), CoMm=os.linesep.join(Comm))
            Text = os.linesep.join(Text.split('\n'))
            self.Sender(Text)
        else:
            self.SetStatusBar('red', 'Comm is empty')

        self.Cleana()
        return

    def mbfqSend(self):
        mbfqData = {}
        for line in self.ui.tracerDose.findChildren(QtWidgets.QLineEdit):
            mbfqData[line.objectName()] = line.text()
            line.clear()
        for line in self.ui.stressVitalSign.findChildren(QtWidgets.QLineEdit):
            mbfqData[line.objectName()] = line.text()
            line.clear()
        for box in self.ui.qcBox.findChildren(QtWidgets.QComboBox):
            mbfqData[box.objectName()] = box.currentText()
            box.setCurrentText("Excellent")
        for line in self.ui.flowStatusBox.findChildren(QtWidgets.QLineEdit):
            mbfqData[line.objectName()] = line.text()
            line.clear()
        for box in self.ui.flowStatusBox.findChildren(QtWidgets.QComboBox):
            mbfqData[box.objectName()] = box.currentText()
            mbfqData[box.objectName() + 'index'] = box.currentIndex()
            box.setCurrentText("Normal")
        patsymp = []
        for checkbox in self.ui.qcBox.findChildren(QtWidgets.QCheckBox):
            if checkbox.isChecked():
                patsymp.append(checkbox.text().lower())
                checkbox.setCheckState(0)

        if patsymp:
            print(patsymp)
            mbfqData['patsym'] = "During stressed, the patient developed {}.".format(", ".join(patsymp))
        else:
            print(patsymp)
            mbfqData['patsym'] = ""

        mbfqData['MDFStatus'] = self.ui.MDFStatus.currentText().lower()
        #self.ui.MDFStatus.setCurrentText("Normal")
        mbfqData['MDFArea'] = self.ui.MDFArea.text()
        #self.ui.MDFArea.clear()

        mbfqData['RMibiDose'] = float(mbfqData['RMibiDose']) - float(mbfqData['RRMibiDose'])
        mbfqData['SMibiDose'] = float(mbfqData['SMibiDose']) - float(mbfqData['RSMibiDose'])

        n = int(self.ui.MDFStatus.currentIndex())
        statusList = [['0','NL','NLA'], ['1','Miab','MiabA'], ['2','Moab','MoabA'], ['3','Isch','IschA'], ['4','Isch','IschA'], ['5','Steal','SteA'], ['6','Infar','InfA']]
        chineseList = ['LAD','LCX','RCA','LV']
        commChinese = ['血流狀態在正常範圍', '血流區域有 {}% 為輕微異常', '血流區域有 {}% 為中度異常', '血流區域有 {}% 為缺血', '血流區域有 {}% 為嚴重缺血', '血流區域有 {}% 為血流靜滯', '血流區域有 {}% 為心肌梗塞']

        for list in statusList:
            if int(list[0]) == n:
                mbfqData[list[1]] = "V"
                mbfqData[list[2]] = self.ui.MDFArea.text() + "%"
            elif not list[1] in mbfqData:
                mbfqData[list[1]] = ""
                mbfqData[list[2]] = ""

        for list in chineseList:
            mbfqData[list + 'chinese'] = commChinese[mbfqData[list + "Statusindex"]].format(mbfqData[list + "Area"])

        Text = MBFQ_Text.format(mbfqData)
        self.ui.MDFStatus.setCurrentText("Normal")
        self.ui.MDFArea.clear()
        self.Sender(Text)
        return


if __name__ == '__main__':


    global WallDic
    global TLesion
    WallDic = {'Ant':'anterior','AL':'anterolateral','Lat':'lateral','IL':'inferolateral','Inf':'inferior',\
               'IS':'inferoseptal','Sep':'septal','AS':'anteroseptal','AP':'apical'}
    TLesion = []
    #DefaultTracer = ["TL - 201"]
    procd = {}
    config = configparser.ConfigParser()
    config.read('./setting/setup.ini', encoding='utf-8-sig')
    procd['Tl201'] = config['proc']['Tl201']
    procd['MiBi'] = config['proc']['MiBi']

    global MainText
    try:
        f = open('./setting/MainText.txt', 'r', encoding='utf-8-sig')
    except UnicodeDecodeError:
        f = open('./setting/MainText.txt,' 'r', encoding='cp950')
    else:
        MainText = os.linesep.join(f.read().split('\n'))
        f.close()
    global MBFQ_Text
    try:
        f = open('./setting/MBFQText.txt', 'r', encoding='utf-8-sig')
    except UnicodeDecodeError:
        f = open('./setting/MBFQText.txt,' 'r', encoding='cp950')
    else:
        MBFQ_Text = os.linesep.join(f.read().split('\n'))
        f.close()

    app = QtWidgets.QApplication(sys.argv)
    myapp = MyWin()
    myapp.show()
    sys.exit(app.exec_())