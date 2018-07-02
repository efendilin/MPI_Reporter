
import sys
import win32clipboard
import win32con
import os

from PyQt5 import QtCore, QtWidgets
from mpi import Ui_MainWindow
from pywinauto.application import Application
from pywinauto.findwindows import find_windows

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

    def SignF(self):
        w_handle = find_windows(title='EBM UniReport') # find the EBM if exist or not
        if not w_handle:
            self.SetStatusBar('red', 'EMB report is not found, 已拷貝到剪貼簿')
        else:
            # print("I find the EBM Report!!!")
            app = Application().connect(title_re="EBM UniReport")
            title = str(app.TMainForm.element_info)
            Title = title.split()
            if 'ID' in Title:
                if self.ui.RBox.currentIndex():
                    app.TMainForm.Edit4.SetText(self.ui.RBox.currentText())
                if self.ui.VBox.currentIndex():
                    app.TMainForm.Edit3.SetText(self.ui.VBox.currentText())
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
            win32clipboard.OpenClipboard()
            Text = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
            win32clipboard.CloseClipboard()
        else:
            # print("I find the EBM Report!!!")
            app = Application().connect(title_re="EBM UniReport")
            title = str(app.TMainForm.element_info)
            Title = title.split()
            if 'ID' in Title:
                Text = app.TMainForm.Edit0.WindowText()
            else:
                self.SetStatusBar('red', 'EBM report is not in EIDT window')

        if Text.find("NUCLEAR MEDICINE REPORT") != -1:
            Text = Text.split('NUCLEAR CARDIOLOGY DIAGNOSIS :\r\n')
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
        w_handle = find_windows(title='EBM UniReport') # find the EBM if exist or not
        if not w_handle:
            self.SetStatusBar('red', 'EMB report is not found, 已拷貝到剪貼簿')
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT,text)
            win32clipboard.CloseClipboard()
        else:
            # print("I find the EBM Report!!!")
            app = Application().connect(title_re="EBM UniReport")
            title = str(app.TMainForm.element_info)
            Title = title.split()
            if 'ID' in Title:
                app.TMainForm.Edit0.SetText(text)
                if self.ui.RBox.currentIndex():
                    app.TMainForm.Edit4.SetText(self.ui.RBox.currentText())
                if self.ui.VBox.currentIndex():
                    app.TMainForm.Edit3.SetText(self.ui.VBox.currentText())     
                    
                # self.SetStatusBar('black', '報告已貼上')
            else:
                self.SetStatusBar('red', 'EBM report is not in EIDT window')
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

    procd['Tl201'] = '''\
Tl-201 myocardial perfusion scan was done with dipyridamole stress. Dipyridamole 0.57mg/kg was IV injected for 5 minutes. \
Tl-201 2mCi was injected 1 minutes after pharmacological stress. Images were started 5 min and 4 hrs after injection with SPECT.\
Images were reconstructed with iterative reconstruction technique and displayed in multiple transaxial planes and 3D pictures.\r\n\
    '''
    procd['MiBi'] = '''\
Tc-99m sestamibi myocardial perfusion scan was done with dipyridamole stress. Dipyridamole 0.57 mg/kg was IV injected for 5 minutes.\
Tc-99m sestamibi was injected after pharmacological stress and at rest, 10mCi and 25mCi respectively. Images were started 60 min after \
1st and 2nd injection with SPECT. Images were reconstructed with iterative reconstruction technique and displayed in multiple transaxial\
 planes and 3D pictures.\r\n\
    '''

    MainText = """\
                 NUCLEAR MEDICINE REPORT\r\n\
              ===========================\r\n\
         {tracerT} MYOCARDIAL PERFUSION SCAN\r\n\
    \r\n\
    {Patient}\
    {TracerProc}\
\r\n\
Analysis of stress/ redistribution images show :\r\n\
  1. Anterior wall : {Antstatus}\r\n\
  2. Septal wall   : {Sepstatus}\r\n\
  3. Inferior wall : {Infstatus}\r\n\
  4. Lateral wall  : {Latstatus}\r\n\
  5. Apical wall   : {Apstatus}\r\n\
\r\n\
The lung uptake was {LungC} L/H ratio was {SLH} (stress)\r\n\
                             L/H ratio was {RLH} (rest)\r\n\
    \r\n\
NUCLEAR CARDIOLOGY DIAGNOSIS :\r\n\
{CoMm}\r\n\
        """

    MBFQ_Text = """\
                        NUCLEAR MEDICINE REPORT\r\n\
             ===================================\r\n\
           Myocardial Blood Flow Quantitation (MBFQ)\r\n\
  \r\n\
檢查流程：Tc-99m-MIBI Rest/Stress DySPECT \r\n\
{0[RMibiDose]:4.2f} mCi of Tc-99m-sestamibi (MIBI) was administrated via IV injection at rest after the patient drank 300cc of water.\
The rest listmode dynamic SPECT (DySPECT) was started 10 sec prior to the MIBI injection and continued for 10 minutes.\
After 2.5 hours,  {0[SMibiDose]:4.2f} mCi of MIBI was reinjected via IV injection at dipyridamole-stress after the patient drank another 300cc of water.\
The stress listmode DySPECT was performed 10 sec prior to the stress MIBI injection and continued for another 10 minutes.\r\n\
 \r\n\
藥物負荷流程\r\n\
The patient was infused with {0[DipDose]:s} mg of dipyridamole for 4 minutes and achieved peak hyperemia at the 7th minutes. \
Baseline heart rate was  {0[BHR]:s} bpm and increased to  {0[AHR]:s} bpm at the peak hyperemia. Baseline blood pressure\
 was  {0[BBP]:s} mmHg and increased to  {0[ABP]:s} mmHg at the peak hyperemia, \
which is normal response to pharmacological stress. {0[patsym]:}\r\n\
\r\n\
檢查結果\r\n\
Study Quality: Stress: {0[SQC]:s}, Rest: {0[RQC]:s}\r\n\
Technical Issues: {0[TechQC]:s}\r\n\
Degraded Flow Status (Extent):\r\n\
LAD={0[LADStatus]:s}({0[LADArea]:s}%)\r\n\
LCX={0[LCXStatus]:s}({0[LCXArea]:s}%)\r\n\
RCA={0[RCAStatus]:s}({0[RCAArea]:s}%)\r\n\
LV= {0[LVStatus]:s}({0[LVArea]:s}%)\r\n\
\r\n\
OVERALL IMPRESSION:\r\n\
The most degraded flow status was {0[MDFStatus]:s} that occupied {0[MDFArea]:s}% of the myocardium.\r\n\
+----------+-----+--------+----------+----------+------+\r\n\
|          |     |        | Moderate |   Mild   |Normal|\r\n\
|Infarction|Steal|Ischemia| abnormal | abnormal |Limit |\r\n\
+----------+-----+--------+----------+----------+------+\r\n\
|{0[Infar]:^10}|{0[Steal]:^5}|{0[Isch]:^8}|{0[Moab]:^10}|{0[Miab]:^10}|{0[NL]:^6}|\r\n\
+----------+-----+--------+----------+----------+------+\r\n\
|{0[InfA]:^10}|{0[SteA]:^5}|{0[IschA]:^8}|{0[MoabA]:^10}|{0[MiabA]:^10}|{0[NLA]:^6}|\r\n\
+----------+-----+--------+----------+----------+------+\r\n\
\r\n\


MBFQ血流狀態圖顯示:\r\n\
 1. 左前降支動脈(LAD){0[LADchinese]:s}。\r\n\
 2. 左迴旋支動脈(LCX){0[LCXchinese]:s} 。\r\n\
 3. 右側冠狀動脈(RCA){0[RCAchinese]:s} 。\r\n\
 4. 整個左心室(LV){0[LVchinese]:s} 。 \r\n\
  """


    app = QtWidgets.QApplication(sys.argv)
    myapp = MyWin()
    myapp.show()
    sys.exit(app.exec_())