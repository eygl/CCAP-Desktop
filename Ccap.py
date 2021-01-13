#!/usr/bin/python
# pylint: disable-msg=C0103
# pylint: disable-msg=C0301
"""
PillCap Communication Program
Version Date			Changes
0.10		01-31-14		First version
0.20		02-25-14		Demo
0.30		03-12-14		Prototype release
0.31		03-25-14		Remove ReportLab
							Fixed calculation and spacing of compliance rate in txt report
0.32		04-13-14		Enable use without Basestation or com port
							Auto-detect Com Port and Base Station if connection is made after software is started
							Added Reading, Erasing and Writing Cap message to status bar, replaces incorrect error message
							Switched from 24 to 12 hour clock
							Version in title of main page
							Doubled the rate at which the software gets status from cap
							Prevented direct exit from 'x' button in title bar from pages other than main page, would cause lockup
0.40		08-04-14		Removed popup are you sure boxes
							Added patient, client and treatment data lists
							Added read cap current time
							Added buzzer on/off checkbox
							updated status info from cap
							Added client, patient and data database
0.50		09-16-14		Sorting of Client and Patient names in text edit and drop down boxes
							Sorting out of punctuation and capitalization with database
							Issue com timer off to cap after write complete
							Buzzer on by default, save client's previous choice
0.51		02-21-15		Removed Windows Sorting of Client and Patient names in text edit and drop down boxes
0.52		07-18-15		Detect if OS is Windows to prevent name lookup errors
							Saving previous value of doctor and treatment in write window
							Database location selectable from Settings page
							Settings saved as soon as they are changed, was on program exit
							Immediate saving and displaying of data on cap read
							Fixed date being 1 day early on OSX write page
							Removed 3 and 4/day option, added every other day and every three days
							Added battery life progress bar in status bars
							Fixed error if writing cap, no client and no patient selected and buzzer selected
							Renamed to Ccap
							Fixed Windows GUI problem on write page, couldn't change minutes, AM/PM hidden
0.60		09-14-15		Fixed issue with every other day and every three day display
							Added more dose frequency options, modify output frequency units according to dose frequency selection
							Added reporting dose length to report
							Added buttons to reset battry timer, powerdown cap and enter demo mode
Above Developed by Jeffery Soohoo
Email as of 2021:

0.61        01-13-21

Aboved Developed by Erick Gonzalez
Email as of 2021: erickygonzalez@gmail.com
"""
SWVer = '0.60'

import wx
import wx.adv # Calendar
import sys
import platform
import os.path
import time
from time import sleep
import datetime
import xml.etree.ElementTree as ET
import xml.etree.cElementTree as ET
import Reports
import CapIO
import configparser
import glob

import qrcode
import json
from PIL import Image, ImageDraw, ImageFont


MAX_STRING_LENGTH = 20
MAX_BATTERY_LIFE = 365

GlobalSerialPort = ''
GlobalDBLocation = ''


DosePatternsText = {'2/day': 'twice a day', '1/day': 'once a day', '1/2 days': 'every other day', '1/3 days': 'once every three days', '1/week': 'once a week', '1/month': 'once a month',
                   '1/3 months': 'once every 3 months'}
DosePatterns = {'12': '2/day', '24': '1/day', '48': '1/2 days', '72': '1/3 days', '168': '1/week', '720': '1/month',
                '2160': '1/3 months'}
DosePatternsRev = {'2/day': '12', '1/day': '24', '1/2 days': '48', '1/3 days': '72', '1/week': '168', '1/month': '720',
                   '1/3 months': '2160'}
CapStatus = {'0': 'Unprogrammed', '1': 'Programmed', '3': 'Running', '7': 'Done', '8': 'Unprogrammed Demo',
             '9': 'Programmed Demo', ';': 'Running Demo', '?': 'Done Demo'}


def SaveConfigFile(ConfigList):
    # open config file with parser
    parser = configparser.ConfigParser()
    parser.read('Ccap.cfg')

    # save settings
    for Setting in ConfigList:
        parser.set('Settings', Setting, ConfigList[Setting])

    # Write configuration file to 'Ccap.cfg'
    with open('Ccap.cfg', 'wb') as configfile:
        parser.write(configfile)


def XMLTagTextPrepare(Text):
    ReturnText = Text.lower()
    if Text:
        ReturnText = ' '.join([s[0].upper() + s[1:] for s in ReturnText.split(' ')])
    ReturnText = ReturnText.replace(' ', '_')
    ReturnText = ReturnText.replace(',', '-')
    return ReturnText


def XMLTagTextExtract(Tag):
    ReturnText = Tag.replace('_', ' ')
    ReturnText = ReturnText.replace('-', ',')
    return ReturnText


def DataBaseClientNameCheck(ClientList, ClientNameSubString):
    ClientSubList = []
    SubStringFound = False
    for Client in ClientList:
        if ClientNameSubString.lower() in Client.lower():
            ClientSubList.append(Client)
            SubStringFound = True
    return (SubStringFound, ClientSubList)


def XMLCharacterCheck(XMLText):
    StringValid = True
    i = 0
    for c in XMLText:
        if ((i == 0) and (c.isalpha() == False)) or \
                ((c.isalnum() == False) and (c.isspace() == False) and (c != '.') and (c != ',')) and (c != '-'):
            StringValid = False
        i += 1
    return StringValid


def XMLIndent(elem, level=0):
    i = "\n" + level * "  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
        for elem in elem:
            XMLIndent(elem, level + 1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i


def CalcFormLength():
    if platform.system() == "Darwin":
        FormLength = 610
    elif platform.system() == "Windows":
        FormLength = 605
    else:
        FormLength = 600
    return FormLength


def CalcFormHeight():
    if platform.system() == "Darwin":
        FormHeight = 430
    elif platform.system() == "Windows":
        FormHeight = 450
    else:
        FormHeight = 424
    return FormHeight




class WriteWindow(wx.Frame):
    CapErased = False

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, title='Write Cap', size=(CalcFormLength(), CalcFormHeight() +150),
                          style=wx.DEFAULT_FRAME_STYLE & ~(wx.RESIZE_BORDER | wx.CLOSE_BOX | wx.MAXIMIZE_BOX))

        wx.Frame.CenterOnScreen(self)
        panel = wx.Panel(self, -1)
        BackGroundColour = (233, 228, 214)
        panel.SetBackgroundColour(BackGroundColour)
        font16 = wx.Font(16, wx.FONTFAMILY_SWISS, wx.NORMAL, wx.NORMAL)
        font12 = wx.Font(12, wx.FONTFAMILY_SWISS, wx.NORMAL, wx.NORMAL)
        font10 = wx.Font(10, wx.FONTFAMILY_SWISS, wx.NORMAL, wx.NORMAL)

        # Status bar
        self.StatusBar = wx.StatusBar(self, -1)
        self.StatusBar.SetFieldsCount(5)

        # battery life progress bar in status bar
        self.ProgressBar = wx.Gauge(self.StatusBar, -1, style=wx.GA_HORIZONTAL | wx.GA_SMOOTH)
        rect = self.StatusBar.GetFieldRect(4)
        self.ProgressBar.SetPosition((rect.x + 2, rect.y + 2))
        self.ProgressBar.SetSize((rect.width - 4, rect.height - 4))

        self.WriteFacilityText = wx.StaticText(panel, -1, u'Facility', pos=(15, 25), size=(120, 30))
        self.WriteFacilityText.SetFont(font16)

        self.WriteFacilityTextCtrl = wx.TextCtrl(panel, -1, '', pos=(150, 25), size=(175, -1))
        self.WriteFacilityTextCtrl.SetFont(font10)
        self.WriteFacilityTextCtrl.SetMaxLength(MAX_STRING_LENGTH)

        self.WriteDrText = wx.StaticText(panel, -1, u'Doctor', pos=(15, 50), size=(120, 30))
        self.WriteDrText.SetFont(font16)
        self.WriteDrChoice = wx.ComboBox(panel, -1, pos=(150, 50), size=(175, 25), style=wx.CB_DROPDOWN)
        self.WriteDrChoice.SetFont(font12)

        self.WriteTreatmentText = wx.StaticText(panel, -1, u'Treatment', pos=(15, 75), size=(120, 30))
        self.WriteTreatmentText.SetFont(font16)

        self.WriteTreatmentChoice = wx.ComboBox(panel, -1, pos=(150, 75), size=(175, 25), style=wx.CB_DROPDOWN)
        self.WriteTreatmentChoice.SetFont(font12)

        self.WriteDoseText = wx.StaticText(panel, -1, u'Dose Freq.', pos=(15, 100), size=(120, 30))
        self.WriteDoseText.SetFont(font16)

        self.WriteDoseChoice = wx.ComboBox(panel, -1, '1/day', pos=(150, 100), size=(175, 25),
                                           choices=[u'2/day', u'1/day', u'1/2 days', u'1/3 days', u'1/week', u'1/month',
                                                    u'1/3 months'], style=wx.CB_DROPDOWN | wx.CB_READONLY)
        self.WriteDoseChoice.SetFont(font12)
        self.Bind(wx.EVT_COMBOBOX, self.on_DoseChange, self.WriteDoseChoice)

        self.WriteDoseCountText = wx.StaticText(panel, -1, u'Dose Count', pos=(15, 125), size=(120, 30))
        self.WriteDoseCountText.SetFont(font16)

        self.WriteDoseCount = wx.SpinCtrl(panel, -1, pos=(152, 125), size=(175, 25), initial=30, value='30', min=1,
                                          max=120)
        self.WriteDoseCount.SetFont(font12)
        self.Bind(wx.EVT_SPINCTRL, self.on_DoseChange, self.WriteDoseCount)

        self.WriteTreatmentLengthOutText = wx.StaticText(panel, -1, u'30 Days', pos=(350, 125), size=(120, 30))
        self.WriteTreatmentLengthOutText.SetFont(font16)

        self.WriteClientText = wx.StaticText(panel, -1, u'Client', pos=(15, 150), size=(120, 30))
        self.WriteClientText.SetFont(font16)

        self.WriteClientChoice = wx.ComboBox(panel, -1, pos=(150, 150), size=(175, -1), style=wx.CB_DROPDOWN)
        self.WriteClientChoice.SetFont(font12)
        self.Bind(wx.EVT_COMBOBOX, self.on_ClientChange, self.WriteClientChoice)
        #		self.Bind(wx.EVT_TEXT, self.on_ClientTextChange, self.WriteClientChoice)

        self.WritePatientText = wx.StaticText(panel, -1, u'Patient', pos=(15, 175), size=(120, 30))
        self.WritePatientText.SetFont(font16)

        self.WritePatientChoice = wx.ComboBox(panel, -1, pos=(150, 175), size=(175, -1), style=wx.CB_DROPDOWN)
        self.WritePatientChoice.SetFont(font12)
        # self.WritePatientTextCtrl.SetMaxLength(MAX_STRING_LENGTH)
        #		self.Bind(wx.EVT_TEXT, self.on_PatientTextChange, self.WritePatientChoice)

        self.WriteRouteText = wx.StaticText(panel, -1, u'Route', pos=(15, 200), size=(120, 30))
        self.WriteRouteText.SetFont(font16)

        self.WriteRouteChoice = wx.ComboBox(panel, -1, '1/day', pos=(150, 200), size=(175, 25),
                                           choices=[u'Mouth',u'Mouth with food'], style=wx.CB_DROPDOWN | wx.CB_READONLY)
        self.WriteRouteChoice.SetFont(font12)
        self.Bind(wx.EVT_COMBOBOX, self.on_RouteChange, self.WriteRouteChoice)


        self.WriteStartDateText = wx.StaticText(panel, -1, u'Start Date', pos=(15, 240), size=(120, 30))
        self.WriteStartDateText.SetFont(font16)

        self.Calendar = wx.adv.CalendarCtrl(panel, 10, wx.DateTime.Now(),pos=(150,240))
        #self.Calendar.SetDateRange(lowerdate=wx.DateTime.Now())
        self.Calendar.Bind(wx.adv.EVT_CALENDAR_SEL_CHANGED, self.OnDate)

        self.WriteStartTimeText = wx.StaticText(panel, -1, u'Start Time', pos=(15, 250+150), size=(120, 30))
        self.WriteStartTimeText.SetFont(font16)
        self.WriteColonText = wx.StaticText(panel, -1, u':', pos=(200, 250+150), size=(15, 30))
        self.WriteColonText.SetFont(font16)

        now = datetime.datetime.now()
        if now.hour < 12:
            InitHour = now.hour + 1
            AMOrPM = u'AM'
        else:
            InitHour = now.hour - 11
            AMOrPM = u'PM'
        self.WriteStartHours = wx.SpinCtrl(panel, -1, pos=(150, 250+150), size=(50, 25), initial=InitHour,
                                           value=str(InitHour), min=1, max=12)
        self.WriteStartHours.SetFont(font12)
        self.WriteStartMins = wx.SpinCtrl(panel, -1, pos=(215, 250+150), size=(50, 25), initial=0, value='0', min=0, max=59)
        self.WriteStartMins.SetFont(font12)

        self.WriteAMPMChoice = wx.ComboBox(panel, -1, AMOrPM, pos=(270, 250+150), size=(60, 25),
                                           choices=[u'AM', u'PM'], style=wx.CB_DROPDOWN | wx.CB_READONLY)
        self.WriteAMPMChoice.SetFont(font12)

        self.WriteBuzzerEnableCheckBox = wx.CheckBox(panel, -1, 'Enable Buzzer', (150, 280+150), size=(150, 30))
        self.WriteBuzzerEnableCheckBox.SetFont(font12)

        self.WriteNewBatteryButton = wx.Button(panel, -1, label='New Battery', pos=(300, 300+150), size=(150, 50))
        self.WriteNewBatteryButton.SetFont(font16)
        self.Bind(wx.EVT_BUTTON, self.on_WriteNewBatteryButton, self.WriteNewBatteryButton)

        self.WriteDemoButton = wx.Button(panel, -1, label='Demo Mode', pos=(450, 300+150), size=(150, 50))
        self.WriteDemoButton.SetFont(font16)
        self.Bind(wx.EVT_BUTTON, self.on_WriteDemoButton, self.WriteDemoButton)

        self.WriteWriteButton = wx.Button(panel, -1, label='Write', pos=(0, 350+150), size=(150, 50))
        self.WriteWriteButton.SetFont(font16)
        self.Bind(wx.EVT_BUTTON, self.on_WriteWriteButton, self.WriteWriteButton)

        self.WriteEraseButton = wx.Button(panel, -1, label='Erase', pos=(150, 350+150), size=(150, 50))
        self.WriteEraseButton.SetFont(font16)
        self.Bind(wx.EVT_BUTTON, self.on_WriteEraseButton, self.WriteEraseButton)

        self.WriteQRCodeButton = wx.Button(panel, -1, label='QR Code', pos=(300, 350+150), size=(150, 50))
        self.WriteQRCodeButton.SetFont(font16)
        self.Bind(wx.EVT_BUTTON, self.on_WriteQRCodeButton, self.WriteQRCodeButton)

        self.WriteBackButton = wx.Button(panel, -1, label='Back', pos=(450, 350+150), size=(150, 50))
        self.WriteBackButton.SetFont(font16)
        self.Bind(wx.EVT_BUTTON, self.on_WriteBackButton, self.WriteBackButton)

    def on_RouteChange(self,event):
        if self.WriteRouteChoice.GetStringSelection() == 'Mouth':
            print(self.WriteRouteChoice.GetStringSelection())

    def OnDate(self,event):
        print(self.Calendar.GetDate())
        print(type(self.Calendar.GetDate()))

    def on_PatientTextChange(self, event):
        if platform.system() != "Windows":
            EnteredText = self.WritePatientChoice.GetValue()
            if XMLCharacterCheck(EnteredText) == False:
                self.WritePatientChoice.SetValue(EnteredText[:-1])

    def on_ClientTextChange(self, event):
        if platform.system() != "Windows":
            global GlobalDBLocation
            EnteredText = self.WriteClientChoice.GetValue()
            if XMLCharacterCheck(EnteredText) == False:
                self.WriteClientChoice.SetValue(EnteredText[:-1])

            ClientDB = ET.parse(os.path.join(GlobalDBLocation, 'PatientDB.XML'))
            ClientList = []
            for Clients in ClientDB.getroot():
                ClientList.append(XMLTagTextExtract(Clients.tag))

            ClientCheck = DataBaseClientNameCheck(ClientList, EnteredText)
            if ClientCheck[0] == True:
                self.WriteClientChoice.Clear()
                for Client in sorted(ClientCheck[1], key=lambda s: s.lower()):
                    self.WriteClientChoice.Append(Client)
            else:
                self.WriteClientChoice.Clear()

            if self.WriteClientChoice.GetValue() in ClientList \
                    or len(self.WriteClientChoice.GetValue()) == 0:
                self.on_ClientChange(self.on_ClientChange)

    def on_ClientChange(self, event):
        # Fill Client list and select first patient
        global GlobalDBLocation
        ClientDB = ET.parse(os.path.join(GlobalDBLocation, 'PatientDB.XML'))
        ClientElementList = ClientDB.getroot()
        SelectedClient = ClientElementList.find(XMLTagTextPrepare(self.WriteClientChoice.GetValue()))
        self.WritePatientChoice.Clear()
        # get patient list and sort alphabetically
        if SelectedClient:
            PatientList = []
            for Patients in SelectedClient:
                PatientList.append(XMLTagTextExtract(Patients.tag))
            for Patient in sorted(PatientList, key=lambda s: s.lower()):
                self.WritePatientChoice.Append(Patient)

            self.WritePatientChoice.SetSelection(0)
            if 'Buzzer' not in SelectedClient.attrib:
                self.WriteBuzzerEnableCheckBox.SetValue(True)
            elif SelectedClient.attrib['Buzzer'] == 'OFF':
                self.WriteBuzzerEnableCheckBox.SetValue(False)
            else:
                self.WriteBuzzerEnableCheckBox.SetValue(True)
        else:
            self.WritePatientChoice.Clear()
            self.WritePatientChoice.SetSelection(0)
            self.WriteBuzzerEnableCheckBox.SetValue(True)

    def on_DoseChange(self, event):
        if self.WriteDoseChoice.GetSelection() < 4:
            TreatmentLength = self.WriteDoseCount.GetValue() * int(
                DosePatternsRev[self.WriteDoseChoice.GetValue()]) / 24
            if self.WriteDoseCount.GetValue() % 2 != 0 and self.WriteDoseChoice.GetSelection() == 0:
                TreatmentLength += 1
            self.WriteTreatmentLengthOutText.SetLabel(str(TreatmentLength) + ' Days')
        elif self.WriteDoseChoice.GetSelection() == 4:
            TreatmentLength = self.WriteDoseCount.GetValue() * int(
                DosePatternsRev[self.WriteDoseChoice.GetValue()]) / 168
            self.WriteTreatmentLengthOutText.SetLabel(str(TreatmentLength) + ' Weeks')
        else:
            TreatmentLength = self.WriteDoseCount.GetValue() * int(
                DosePatternsRev[self.WriteDoseChoice.GetValue()]) / 720
            self.WriteTreatmentLengthOutText.SetLabel(str(TreatmentLength) + ' Months')


    def isWriteFormFilled(self):
        if (self.WriteFacilityTextCtrl.GetValue() == '' or
            self.WriteDrChoice.GetValue() == '' or
            str(self.WriteDoseCount.GetValue()) == '' or
            self.WriteClientChoice.GetValue() == '' or
            self.WritePatientChoice.GetValue() == '' or
            self.WriteRouteChoice.GetValue() == ''):
            return False
        else:
            return True

    def on_WriteWriteButton(self, event):
        global GlobalSerialPort
        global GlobalDBLocation
        MaxBatteryLife = 365

        
        if GlobalSerialPort == CapIO.NO_COM_PORT_TEXT:
            dlg = wx.MessageDialog(None, 'No Com Port', 'Error', wx.OK)
            DialogReturn = dlg.ShowModal()
        elif str(GlobalSerialPort).find(CapIO.NO_BASE_STATION_TEXT) == 0:
            dlg = wx.MessageDialog(None, 'No BaseStation', 'Error', wx.OK)
            DialogReturn = dlg.ShowModal()
        elif not self.isWriteFormFilled():
            dlg = wx.MessageDialog(None, 'Some fields are missing information.', 'Error', wx.OK)
            DialogReturn = dlg.ShowModal()
        else:
            WriteDict = {}
            self.StatusBar.SetStatusText('Erasing Cap', 3)

            if CapIO.Erase(GlobalSerialPort):
                self.StatusBar.SetStatusText('Writing Cap', 3)
                # first save buzzer selection to database
                ClientDB = ET.parse(os.path.join(GlobalDBLocation, 'PatientDB.XML'))
                ClientElementList = ClientDB.getroot()
                SelectedClient = ClientElementList.find(XMLTagTextPrepare(self.WriteClientChoice.GetValue()))
                if SelectedClient:
                    if self.WriteBuzzerEnableCheckBox.GetValue() == False:
                        SelectedClient.attrib['Buzzer'] = 'OFF'
                    else:
                        SelectedClient.attrib['Buzzer'] = 'ON'
                XMLIndent(ClientElementList)
                ClientDB.write(os.path.join(GlobalDBLocation, 'PatientDB.XML'))
                # then write to cap
                WriteDict['Facility'] = self.WriteFacilityTextCtrl.GetValue()
                WriteDict['Doctor'] = self.WriteDrChoice.GetValue()
                WriteDict['Treatment'] = self.WriteTreatmentChoice.GetValue()
                WriteDict['DosePattern'] = DosePatternsRev[self.WriteDoseChoice.GetValue()]
                WriteDict['DoseCount'] = str(self.WriteDoseCount.GetValue())
                WriteDict['Client'] = self.WriteClientChoice.GetValue()
                WriteDict['Patient'] = self.WritePatientChoice.GetValue()
                if self.WriteBuzzerEnableCheckBox.GetValue():
                    WriteDict['BuzzerEnable'] = '1'
                else:
                    WriteDict['BuzzerEnable'] = '0'
                BatteryAge = CapIO.ReadBatteryAge(GlobalSerialPort)
                BatteryLife = MaxBatteryLife - BatteryAge
                TreatmentTime = self.WriteDoseCount.GetValue() * int(
                    DosePatternsRev[self.WriteDoseChoice.GetValue()]) / 24
                if BatteryLife < TreatmentTime:
                    MessageString = 'Insufficient Battery Life. Replace Battery\n'
                    wx.MessageBox(MessageString, '')
                else:
                    now = datetime.datetime.now()  # # set clock
                    CapIO.WriteString(GlobalSerialPort,
                                      'B:{:0>2}:{:0>2}:{:0>2}:{:0>2}:{:0>2}:{:0>2}:{:0>2}\r'.format(now.second,
                                                                                                    now.minute,
                                                                                                    now.hour, now.day,
                                                                                                    now.weekday() + 1,
                                                                                                    now.month,
                                                                                                    now.year - 2000))
                    # Get Value from Calendar
                    #TODO
                    StartTime = self.Calendar.GetDate()
                    
                    if self.WriteAMPMChoice.GetValue() == 'AM':
                        StartHours = self.WriteStartHours.GetValue()
                        if StartHours == 12:
                            StartHours = 0
                    else:
                        StartHours = self.WriteStartHours.GetValue() + 12
                        if StartHours == 24:
                            StartHours = 12
                    CapIO.WriteString(GlobalSerialPort, 'F:00:{:0>2}:{:0>2}:{:0>2}:{:0>2}:{:0>2}:{:0>2}\r'.format(
                        self.WriteStartMins.GetValue(), StartHours, StartTime.GetDay(), StartTime.GetWeekDay(),
                        StartTime.GetMonth() + 1, StartTime.GetYear() - 2000))
                    CapIO.WriteSettings(GlobalSerialPort, WriteDict)
                    CapIO.WriteString(GlobalSerialPort, 'J\r')  # start

                    


                    Timer = 10
                    while Timer:
                        Timer = Timer - 1
                        if Timer == 0:
                            dlg = wx.MessageDialog(None, 'Cap Program Timeout.', 'Failure', wx.OK)
                            DialogReturn = dlg.ShowModal()
                            break
                        PillCapTitleList = CapIO.Ping(GlobalSerialPort)
                        if len(PillCapTitleList) >= 3:
                            if PillCapTitleList[3] != '0' and PillCapTitleList[3] != '8':
                                CapIO.WriteString(GlobalSerialPort, 'Q\r')  # clear timer

                                
                                self.on_GenerateQRCode()
                                dlg = wx.MessageDialog(None, 'Cap Programmed.', 'Success', wx.OK)
                                DialogReturn = dlg.ShowModal()
                                break
                        else:
                            sleep(1)
            else:
                dlg = wx.MessageDialog(None, 'Erase Error:\nCheck cap alignment.\nPress cap.\nCheck battery.', 'Error',
                                       wx.OK)
                DialogReturn = dlg.ShowModal()

    def on_WriteNewBatteryButton(self, event):
        global GlobalSerialPort
        if GlobalSerialPort == CapIO.NO_COM_PORT_TEXT:
            dlg = wx.MessageDialog(None, 'No Com Port', 'Error', wx.OK)
            DialogReturn = dlg.ShowModal()
        else:
            self.StatusBar.SetStatusText('Resetting Battery Timer', 3)
            now = datetime.datetime.now()
            CapIO.WriteString(GlobalSerialPort,
                              'B:{:0>2}:{:0>2}:{:0>2}:{:0>2}:{:0>2}:{:0>2}:{:0>2}\r'.format(now.second, now.minute,
                                                                                            now.hour, now.day,
                                                                                            now.weekday() + 1,
                                                                                            now.month, now.year - 2000))
            CapIO.WriteString(GlobalSerialPort, 'N\r')

    def on_WriteEraseButton(self, event):
        global GlobalSerialPort
        if GlobalSerialPort == CapIO.NO_COM_PORT_TEXT:
            dlg = wx.MessageDialog(None, 'No Com Port', 'Error', wx.OK)
            DialogReturn = dlg.ShowModal()
        else:
            self.StatusBar.SetStatusText('Erasing Cap', 3)
            if CapIO.Erase(GlobalSerialPort) == 0:
                dlg = wx.MessageDialog(None, 'Erase Error:\nCheck cap alignment.\nPress cap.\nCheck battery.', 'Error',
                                       wx.OK)
                DialogReturn = dlg.ShowModal()

    def on_GenerateQRCode(self):
        StartTime = datetime.datetime.now().isoformat()

        self.rx = [{
            "treatment": self.WriteTreatmentChoice.GetValue(),
            "dosecount": str(self.WriteDoseCount.GetValue()),
            "dosefreq": DosePatternsText[self.WriteDoseChoice.GetValue()],
            "startdate": StartTime,
            "patient" : self.WritePatientChoice.GetValue(),
            "client" : self.WriteClientChoice.GetValue(),
            "doctor" : self.WriteDrChoice.GetValue(),
            "facility" : self.WriteFacilityTextCtrl.GetValue(),
            "route" :  str(self.WriteRouteChoice.GetValue()).lower(),
            "medication" : self.WriteTreatmentChoice.GetValue(),
            "id": int(round(time.time() * 1000)),
            "note":"N/A",
            "active":True,
        }]

        
        obj = json.dumps(self.rx)
        print(obj)
        qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=6,
                border=4,
        )
        qr.add_data(obj)
        qr.make(fit=True)
        qr_img = qr.make_image(fill='black', back_color='white')
        
        width, height = qr_img.size
        margin = 180
        imageCanvas = Image.new('RGB', (width, height + margin),(255, 255, 255))
        imageCanvas.paste(qr_img)

        draw = ImageDraw.Draw(imageCanvas)             
        boldFont = ImageFont.truetype('Roboto-Bold.ttf',size=18)
        regularFont = ImageFont.truetype('Roboto-Regular.ttf',size=14)
        color = 'rgb(0, 0, 0)' 

        now = datetime.datetime.now()
        nowString =  now.strftime("%m/%d/%Y, %I:%M %p")

        
            
        draw.text((20,height - 20),self.rx[0]['facility'],fill=color,font=boldFont,align='center')
        
        headlinerText = "Date: " + nowString + "   |   Doctor: " + self.rx[0]['doctor']
        draw.text((20,height), headlinerText,fill=color,font=regularFont,align='center')
        
        
        draw.text((20,height + 24), "Patient: ", fill=color, font=regularFont, align='left')
        draw.text((70,height + 20), self.rx[0]['patient'], fill=color, font=boldFont, align='left')

        draw.text((20,height + 40), "Client: ", fill=color, font=regularFont, align='left')
        draw.text((60,height + 36), self.rx[0]['client'], fill=color, font=boldFont, align='left')

        treatmentString = "Give " + self.rx[0]['treatment'] + " by " + self.rx[0]['route'] + " " + self.rx[0]['dosefreq'].lower()


        #Intellegently add newline in between words and adjust bottom text
        spaceIndexList = []
        for char in range(len(treatmentString)):
            if treatmentString[char] == ' ':
                spaceIndexList.append(char)
        offset = 0
        maxCharWidth = 60
        numberOfNewlines = 0
        for i in range(len(spaceIndexList)):
            if spaceIndexList[i] - offset > maxCharWidth :
                treatmentString = treatmentString[:spaceIndexList[i-1]] + '\n' + treatmentString[spaceIndexList[i-1]+1:]
                offset = spaceIndexList[i-1]
                numberOfNewlines += 1

                
        draw.text((20,height + 60), treatmentString.upper(), fill=color,font=regularFont,align='left')
        draw.text((20,height + 100 + 14 * numberOfNewlines), self.rx[0]['medication'], fill=color, font=boldFont,align='left') 
        draw.text((20,height + 120 + 14 * numberOfNewlines), "QTY: " + self.rx[0]['dosecount'], fill=color, font= regularFont, align='left')




        global QRCodeWidth
        global QRCodeHeight
        global QRCodeFileName
        QRCodeWidth = width + 100
        QRCodeHeight = height + margin + 100
        QRCodeFileName = "qrcode.png"
        ##ADD TEXT HERE
        imageCanvas.save(QRCodeFileName)

        QRCodeDisplay = QRCodeWindow()

    def on_WriteQRCodeButton(self, event):  
        self.on_GenerateQRCode()

    def on_WriteDemoButton(self, event):
        global GlobalSerialPort
        if GlobalSerialPort == CapIO.NO_COM_PORT_TEXT:
            dlg = wx.MessageDialog(None, 'No Com Port', 'Error', wx.OK)
            DialogReturn = dlg.ShowModal()
        else:
            self.StatusBar.SetStatusText('Enabling Demo Mode', 3)
            CapIO.WriteString(GlobalSerialPort, 'AR/r')

    def on_WriteBackButton(self, event):
        # SettingsDict = {}
        # SettingsDict['lasttreatment'] = self.WriteTreatmentChoice.GetValue()
        # SettingsDict['lastdoctor'] = self.WriteDrChoice.GetValue()
        # SaveConfigFile(SettingsDict)
        self.WriteDoseCount.SetValue(30)
        # self.Calendar.SetRange(wx.DateTime_Now() - wx.DateSpan(days=1), wx.DateTime_Now() + wx.DateSpan(days=90))
        now = datetime.datetime.now()
        if now.hour < 12:
            InitHour = now.hour + 1
            AMOrPM = u'AM'
        else:
            InitHour = now.hour - 11
            AMOrPM = u'PM'
        self.WriteStartHours.SetValue(InitHour)
        self.WriteStartMins.SetValue(0)
        self.WriteAMPMChoice.SetValue(AMOrPM)
        self.WriteBuzzerEnableCheckBox.SetValue(False)
        self.WriteDoseChoice.SetSelection(0)
        # StartTime = self.Calendar.GetValue()
        self.WriteClientChoice.Clear()
        self.Hide()
        frame.Show()

class QRCodeWindow(wx.Frame):
    def __init__(self, parent=None):
        global QRCodeFileName
        global QRCodeWidth
        global QRCodeHeight
        wx.Frame.__init__(self, parent, title='QRCode', size=(QRCodeWidth, QRCodeHeight+50),
                          style=wx.DEFAULT_FRAME_STYLE & ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX))
        png = wx.Image(QRCodeFileName, wx.BITMAP_TYPE_ANY).ConvertToBitmap()
        panel = wx.Panel(self, -1)
        wx.StaticBitmap(panel, -1, png, (0,0), (QRCodeWidth,QRCodeHeight))
        
        self.Show()

class ReportsWindow(wx.Frame):
    FacilityText = ''
    DateTag = ''

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, title='Get Report', size=(CalcFormLength(), CalcFormHeight()),
                          style=wx.DEFAULT_FRAME_STYLE & ~(wx.RESIZE_BORDER | wx.CLOSE_BOX | wx.MAXIMIZE_BOX))
        wx.Frame.CenterOnScreen(self)
        panel = wx.Panel(self, -1)
        BackGroundColour = (233, 228, 214)
        panel.SetBackgroundColour(BackGroundColour)
        font16 = wx.Font(16, wx.FONTFAMILY_SWISS, wx.NORMAL, wx.NORMAL)
        font14 = wx.Font(14, wx.FONTFAMILY_SWISS, wx.NORMAL, wx.NORMAL)
        font12 = wx.Font(12, wx.FONTFAMILY_SWISS, wx.NORMAL, wx.NORMAL)
        font10 = wx.Font(10, wx.FONTFAMILY_SWISS, wx.NORMAL, wx.NORMAL)

        # Status bar
        self.StatusBar = wx.StatusBar(self, -1)
        self.StatusBar.SetFieldsCount(5)

        # battery life progress bar in status bar
        self.ProgressBar = wx.Gauge(self.StatusBar, -1, style=wx.GA_HORIZONTAL | wx.GA_SMOOTH)
        rect = self.StatusBar.GetFieldRect(4)
        self.ProgressBar.SetPosition((rect.x + 2, rect.y + 2))
        self.ProgressBar.SetSize((rect.width - 4, rect.height - 4))

        self.ReportSelectClientText = wx.StaticText(panel, -1, u'Select Client', pos=(15, 50), size=(120, 30))
        self.ReportSelectClientText.SetFont(font16)

        self.ReportClientChoice = wx.ComboBox(panel, -1, pos=(200, 50), size=(175, 25), style=wx.CB_DROPDOWN)
        self.ReportClientChoice.SetFont(font12)
        self.Bind(wx.EVT_COMBOBOX, self.on_ClientChange, self.ReportClientChoice)
        #		self.Bind(wx.EVT_TEXT, self.on_ClientTextChange, self.ReportClientChoice)

        self.ReportSelectPatientText = wx.StaticText(panel, -1, u'Select Patient', pos=(15, 75), size=(120, 30))
        self.ReportSelectPatientText.SetFont(font16)

        self.ReportPatientChoice = wx.ComboBox(panel, -1, pos=(200, 75), size=(175, 25),
                                               style=wx.CB_DROPDOWN | wx.CB_READONLY)
        self.ReportPatientChoice.SetFont(font12)
        self.Bind(wx.EVT_COMBOBOX, self.on_PatientChange, self.ReportPatientChoice)

        self.ReportSelectPatientText = wx.StaticText(panel, -1, u'Select Date', pos=(15, 100), size=(120, 30))
        self.ReportSelectPatientText.SetFont(font16)

        self.ReportDateChoice = wx.ComboBox(panel, -1, pos=(200, 100), size=(350, 25),
                                            style=wx.CB_DROPDOWN | wx.CB_READONLY)
        self.ReportDateChoice.SetFont(font12)

        self.ReportFormatText = wx.StaticText(panel, -1, u'Report Format', pos=(15, 125), size=(120, 30))
        self.ReportFormatText.SetFont(font16)

        self.ReportFormatChoice = wx.ComboBox(panel, -1, 'Text', pos=(200, 125), size=(100, 25),
                                              choices=[u'Text', u'Excel'], style=wx.CB_DROPDOWN | wx.CB_READONLY)
        self.ReportFormatChoice.SetFont(font12)

        self.ReportDestinationText = wx.StaticText(panel, -1, u'Destination', pos = (15, 300), size = (120, 30))
        self.ReportDestinationText.SetFont(font16)

        self.ReportDestination = wx.TextCtrl(panel, -1, u'', pos = (150, 300), size = (450, 25), style = wx.TE_READONLY)
        self.ReportDestination.SetFont(font12)
        self.ReportDestination.Bind(wx.EVT_LEFT_DOWN, self.on_ReportDestination_mouseDown)

        self.ReportsGenerateButton = wx.Button(panel, -1, label='Create Report', pos=(0, 350), size=(150, 50))
        self.ReportsGenerateButton.SetFont(font14)
        self.Bind(wx.EVT_BUTTON, self.on_ReportsGenerateButton, self.ReportsGenerateButton)

        self.ReportsBackButton = wx.Button(panel, -1, label='Back', pos=(450, 350), size=(150, 50))
        self.ReportsBackButton.SetFont(font16)
        self.Bind(wx.EVT_BUTTON, self.on_ReportsBackButton, self.ReportsBackButton)

    def on_ClientTextChange(self, event):
        if platform.system() != "Windows":
            global GlobalDBLocation
            ClientDB = ET.parse(os.path.join(GlobalDBLocation, 'PatientDB.XML'))
            ClientList = []
            for Clients in ClientDB.getroot():
                ClientList.append(XMLTagTextExtract(Clients.tag))

            EnteredText = self.ReportClientChoice.GetValue()
            ClientCheck = DataBaseClientNameCheck(ClientList, EnteredText)
            if ClientCheck[0] == True:
                self.ReportClientChoice.Clear()
                for Client in sorted(ClientCheck[1], key=lambda s: s.lower()):
                    self.ReportClientChoice.Append(Client)
            else:
                self.ReportClientChoice.SetValue(EnteredText[:-1])

            if self.ReportClientChoice.GetValue() in ClientList \
                    or len(self.ReportClientChoice.GetValue()) == 0:
                self.on_ClientChange(self.on_ClientChange)

    def on_ClientChange(self, event):
        # Fill patient list and select first patient
        global GlobalDBLocation
        ClientDB = ET.parse(os.path.join(GlobalDBLocation, 'PatientDB.XML'))
        ClientElementList = ClientDB.getroot()
        PatientElementList = ClientElementList.find(XMLTagTextPrepare(self.ReportClientChoice.GetValue()))
        self.ReportPatientChoice.SetValue('')
        self.ReportPatientChoice.Clear()
        self.ReportDateChoice.SetValue('')
        self.ReportDateChoice.Clear()

        # get patient list and sort alphabetically
        if PatientElementList:
            PatientList = []
            for Patients in PatientElementList:
                PatientList.append(XMLTagTextExtract(Patients.tag))
            for Patient in sorted(PatientList, key=lambda s: s.lower()):
                self.ReportPatientChoice.Append(Patient)
            # print Patient
            # print PatientList
            self.ReportPatientChoice.SetSelection(0)
            self.on_PatientChange(self.on_PatientChange)

    def on_PatientChange(self, event):
        # Fill date list and select first date
        global GlobalDBLocation
        ClientDB = ET.parse(os.path.join(GlobalDBLocation, 'PatientDB.XML'))
        ClientElementList = ClientDB.getroot()
        PatientElementList = ClientElementList.find(XMLTagTextPrepare(self.ReportClientChoice.GetValue()))
        DateElementList = PatientElementList.find(XMLTagTextPrepare(self.ReportPatientChoice.GetValue()))
        self.ReportDateChoice.Clear()
        if DateElementList:
            for Dates in reversed(DateElementList):
                # reformat date time string.  Saved in awkward format due to XML requirements
                DataDateTime = time.strptime(Dates.tag[5:], "%Y-%m-%d-%H-%M-%S")
                self.ReportDateChoice.Append(
                    Dates.attrib['Treatment'] + time.strftime(" %a, %d %b %Y %H:%M:%S", DataDateTime))
            if len(DateElementList) > 1:
                self.ReportDateChoice.Append('All Dates')
            self.ReportsGenerateButton.Enable()
        else:
            self.ReportDateChoice.Append('No Dates')
            self.ReportsGenerateButton.Disable()

        self.ReportDateChoice.SetSelection(0)

    def on_ReportDestination_mouseDown(self, event):
    # """ get directory for report saving"""
        DirDialog = wx.DirDialog(None, message = 'Choose a directory to save Reports')
        if DirDialog.ShowModal() == wx.ID_OK:
            if os.access(DirDialog.GetPath(), os.W_OK):
                self.ReportDestination.SetValue(DirDialog.GetPath())
            else:
                wx.MessageBox('Cannot write to this directory. Reselect.', '')
                DirDialog.Destroy()

    def on_ReportsGenerateButton(self, event):
        global GlobalDBLocation
        # verify destination exists
        if os.path.isdir(self.ReportDestination.GetValue()) == False:
            self.on_ReportDestination_mouseDown(self.on_ReportDestination_mouseDown)
        # get basic information
        InfoDict = {}
        InfoDict['Facility'] = self.FacilityText
        InfoDict['Client'] = self.ReportClientChoice.GetValue()
        InfoDict['Patient'] = self.ReportPatientChoice.GetValue()

        # read from database
        ClientDB = ET.parse(os.path.join(GlobalDBLocation, 'PatientDB.XML'))
        ClientElementList = ClientDB.getroot()
        PatientElementList = ClientElementList.find(XMLTagTextPrepare(self.ReportClientChoice.GetValue()))
        DateElementList = PatientElementList.find(XMLTagTextPrepare(self.ReportPatientChoice.GetValue()))
        Dates = []
        if self.ReportDateChoice.GetValue() == 'All Dates':
            Dates = DateElementList
        else:
            ListLen = len(DateElementList)
            Dates = DateElementList[
                    ListLen - self.ReportDateChoice.GetSelection() - 1:ListLen - self.ReportDateChoice.GetSelection()]

        for Date in Dates:
            InfoDict['Doctor'] = Date.attrib['Doctor']
            InfoDict['Treatment'] = Date.attrib['Treatment']
            InfoDict['DosePattern'] = Date.attrib['DosePattern']
            InfoDict['DoseCount'] = Date.attrib['DoseCount']

            Reports.GenerateReport(self.ReportDestination.GetValue(), self.ReportFormatChoice.GetValue(), InfoDict, Date.text)
        # Reports.GenerateReport(self.ReportDestination.GetValue(), self.ReportFormatChoice.GetValue(), InfoDict, Date.text)

    def on_ReportsBackButton(self, event):
        # SettingsDict = {}
        # SettingsDict['reportformat'] = self.ReportFormatChoice.GetValue()
        # SaveConfigFile(SettingsDict)
        self.ReportClientChoice.SetValue('')
        self.ReportPatientChoice.SetValue('')
        self.ReportDateChoice.SetValue('')
        self.Hide()
        frame.Show()


class AddPatientWindow(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, title='Add Client/Patient', size=(CalcFormLength(), CalcFormHeight()),
                          style=wx.DEFAULT_FRAME_STYLE & ~(wx.RESIZE_BORDER | wx.CLOSE_BOX | wx.MAXIMIZE_BOX))
        wx.Frame.CenterOnScreen(self)
        panel = wx.Panel(self, -1)
        BackGroundColour = (233, 228, 214)
        panel.SetBackgroundColour(BackGroundColour)
        font16 = wx.Font(16, wx.FONTFAMILY_SWISS, wx.NORMAL, wx.NORMAL)
        font12 = wx.Font(12, wx.FONTFAMILY_SWISS, wx.NORMAL, wx.NORMAL)
        font10 = wx.Font(10, wx.FONTFAMILY_SWISS, wx.NORMAL, wx.NORMAL)

        # Status bar
        self.StatusBar = wx.StatusBar(self, -1)
        self.StatusBar.SetFieldsCount(5)

        # battery life progress bar in status bar
        self.ProgressBar = wx.Gauge(self.StatusBar, -1, style=wx.GA_HORIZONTAL | wx.GA_SMOOTH)
        rect = self.StatusBar.GetFieldRect(4)
        self.ProgressBar.SetPosition((rect.x + 2, rect.y + 2))
        self.ProgressBar.SetSize((rect.width - 4, rect.height - 4))

        self.AddPatientAddText = wx.StaticText(panel, -1, u'Add Client/Patient:', pos=(15, 0), size=(120, 30))
        self.AddPatientAddText.SetFont(font16)

        self.AddPatientClientText = wx.StaticText(panel, -1, u'New Client', pos=(15, 25), size=(120, 30))
        self.AddPatientClientText.SetFont(font16)

        self.AddPatientNewClientTextCtrl = wx.TextCtrl(panel, -1, '', pos=(200, 25), size=(175, -1))
        self.AddPatientNewClientTextCtrl.SetFont(font10)
        self.AddPatientNewClientTextCtrl.SetMaxLength(MAX_STRING_LENGTH)
        #		self.Bind(wx.EVT_TEXT, self.on_AddPatientNewClientTextChange, self.AddPatientNewClientTextCtrl)

        self.AddPatientClientText = wx.StaticText(panel, -1, u'OR', pos=(150, 50), size=(120, 30))
        self.AddPatientClientText.SetFont(font16)

        self.AddPatientSelectClientText = wx.StaticText(panel, -1, u'Existing Client', pos=(15, 75), size=(120, 30))
        self.AddPatientSelectClientText.SetFont(font16)

        self.AddPatientClientChoice = wx.ComboBox(panel, -1, pos=(200, 75), size=(175, 25), style=wx.CB_DROPDOWN)
        self.AddPatientClientChoice.SetFont(font12)
        self.Bind(wx.EVT_COMBOBOX, self.on_AddPatientClientChoiceChange, self.AddPatientClientChoice)
        #		self.Bind(wx.EVT_TEXT, self.on_AddPatientClientTextChange, self.AddPatientClientChoice)

        self.AddPatientNewPatientText = wx.StaticText(panel, -1, u'New Patient', pos=(15, 100), size=(120, 30))
        self.AddPatientNewPatientText.SetFont(font16)

        self.AddPatientNewPatientTextCtrl = wx.TextCtrl(panel, -1, '', pos=(200, 100), size=(175, -1))
        self.AddPatientNewPatientTextCtrl.SetFont(font10)
        self.AddPatientNewPatientTextCtrl.SetMaxLength(MAX_STRING_LENGTH)
        self.Bind(wx.EVT_TEXT, self.on_AddPatientNewPatientTextCtrl, self.AddPatientNewPatientTextCtrl)

        self.AddPatientClearButton = wx.Button(panel, -1, label='Clear', pos=(50, 125), size=(150, 50))
        self.AddPatientClearButton.SetFont(font16)
        self.Bind(wx.EVT_BUTTON, self.on_AddPatientClearButton, self.AddPatientClearButton)

        self.AddPatientAddButton = wx.Button(panel, -1, label='Add', pos=(200, 125), size=(150, 50))
        self.AddPatientAddButton.SetFont(font16)
        self.Bind(wx.EVT_BUTTON, self.on_AddPatientAddButton, self.AddPatientAddButton)

        self.DeletePatientAddText = wx.StaticText(panel, -1, u'Delete Client/Patient:', pos=(15, 200), size=(120, 30))
        self.DeletePatientAddText.SetFont(font16)

        self.WriteClientText = wx.StaticText(panel, -1, u'Client', pos=(15, 225), size=(120, 30))
        self.WriteClientText.SetFont(font16)

        self.DeleteClientChoice = wx.ComboBox(panel, -1, pos=(200, 225), size=(175, -1), style=wx.CB_DROPDOWN)
        self.DeleteClientChoice.SetFont(font12)
        self.Bind(wx.EVT_COMBOBOX, self.on_DeleteClientChange, self.DeleteClientChoice)
        #		self.Bind(wx.EVT_TEXT, self.on_DeleteClientTextChange, self.DeleteClientChoice)

        self.DeletePatientText = wx.StaticText(panel, -1, u'Patient', pos=(15, 250), size=(120, 30))
        self.DeletePatientText.SetFont(font16)

        self.DeletePatientChoice = wx.ComboBox(panel, -1, pos=(200, 250), size=(175, -1),
                                               style=wx.CB_DROPDOWN | wx.CB_READONLY)
        self.DeletePatientChoice.SetFont(font12)

        self.DeletePatientClearButton = wx.Button(panel, -1, label='Clear', pos=(50, 275), size=(150, 50))
        self.DeletePatientClearButton.SetFont(font16)
        self.Bind(wx.EVT_BUTTON, self.on_DeletePatientClearButton, self.DeletePatientClearButton)

        self.DeletePatientButton = wx.Button(panel, -1, label='Delete', pos=(200, 275), size=(150, 50))
        self.DeletePatientButton.SetFont(font16)
        self.Bind(wx.EVT_BUTTON, self.on_DeletePatientButton, self.DeletePatientButton)

        self.AddPatientBackButton = wx.Button(panel, -1, label='Back', pos=(450, 350), size=(150, 50))
        self.AddPatientBackButton.SetFont(font16)
        self.Bind(wx.EVT_BUTTON, self.on_AddPatientBackButton, self.AddPatientBackButton)

    def on_DeleteClientTextChange(self, event):
        if platform.system() != "Windows":
            global GlobalDBLocation
            self.DeletePatientClearButton.Enable()
            ClientDB = ET.parse(os.path.join(GlobalDBLocation, 'PatientDB.XML'))
            ClientList = []
            for Clients in ClientDB.getroot():
                ClientList.append(XMLTagTextExtract(Clients.tag))

            EnteredText = self.DeleteClientChoice.GetValue()
            ClientCheck = DataBaseClientNameCheck(ClientList, EnteredText)
            if ClientCheck[0] == True:
                self.DeleteClientChoice.Clear()
                for Client in sorted(ClientCheck[1], key=lambda s: s.lower()):
                    self.DeleteClientChoice.Append(Client)
            else:
                self.DeleteClientChoice.SetValue(EnteredText[:-1])

            if self.DeleteClientChoice.GetValue() in ClientList \
                    or len(self.DeleteClientChoice.GetValue()) == 0:
                self.on_DeleteClientChange(self.on_DeleteClientChange)

    def on_DeleteClientChange(self, event):
        # Fill patient list and select first patient
        global GlobalDBLocation
        ClientDB = ET.parse(os.path.join(GlobalDBLocation, 'PatientDB.XML'))
        ClientElementList = ClientDB.getroot()
        PatientElementList = ClientElementList.find(XMLTagTextPrepare(self.DeleteClientChoice.GetValue()))
        self.DeletePatientChoice.SetValue('')
        self.DeletePatientChoice.Clear()
        self.DeletePatientButton.Disable()

        # get patient list and sort alphabetically
        if PatientElementList:
            PatientList = []
            for Patients in PatientElementList:
                PatientList.append(XMLTagTextExtract(Patients.tag))
            for Patient in sorted(PatientList, key=lambda s: s.lower()):
                self.DeletePatientChoice.Append(Patient)
        self.DeletePatientButton.Enable()
        self.DeletePatientChoice.Append('Delete Client')
        self.DeletePatientChoice.SetSelection(0)

    def on_DeletePatientClearButton(self, event):
        global GlobalDBLocation
        self.DeletePatientChoice.SetValue('')
        self.DeleteClientChoice.Enable()
        self.DeleteClientChoice.SetValue('')
        self.DeletePatientClearButton.Disable()
        self.DeletePatientButton.Disable()
        self.DeletePatientButton.Disable()
        ''' Fill Client listbox'''
        ClientDB = ET.parse(os.path.join(GlobalDBLocation, 'PatientDB.XML'))
        self.DeleteClientChoice.Clear()
        ClientList = []
        for Clients in ClientDB.getroot():
            ClientList.append(XMLTagTextExtract(Clients.tag))
        for Client in sorted(ClientList):
            self.DeleteClientChoice.Append(Client)

    def on_DeletePatientButton(self, event):
        global GlobalDBLocation
        ClientDB = ET.parse(os.path.join(GlobalDBLocation, 'PatientDB.XML'))
        ClientElementList = ClientDB.getroot()
        PatientElementList = ClientElementList.find(XMLTagTextPrepare(self.DeleteClientChoice.GetValue()))
        if self.DeletePatientChoice.GetValue() == 'Delete Client':
            ClientToDelete = ClientElementList.find(XMLTagTextPrepare(self.DeleteClientChoice.GetValue()))
            ClientElementList.remove(ClientToDelete)
        else:
            PatientToDelete = PatientElementList.find(XMLTagTextPrepare(self.DeletePatientChoice.GetValue()))
            PatientElementList.remove(PatientToDelete)
        XMLIndent(ClientElementList)
        ClientDB.write(os.path.join(GlobalDBLocation, 'PatientDB.XML'))
        self.on_DeletePatientClearButton(self.on_DeletePatientClearButton)

    def on_AddPatientNewPatientTextCtrl(self, event):
        EnteredText = self.AddPatientNewPatientTextCtrl.GetValue()
        if XMLCharacterCheck(EnteredText):
            self.AddPatientAddButton.Enable()
            self.AddPatientClearButton.Enable()
        else:
            self.AddPatientNewPatientTextCtrl.ChangeValue(EnteredText[:-1])

    def on_AddPatientClearButton(self, event):
        global GlobalDBLocation
        self.AddPatientNewPatientTextCtrl.SetValue('')
        self.AddPatientNewClientTextCtrl.SetValue('')
        self.AddPatientNewClientTextCtrl.Enable()
        self.AddPatientClientChoice.Enable()
        self.AddPatientClientChoice.SetValue('')
        self.AddPatientClearButton.Disable()
        self.AddPatientAddButton.Disable()
        ''' Fill Client listbox'''
        ClientDB = ET.parse(os.path.join(GlobalDBLocation, 'PatientDB.XML'))
        self.AddPatientClientChoice.Clear()
        ClientList = []
        for Clients in ClientDB.getroot():
            ClientList.append(XMLTagTextExtract(Clients.tag))
        for Client in sorted(ClientList):
            self.AddPatientClientChoice.Append(Client)

    def on_AddPatientNewClientTextChange(self, event):
        if platform.system() != "Windows":
            EnteredText = self.AddPatientNewClientTextCtrl.GetValue()
            if XMLCharacterCheck(EnteredText):
                self.AddPatientClientChoice.SetValue('')
                self.AddPatientClientChoice.Disable()
                self.AddPatientClearButton.Enable()
            else:
                self.AddPatientNewClientTextCtrl.ChangeValue(EnteredText[:-1])

    def on_AddPatientClientTextChange(self, event):
        if platform.system() != "Windows":
            global GlobalDBLocation
            ClientDB = ET.parse(os.path.join(GlobalDBLocation, 'PatientDB.XML'))
            ClientList = []
            for Clients in ClientDB.getroot():
                ClientList.append(XMLTagTextExtract(Clients.tag))

            EnteredText = self.AddPatientClientChoice.GetValue()
            ClientCheck = DataBaseClientNameCheck(ClientList, EnteredText)
            if ClientCheck[0] == True:
                self.AddPatientClientChoice.Clear()
                for Client in sorted(ClientCheck[1], key=lambda s: s.lower()):
                    self.AddPatientClientChoice.Append(Client)
            else:
                self.AddPatientClientChoice.SetValue(EnteredText[:-1])
            if len(EnteredText) == 0:
                self.on_AddPatientClearButton(self.on_AddPatientClearButton)
            elif EnteredText in ClientCheck[1]:
                self.on_AddPatientClientChoiceChange(self.on_AddPatientClientChoiceChange)

    def on_AddPatientClientChoiceChange(self, event):
        self.AddPatientNewClientTextCtrl.Disable()
        self.AddPatientClearButton.Enable()

    def on_AddPatientAddButton(self, event):
        global GlobalDBLocation
        PatientText = self.AddPatientNewPatientTextCtrl.GetValue().strip()
        if len(self.AddPatientNewClientTextCtrl.GetValue().strip()):
            ClientText = self.AddPatientNewClientTextCtrl.GetValue().strip()
        else:
            ClientText = self.AddPatientClientChoice.GetValue()
        if len(PatientText) and len(ClientText):
            ClientDB = ET.parse(os.path.join(GlobalDBLocation, 'PatientDB.XML'))
            ClientElementList = ClientDB.getroot()
            PatientText = XMLTagTextPrepare(PatientText)
            ClientText = XMLTagTextPrepare(ClientText)

            # if client doesn't exist, add
            if ClientElementList.find(ClientText) == None:
                NewClient = ET.Element(ClientText)
                ClientElementList.append(NewClient)

            # if patient doesn't exist, add
            PatientElementList = ClientElementList.find(ClientText)
            NewPatient = PatientElementList.find(PatientText)
            if NewPatient == None:
                NewPatient = ET.Element(PatientText)
                PatientElementList.insert(0, NewPatient)
            # save file
            XMLIndent(ClientElementList)
            ClientDB.write(os.path.join(GlobalDBLocation, 'PatientDB.XML'))
            self.on_AddPatientClearButton(self.on_AddPatientClearButton)

        else:  # pop up
            dlg = wx.MessageDialog(None, 'Enter Patient name and \nEnter Client Name or Select Existing Client',
                                   'Error', wx.OK | wx.ICON_EXCLAMATION)
            DialogReturn = dlg.ShowModal()

    def on_AddPatientBackButton(self, event):
        self.on_AddPatientClearButton(self.on_AddPatientClearButton)
        self.DeleteClientChoice.SetValue('')
        self.DeletePatientChoice.SetValue('')
        self.Hide()
        frame.Show()


class SettingsWindow(wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, title='Settings', size=(CalcFormLength(), CalcFormHeight()),
                          style=wx.DEFAULT_FRAME_STYLE & ~(wx.RESIZE_BORDER | wx.CLOSE_BOX | wx.MAXIMIZE_BOX))
        wx.Frame.CenterOnScreen(self)
        panel = wx.Panel(self, -1)
        BackGroundColour = (233, 228, 214)
        panel.SetBackgroundColour(BackGroundColour)
        font16 = wx.Font(16, wx.FONTFAMILY_SWISS, wx.NORMAL, wx.NORMAL)
        font12 = wx.Font(12, wx.FONTFAMILY_SWISS, wx.NORMAL, wx.NORMAL)
        font10 = wx.Font(10, wx.FONTFAMILY_SWISS, wx.NORMAL, wx.NORMAL)

        # Status bar
        self.StatusBar = wx.StatusBar(self, -1)
        self.StatusBar.SetFieldsCount(5)

        # battery life progress bar in status bar
        self.ProgressBar = wx.Gauge(self.StatusBar, -1, style=wx.GA_HORIZONTAL | wx.GA_SMOOTH)
        rect = self.StatusBar.GetFieldRect(4)
        self.ProgressBar.SetPosition((rect.x + 2, rect.y + 2))
        self.ProgressBar.SetSize((rect.width - 4, rect.height - 4))

        self.SettingsFacilityText = wx.StaticText(panel, -1, u'Facility', pos=(15, 25), size=(120, 30))
        self.SettingsFacilityText.SetFont(font16)

        self.SettingsFacilityTextCtrl = wx.TextCtrl(panel, -1, '', pos=(150, 25), size=(175, -1))
        self.SettingsFacilityTextCtrl.SetFont(font12)
        self.SettingsFacilityTextCtrl.SetMaxLength(MAX_STRING_LENGTH)

        self.SettingsCommandText = wx.StaticText(panel, -1, u'Configure Commands', pos=(15, 300), size=(120, 30))
        self.SettingsCommandText.SetFont(font16)

        self.SettingsCommandTextCtrl = wx.TextCtrl(panel, -1, '', pos=(225, 300), size=(175, -1),
                                                   style=wx.TE_PROCESS_ENTER)
        self.SettingsCommandTextCtrl.SetFont(font12)
        self.SettingsCommandTextCtrl.SetMaxLength(MAX_STRING_LENGTH)
        self.Bind(wx.EVT_TEXT_ENTER, self.on_SettingsCommandTextCtrl_Enter, self.SettingsCommandTextCtrl)

        self.SettingsDrText = wx.StaticText(panel, -1, u'Doctors', pos=(15, 75), size=(120, 30))
        self.SettingsDrText.SetFont(font16)

        self.SettingsDrList = wx.TextCtrl(parent=panel, id=-1, pos=(150, 75), size=(175, 100), style=wx.TE_MULTILINE)
        self.SettingsDrList.SetFont(font12)

        self.SettingsMedText = wx.StaticText(panel, -1, u'Treatments', pos=(15, 200), size=(120, 30))
        self.SettingsMedText.SetFont(font16)

        self.SettingsMedList = wx.TextCtrl(panel, -1, pos=(150, 200), size=(175, 100), style=wx.TE_MULTILINE)
        self.SettingsMedList.SetFont(font12)

        self.SettingsBackButton = wx.Button(panel, -1, label='Back', pos=(450, 350), size=(150, 50))
        self.SettingsBackButton.SetFont(font16)
        self.Bind(wx.EVT_BUTTON, self.on_SettingsBackButton, self.SettingsBackButton)

        self.SettingsDatabaseLocationText = wx.StaticText(panel, -1, u'Database Location', pos=(15, 325),
                                                          size=(120, 30))
        self.SettingsDatabaseLocationText.SetFont(font16)

        self.SettingsDatabaseLocation = wx.TextCtrl(panel, -1, u'/', pos=(225, 325), size=(350, 25),
                                                    style=wx.TE_READONLY)
        self.SettingsDatabaseLocation.SetFont(font12)
        self.SettingsDatabaseLocation.Bind(wx.EVT_LEFT_DOWN, self.on_SettingsDatabaseDestination_mouseDown)

    def on_SettingsCommandTextCtrl_Enter(self, event):
        global GlobalSerialPort
        GlobalSerialPortText = str(GlobalSerialPort)  # windows GSP is an int, convert to str to prevent error
        if GlobalSerialPortText.find('No', 0, 3) == 0 or GlobalSerialPortText.find('Ba', 0, 3) == 0:
            dlg = wx.MessageDialog(None, GlobalSerialPort, 'Error', wx.OK)
            DialogReturn = dlg.ShowModal()
        else:
            if str(self.SettingsCommandTextCtrl.GetValue())[:3] == 'new':
                CapIO.WriteString(GlobalSerialPort, 'T\r')
                now = datetime.datetime.now()
                CapIO.WriteString(GlobalSerialPort,
                                  'B:{:0>2}:{:0>2}:{:0>2}:{:0>2}:{:0>2}:{:0>2}:{:0>2}\r'.format(now.second, now.minute,
                                                                                                now.hour, now.day,
                                                                                                now.weekday() + 1,
                                                                                                now.month,
                                                                                                now.year - 2000))
                CapIO.WriteString(GlobalSerialPort, 'N\r')
                self.SettingsCommandTextCtrl.SetValue('')
            if str(self.SettingsCommandTextCtrl.GetValue())[:3] == 'setclock':
                now = datetime.datetime.now()
                CapIO.WriteString(GlobalSerialPort,
                                  'B:{:0>2}:{:0>2}:{:0>2}:{:0>2}:{:0>2}:{:0>2}:{:0>2}\r'.format(now.second, now.minute,
                                                                                                now.hour, now.day,
                                                                                                now.weekday() + 1,
                                                                                                now.month,
                                                                                                now.year - 2000))
                self.SettingsCommandTextCtrl.SetValue('')
            elif str(self.SettingsCommandTextCtrl.GetValue())[:3] == 'now':
                self.SettingsCommandTextCtrl.SetValue(CapIO.ReadTime(GlobalSerialPort, 'C\r'))
            elif str(self.SettingsCommandTextCtrl.GetValue())[:4] == 'demo':
                CapIO.WriteString(GlobalSerialPort, 'AR\r')
                self.SettingsCommandTextCtrl.SetValue('')
            elif str(self.SettingsCommandTextCtrl.GetValue())[:4] == 'age':
                BatAge = CapIO.ReadBatteryAge(GlobalSerialPort)
                # print BatAge
                if BatAge < 0:
                    self.SettingsCommandTextCtrl.SetValue('Read Error')
                else:
                    self.SettingsCommandTextCtrl.SetValue('{0:.2f} Days'.format(BatAge))

    def on_SettingsDatabaseDestination_mouseDown(self, event):
        """ get directory for database"""
        global GlobalDBLocation
        DirDialog = wx.DirDialog(None, message='Choose the location of the patient database')
        if DirDialog.ShowModal() == wx.ID_OK:
            if os.access(DirDialog.GetPath(), os.W_OK):
                self.SettingsDatabaseLocation.SetValue(DirDialog.GetPath())
                GlobalDBLocation = self.SettingsDatabaseLocation.GetValue()
            else:
                wx.MessageBox('Cannot write to this directory. Reselect.', '')
        DirDialog.Destroy()

    def on_SettingsBackButton(self, event):
        # SettingsDict = {}
        # SettingsDict['Facility'] = self.SettingsFacilityTextCtrl.Value
        # SettingsDict['Treatments'] = self.SettingsMedList.GetValue()
        # SettingsDict['Doctors'] = self.SettingsDrList.GetValue()
        # SettingsDict['DatabaseLocation'] = self.SettingsDatabaseLocation.GetValue()
        # SaveConfigFile(SettingsDict)
        self.SettingsCommandTextCtrl.SetValue('')
        self.Hide()
        frame.Show()


class MainWindow(wx.Frame):
    """ real main"""

    def __init__(self):
        global GlobalSerialPort
        global GlobalDBLocation
        wx.Frame.__init__(self, None, wx.ID_ANY, 'Ccap v' + SWVer,
                          style=wx.DEFAULT_FRAME_STYLE & ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX),
                          size=(CalcFormLength(), CalcFormHeight()))
        panel = wx.Panel(self, -1)
        wx.Frame.CenterOnScreen(self)
        BackGroundColour = (233, 228, 214)
        panel.SetBackgroundColour(BackGroundColour)

        self.WritePanel = WriteWindow(self)
        self.AddPatientPanel = AddPatientWindow(self)
        self.ReportsPanel = ReportsWindow(self)
        self.SettingsPanel = SettingsWindow(self)

        font16 = wx.Font(16, wx.FONTFAMILY_SWISS, wx.NORMAL, wx.NORMAL)
        font15 = wx.Font(15, wx.FONTFAMILY_SWISS, wx.NORMAL, wx.NORMAL)
        font12 = wx.Font(12, wx.FONTFAMILY_SWISS, wx.NORMAL, wx.NORMAL)
        font10 = wx.Font(10, wx.FONTFAMILY_SWISS, wx.NORMAL, wx.NORMAL)

        # Main Page
        self.timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.OnTimer, self.timer)
        self.timer.Start(3000)

        # Status bar
        self.StatusBar = wx.StatusBar(self, -1)
        self.StatusBar.SetFieldsCount(5)

        # battery life progress bar in status bar
        self.ProgressBar = wx.Gauge(self.StatusBar, -1, style=wx.GA_HORIZONTAL | wx.GA_SMOOTH)
        rect = self.StatusBar.GetFieldRect(4)
        self.ProgressBar.SetPosition((rect.x + 2, rect.y + 2))
        self.ProgressBar.SetSize((rect.width - 4, rect.height - 4))

        jpg1 = wx.Image('Ccap0.25.jpg', wx.BITMAP_TYPE_ANY).ConvertToBitmap()
        # bitmap upper left corner is in the position tuple (x, y) = (5, 5)
        wx.StaticBitmap(self, -1, jpg1, (230, 30), (jpg1.GetWidth(), jpg1.GetHeight()))

        # self.MainLogoText = wx.StaticText(panel, -1, u'cCap', pos = (180, 100), size = (120, 30))
        # self.MainLogoText.SetFont(wx.Font(72, wx.FONTFAMILY_SWISS, wx.ITALIC, wx.NORMAL))

        self.ReadDataButton = wx.Button(panel, -1, label='Read Cap', pos=(0, 350), size=(150, 50))
        self.ReadDataButton.SetFont(font16)
        self.Bind(wx.EVT_BUTTON, self.on_ReadDataButton, self.ReadDataButton)

        self.WritePageGotoButton = wx.Button(panel, -1, label='Write Cap', pos=(150, 350), size=(150, 50))
        self.WritePageGotoButton.SetFont(font16)
        self.Bind(wx.EVT_BUTTON, self.on_WritePageGotoButton, self.WritePageGotoButton)
        
        # ---------------------------------------------------------------------------------------------------
        # Disable for Release
        # ---------------------------------------------------------------------------------------------------
        self.ReadRAMButton = wx.Button(panel, -1, label='Dump RAM', pos=(0, 0), size=(150, 50))
        self.ReadRAMButton.SetFont(font16)
        self.Bind(wx.EVT_BUTTON, self.on_ReadRAMButton, self.ReadRAMButton)

        self.ReadEEPROMButton = wx.Button(panel, -1, label='Dump EEPROM', pos=(0, 50), size=(150, 50))
        self.ReadEEPROMButton.SetFont(font16)
        self.Bind(wx.EVT_BUTTON, self.on_ReadEEPROMButton, self.ReadEEPROMButton)
        # ---------------------------------------------------------------------------------------------------
        
        self.SettingsPageGotoButton = wx.Button(panel, -1, label='Settings', pos=(300, 350), size=(150, 50))
        self.SettingsPageGotoButton.SetFont(font16)
        self.Bind(wx.EVT_BUTTON, self.on_SettingsPageGotoButton, self.SettingsPageGotoButton)

        self.ExitButton = wx.Button(panel, -1, label='Exit', pos=(450, 350), size=(150, 50))
        self.ExitButton.SetFont(font16)
        self.Bind(wx.EVT_BUTTON, self.on_ExitButton, self.ExitButton)

        self.ReportPageGotoButton = wx.Button(panel, -1, label='Get Report', pos=(0, 300), size=(150, 50))
        self.ReportPageGotoButton.SetFont(font16)
        self.Bind(wx.EVT_BUTTON, self.on_ReportPageGotoButton, self.ReportPageGotoButton)
        
        self.AddPatientPageGotoButton = wx.Button(panel, -1, label='Add Client/Patient', pos=(150, 300), size=(150, 50))
        self.AddPatientPageGotoButton.SetFont(font12)
        self.Bind(wx.EVT_BUTTON, self.on_AddPatientPageGotoButton, self.AddPatientPageGotoButton)
        # end of GUI initialization

        # load settings
        config = configparser.ConfigParser(
            defaults={'facility': '', 'treatments': '', 'doctors': '', 'reportdestination': '', 'reportformat': 'Text'})
        if not (os.path.isfile('Ccap.cfg')):
            file = open("Ccap.cfg", "w")
            file.write("[Settings]\n")
            file.close()
        config.read('Ccap.cfg')

        # Fill Doctor list
        DoctorList = config.get('Settings', 'doctors')
        self.SettingsPanel.SettingsDrList.AppendText(DoctorList)

        # Fill Medicine list
        TreatmentList = config.get('Settings', 'treatments')
        self.SettingsPanel.SettingsMedList.AppendText(TreatmentList)

        self.SettingsPanel.SettingsFacilityTextCtrl.SetValue(config.get('Settings', 'facility'))
        self.WritePanel.WriteFacilityTextCtrl.SetValue(config.get('Settings', 'facility'))
        self.SettingsPanel.SettingsDatabaseLocation.SetValue(config.get('Settings', 'DatabaseLocation'))
        GlobalDBLocation = config.get('Settings', 'DatabaseLocation')

        self.WritePanel.WriteDrChoice.SetValue(config.get('Settings', 'lastdoctor'))
        self.WritePanel.WriteTreatmentChoice.SetValue(config.get('Settings', 'lasttreatment'))

        self.WritePanel.MaxBatteryLife = config.get('Settings', 'batterylife')

        BaseStationList = CapIO.FindBaseStation()
        GlobalSerialPort = BaseStationList[0]
        SerialPortName = str(BaseStationList[0])
        if SerialPortName.find('No', 0, 2) == 0:
            dlg = wx.MessageDialog(None, GlobalSerialPort, 'Error', wx.OK)
            DialogReturn = dlg.ShowModal()
            SerialPortName = CapIO.NO_COM_PORT_TEXT
            self.Update_Statusbar(SerialPortName, 0)
        else:
            SerialPortHead = ''
            if platform.system() == "Windows":
                if 'COM' in SerialPortName:
                    self.Update_Statusbar(str(SerialPortName),0)
            else:
                SlashLoc = SerialPortName.rfind('/')
                self.Update_Statusbar(SerialPortHead + SerialPortName[SlashLoc + 1:], 0)

            if CapIO.NO_BASE_STATION_TEXT in SerialPortName:
                print("Base station not found")
                dlg = wx.MessageDialog(None, GlobalSerialPort, 'Error', wx.OK)
                DialogReturn = dlg.ShowModal()
                SerialPortName = ''
                self.Update_Statusbar(SerialPortName, 0)
                BaseStationTitle = 'Base Station Error'
                self.Update_Statusbar(BaseStationTitle, 1)
            else:
                BaseStationTitle = BaseStationList[1][2:-5].strip()
                BaseStationTitle = BaseStationTitle.replace(':', ' ')
                BaseStationTitle = BaseStationTitle.replace('Z\\r', '')
                self.Update_Statusbar(str(BaseStationTitle), 1)
        self.OnTimer(wx.EVT_TIMER)

    def on_ReadRAMButton(self,event):
        global GlobalSerialPort
        CapIO.DumpRAM(GlobalSerialPort)

    def on_ReadEEPROMButton(self,event):
        global GlobalSerialPort
        CapIO.DumpEEPROM(GlobalSerialPort)

    def on_WritePageGotoButton(self, event):
        global GlobalDBLocation
        self.WritePanel.WriteFacilityTextCtrl.SetValue(self.SettingsPanel.SettingsFacilityTextCtrl.Value)
        # Fill Doctor list
        TempString = self.SettingsPanel.SettingsDrList.GetValue().strip()
        TempList = TempString.split('\n')
        self.WritePanel.WriteDrChoice.Clear()
        for Items in TempList:
            self.WritePanel.WriteDrChoice.Append(Items)
        # self.WritePanel.WriteDrChoice.SetSelection(0)
        # Fill Medicine list
        TempString = self.SettingsPanel.SettingsMedList.GetValue().strip()
        TempList = TempString.split('\n')
        self.WritePanel.WriteTreatmentChoice.Clear()
        for Items in TempList:
            self.WritePanel.WriteTreatmentChoice.Append(Items)
        # self.WritePanel.WriteTreatmentChoice.SetSelection(0)
        # Fill Client list
        self.WritePanel.WriteClientChoice.Clear()
        ClientDB = ET.parse(os.path.join(GlobalDBLocation, 'PatientDB.XML'))
        ClientList = []
        for Clients in ClientDB.getroot():
            ClientList.append(XMLTagTextExtract(Clients.tag))
        for Client in sorted(ClientList, key=lambda s: s.lower()):
            self.WritePanel.WriteClientChoice.Append(Client)
        self.WritePanel.WriteClientChoice.SetSelection(0)
        # Clear Patient list
        self.WritePanel.WritePatientChoice.Clear()
        self.WritePanel.WritePatientChoice.SetSelection(0)
        self.WritePanel.WriteBuzzerEnableCheckBox.SetValue(True)

        self.WritePanel.Show()
        frame.Hide()

    def on_SettingsPageGotoButton(self, event):
        self.SettingsPanel.Show()
        frame.Hide()

    def on_ReadDataButton(self, event):
        global GlobalSerialPort
        global GlobalDBLocation

        if GlobalSerialPort == CapIO.NO_COM_PORT_TEXT:
            dlg = wx.MessageDialog(None, 'Read Error No Com Port', 'Error', wx.OK)
            DialogReturn = dlg.ShowModal()
        else:
            NotRead = True
            self.StatusBar.SetStatusText('Reading Cap', 3)
            for i in range(0, 2):
                if NotRead:
                    try:
                        # read info from cap
                        RxInfoDict = CapIO.ReadSettings(GlobalSerialPort)
                        # print RxInfoDict
                        TimePointsText = ''
                        TimePoints = CapIO.ReadData(GlobalSerialPort)
                        StartTime = CapIO.ReadTime(GlobalSerialPort, 'H\r')
                        NotRead = False  # success don't try again

                        print("Successfully read from cap.")
                        print("\n-------------------------")
                        print("Start time: " + str(StartTime))

                        # open database file
                        ClientDB = ET.parse(os.path.join(GlobalDBLocation, 'PatientDB.XML'), ET.XMLParser())
                        # find client
                        ClientElementList = ClientDB.getroot()
                        ClientName = XMLTagTextPrepare(RxInfoDict['Client'])
                        if ClientName == '':
                            ClientName = 'Blank_Client'
                        # if client doesn't exist, add
                        if ClientElementList.find(ClientName) == None:
                            NewClient = ET.Element(ClientName)
                            ClientElementList.append(NewClient)
                        ClientElement = ClientElementList.find(ClientName)
                        print("Client Name: " + ClientName)
                        
                        # if patient doesn't exist, add
                        PatientName = XMLTagTextPrepare(RxInfoDict['Patient'])
                        if PatientName == '':
                            PatientName = 'Blank_Patient'
                        NewPatient = ClientElement.find(PatientName)
                        if NewPatient == None:
                            NewPatient = ET.Element(PatientName)
                            ClientElement.insert(0, NewPatient)
                        print("Patient: " + PatientName)
                        
                        # Add datetime element for patient
                        # ONLY ADDS DATE AND DATA IF PATIENT IS NEW
                        DateString = datetime.datetime.now().strftime("Date-%Y-%m-%d-%H-%M-%S")
                        NewDateTimeElement = ET.SubElement(NewPatient, DateString)
                        #print(RxInfoDict)

                        # Add Info
                        RxInfoDict['DosePattern'] = DosePatterns[RxInfoDict['DosePattern']]
                        NewDateTimeElement.attrib['Facility'] = RxInfoDict['Facility']
                        NewDateTimeElement.attrib['Doctor'] = RxInfoDict['Doctor']
                        NewDateTimeElement.attrib['Treatment'] = RxInfoDict['Treatment']
                        NewDateTimeElement.attrib['DosePattern'] = RxInfoDict['DosePattern']
                        NewDateTimeElement.attrib['DoseCount'] = RxInfoDict['DoseCount']

                        print('DosePattern: ' + RxInfoDict['DosePattern'])
                        print('Facility: ' + RxInfoDict['Facility'])
                        print('Doctor: ' + RxInfoDict['Doctor'])
                        print('Treatment: ' + RxInfoDict['Treatment'])
                        print('DosePattern: ' + RxInfoDict['DosePattern'])
                        print('DoseCount: ' + RxInfoDict['DoseCount'])

                        # Add Data
                        print("Time points number: " + str(len(TimePoints)))
                        print(TimePoints)
                        for Points in TimePoints:
                            DataSubList = CapIO.RxdStrParse(str(Points.strip()))
                            TimePointsText += (
                                '{},{},{},{};'.format(DataSubList[3], DataSubList[2], DataSubList[1], DataSubList[0]))
                        NewDateTimeElement.text = TimePointsText
                        print("-------------------------\n")


                        # save file
                        XMLIndent(ClientElementList)
                        ClientDB.write(os.path.join(GlobalDBLocation, 'PatientDB.XML'))
                        print("Saving to PatientDB.xml")
                        print()

                        # # display report
                        # PatientElementList = ClientElementList.find(XMLTagTextPrepare(XMLTagTextExtract(ClientName)))
                        # DateElementList = PatientElementList.find(XMLTagTextPrepare(XMLTagTextExtract(PatientName)))
                        # Date = DateElementList[-1]
                        # Reports.GenerateReport("", u'Text', RxInfoDict, Date.text)
                        # print("Report Displayed")

                        # if cap done
                        PillCapTitleList = CapIO.Ping(GlobalSerialPort)
                        if PillCapTitleList[3].strip() == '7' or PillCapTitleList[3].strip() == '?':
                            dlg = wx.MessageDialog(None,
                                                   'Cap successfully read and data stored.\nPress OK to erase or Cancel to continue.',
                                                   'Success', wx.OK | wx.CANCEL)
                            DialogReturn = dlg.ShowModal()
                            if DialogReturn == wx.ID_OK:
                                self.StatusBar.SetStatusText('Erasing Cap', 3)
                                CapIO.Erase(GlobalSerialPort)
                        else:
                            dlg = wx.MessageDialog(None,
                                                   'Cap successfully read and data stored.\nThere are doses remaining in Cap.\nPress OK to Continue.',
                                                   'Success', wx.OK)
                            DialogReturn = dlg.ShowModal()

                    except:
                        e = sys.exc_info()[0]
                        print(e)
                        if (i == 1):
                            dlg = wx.MessageDialog(None,
                                                   'Read Error:\nCap should be top down.\nPress cap.\nCheck battery.\nCap may be blank.',
                                                   'Error', wx.OK)
                            DialogReturn = dlg.ShowModal()

    def on_AddPatientPageGotoButton(self, event):
        ''' Fill Client listbox'''
        global GlobalDBLocation
        ClientDB = ET.parse(os.path.join(GlobalDBLocation, 'PatientDB.XML'))
        self.AddPatientPanel.AddPatientClientChoice.Clear()
        ClientList = []
        for Clients in ClientDB.getroot():
            ClientList.append(XMLTagTextExtract(Clients.tag))
        for Client in sorted(ClientList):
            self.AddPatientPanel.AddPatientClientChoice.Append(Client)
            self.AddPatientPanel.DeleteClientChoice.Append(Client)
        self.AddPatientPanel.AddPatientClearButton.Disable()
        self.AddPatientPanel.AddPatientAddButton.Disable()
        self.AddPatientPanel.DeletePatientButton.Disable()
        self.AddPatientPanel.DeletePatientClearButton.Disable()
        self.AddPatientPanel.Show()
        frame.Hide()

    def on_ReportPageGotoButton(self, event):
        global GlobalDBLocation
        self.ReportsPanel.FacilityText = self.SettingsPanel.SettingsFacilityTextCtrl.Value
        ClientDB = ET.parse(os.path.join(GlobalDBLocation, 'PatientDB.XML'))
        self.ReportsPanel.ReportClientChoice.Clear()
        ClientList = []
        for Clients in ClientDB.getroot():
            ClientList.append(XMLTagTextExtract(Clients.tag))
        for Client in sorted(ClientList, key=lambda s: s.lower()):
            self.ReportsPanel.ReportClientChoice.Append(Client)
        self.ReportsPanel.ReportsGenerateButton.Disable()
        self.ReportsPanel.Show()
        frame.Hide()

    def Update_Statusbar(self, SBText, SBFrame):
        if SBFrame == 4:
            self.ProgressBar.SetValue(int(SBText))
            self.AddPatientPanel.ProgressBar.SetValue(int(SBText))
            self.WritePanel.ProgressBar.SetValue(int(SBText))
            self.SettingsPanel.ProgressBar.SetValue(int(SBText))
            self.ReportsPanel.ProgressBar.SetValue(int(SBText))
        else:
            self.StatusBar.SetStatusText(SBText, SBFrame)
            self.AddPatientPanel.SetStatusText(SBText, SBFrame)
            self.WritePanel.StatusBar.SetStatusText(SBText, SBFrame)
            self.SettingsPanel.StatusBar.SetStatusText(SBText, SBFrame)
            self.ReportsPanel.StatusBar.SetStatusText(SBText, SBFrame)

    def OnTimer(self, event):
        global GlobalSerialPort
        BatteryLife = 0
        # search for base station
        if GlobalSerialPort == CapIO.NO_COM_PORT_TEXT:
            BaseStationList = CapIO.FindBaseStation()
            print("Global for cap: " + str(BaseStationList[0]))
            GlobalSerialPort = BaseStationList[0]
            SerialPortName = str(BaseStationList[0])
            if SerialPortName.find('No', 0, 2) == -1:
                SerialPortHead = ''
                if platform.system() == "Windows":
                    if 'COM' in SerialPortName:
                        self.Update_Statusbar(str(SerialPortName),0)
                else:
                    SlashLoc = SerialPortName.rfind('/')
                    self.Update_Statusbar(SerialPortHead + SerialPortName[SlashLoc + 1:], 0)

                if CapIO.NO_BASE_STATION_TEXT in  SerialPortName:
                    SerialPortName = ''
                    self.Update_Statusbar(SerialPortName, 0)
                    BaseStationTitle = 'Base Station Error'
                    self.Update_Statusbar(BaseStationTitle, 1)
                else:
                    BaseStationTitle = BaseStationList[1][2:-5].strip()
                    BaseStationTitle = BaseStationTitle.replace(':', ' ')
                    BaseStationTitle = BaseStationTitle.replace('Z\r', '')
                    self.Update_Statusbar(BaseStationTitle, 1)
        else:
            try:
                PillCapTitleList = CapIO.Ping(GlobalSerialPort)
                # print PillCapTitleList
                if len(PillCapTitleList) < 3:
                    print("Response Length: " + str(len(PillCapTitleList)))
                    print("No Cap")
                    CapTitle = 'No Cap'
                    CapStatusStr = ''
                else:
                    
                    CapTitle = PillCapTitleList[1] + ' ' + PillCapTitleList[2]
                    CapStatusStr = CapStatus[PillCapTitleList[3][:1]]

                    
                    print(CapTitle + ": " + CapStatusStr)
                    #TODO
                    BatteryLife = 100
                    #if BatteryLife > 0:
                    #    BatteryLife = (MAX_BATTERY_LIFE - BatteryLife) / MAX_BATTERY_LIFE * 100
            except:
                CapTitle = 'Read Failure'
                CapStatusStr = ''

            self.Update_Statusbar(CapTitle, 2)
            self.Update_Statusbar(CapStatusStr, 3)
            self.Update_Statusbar(BatteryLife, 4)

    def on_ExitButton(self, event):
        for filename in glob.glob('Report *.txt'):
            os.remove(filename)
        self.Close(True)


if __name__ == '__main__':
    app = wx.App(False)
    frame = MainWindow()
    frame.Show()
    app.MainLoop()
