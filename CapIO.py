import serial
import datetime
from time import sleep

EEPROM_WRITE_TIME = .0034 * 20
NO_BASE_STATION_TEXT = 'Base Station Not Found'
BASE_BAUD_RATE = 600
CAP_BAUD_RATE = 1200
DEFAULT_SERIAL_TIMEOUT = 0.3
NO_COM_PORT_TEXT = 'No Com Port'

SettingsDict = {'Facility': '0', 'Doctor': '1', 'Treatment': '2', 'Client': '3', 'Patient': '4', 'DosePattern': '5',
                'DoseCount': '6', 'BuzzerEnable': '7'}

def RxdStrParse(RecievedString):
    """pull out values from string separated by :'s and put into list"""
    ReturnList = []
    MoreDelimiters = True
    FirstDelimLoc = RecievedString.find(':')
    ReturnList.append(str(RecievedString[0:FirstDelimLoc]))
    while MoreDelimiters:
        SecondDelimLoc = RecievedString.find(':', FirstDelimLoc+1)
        if not (SecondDelimLoc > 0):
            MoreDelimiters = False
        ReturnList.append(str(RecievedString[FirstDelimLoc+1:SecondDelimLoc]))
        FirstDelimLoc = SecondDelimLoc
    return ReturnList

def Ping(SPort):
    print("Pinging for cap on: " + SPort)    
    COMPort = serial.Serial(SPort, CAP_BAUD_RATE, timeout=DEFAULT_SERIAL_TIMEOUT)
    COMPort.write(str.encode('T\r'))

    ResponseList = RxdStrParse(str(COMPort.readline()))
    print("Response: " + str(ResponseList))


    COMPort.close()
    return ResponseList


def ReadTime(SPort, TimeSelect):
    print("Reading Time...")

    COMPort = serial.Serial(SPort, CAP_BAUD_RATE, timeout=DEFAULT_SERIAL_TIMEOUT)
    COMPort.write(str.encode('T\r'))
    COMPort.readline()
    COMPort.write(str.encode(TimeSelect))
    ResponseString = str(COMPort.readline().strip())
    
    try:
        # print ResponseString
        print("Raw DT: " + str(ResponseString)[7:-1])
        StartDT = datetime.datetime.strptime(str(ResponseString)[7:-1], "%S:%M:%H:%d:0%w:%m:%y")
        ReturnString = StartDT.strftime("%Y-%m-%d %H:%M:%S")
        print("Parsed Date: " + ReturnString)
    except ValueError:
        ReturnString = 'Read Error'
    except KeyError:
        print("Key error formating time???")
    COMPort.close()
    print("Formated DT: " + ReturnString)
    return ReturnString


def ReadBatteryAge(SPort):
    print("Reading Battery Age")
    '''Return the battery age in days'''
    COMPort = serial.Serial(SPort, CAP_BAUD_RATE, timeout=DEFAULT_SERIAL_TIMEOUT)
    COMPort.write(str.encode('T\r'))
    COMPort.readline()
    COMPort.write(str.encode('M\r'))
    ResponseString = str(COMPort.readline().strip())
    RetVal = 0
    try:
        # print ResponseString
        BatteryInstallDate = datetime.datetime.strptime(ResponseString[4:-2], "%S:%M:%H:%d:0%w:%m:%y")
        Now = datetime.datetime.now()
        TimeDiff = Now - BatteryInstallDate
        RetVal = TimeDiff.days + TimeDiff.seconds / 86400.0
    except ValueError:
        RetVal = -1
    finally:
        pass
    COMPort.close()
    return RetVal


def ReadSettings(SPort):
    print("Reading Settings...")
    ReturnDict = {}
    COMPort = serial.Serial(SPort, CAP_BAUD_RATE, timeout=DEFAULT_SERIAL_TIMEOUT)
    COMPort.write(str.encode('T\r'))
    COMPort.readline()
    for Description, Setting in SettingsDict.items():
        try:
            print(Description + " : " + Setting)
            COMPort.write(str.encode('G'))
            COMPort.write(str.encode(str(Setting)))
            COMPort.write(str.encode('\r'))
            ResponseList = RxdStrParse(str(COMPort.readline().strip()))
            
            ReturnDict[Description] = str(ResponseList[2].strip())
        except TypeError as te:
            print(te)
        except ValueError as ve:
            print(ve)
    COMPort.close()
    return ReturnDict


def ReadData(SPort):
    print("Reading Data...")
    ReturnStringList = []
    COMPort = serial.Serial(SPort, CAP_BAUD_RATE, timeout=DEFAULT_SERIAL_TIMEOUT)
    COMPort.write(str.encode('T\r'))
    COMPort.readline()
    COMPort.write(str.encode('D\r'))

    ResponseData = str(COMPort.readline().strip())
    print("Data: " + ResponseData)
    ResponseList = RxdStrParse(ResponseData)
    
    print("Data Size: " + ResponseList[1])
    for i in range(0, int(ResponseList[1].strip())):
        ReturnStringList.append(str(COMPort.readline().strip()))
    COMPort.close()
    return ReturnStringList


def WriteString(SPort, OutputString):
    print("Writing String: " + OutputString)
    COMPort = serial.Serial(SPort, CAP_BAUD_RATE, timeout=DEFAULT_SERIAL_TIMEOUT)
    OutputChars = list(OutputString)
    # print OutputChars
    for Char in OutputChars:
        COMPort.write(str.encode(str(Char)))
        sleep(0.02)
    # print Char,
    # print
    COMPort.close()


def WriteSettings(SPort, WriteDict):
    Ping(SPort)
    for Description, Settings in SettingsDict.items():
        OutputString = WriteDict[Description]
        OutputString = OutputString[:20]  # limit strings to 20 characters
        print("Writing Setting: " + OutputString)
        WriteString(SPort, 'S' + SettingsDict[Description] + OutputString + '\r')
        sleep(EEPROM_WRITE_TIME)


def Erase(SPort):
    print("Erasing...")
    try:
        COMPort = serial.Serial(SPort, CAP_BAUD_RATE, timeout=3)
        COMPort.write(str.encode('ER\r'))
        EraseReturnString = ''
        EraseReturnString = COMPort.readline()
        COMPort.close()
        RetVal = 0
        if not EraseReturnString == '':
            RetVal = 1
    except:
        RetVal = 0
    return RetVal


def FindBaseStation():
    print("Finding Basestation...")
    # scan for open com ports
    #	if no port return Fail
    # starting with port from INI file if present and open
    # 		send 'z\r\n' to port
    #		wait for response
    #		if not found, next port
    #		if found then return OK
    #		if tried all ports return Fail
    # Error messages begin with 'No', for checking later
    ReturnList = []
    ResponseString = ''
    
    scanResult = serial.tools.list_ports.comports()

    PortList = []
    
    for portInfo in scanResult:
        print(portInfo.device)
        PortList.append(portInfo.device)
    
    if not len(PortList):
        PortName = NO_COM_PORT_TEXT
    else:  # search through open ports, will return last one that responds
        PortName = NO_BASE_STATION_TEXT + '.\nCheck cables.\nBe sure FTDI230X driver is installed.\nRemove from bright lights.\nInserting a cap may help.'
        for Port in PortList:
            try:
                COMPort = serial.Serial(Port, BASE_BAUD_RATE, timeout=DEFAULT_SERIAL_TIMEOUT)
                COMPort.write(str.encode('T'))
                sleep(0.05)
                COMPort.write(str.encode('\r'))
                print(COMPort.readline())
                COMPort.write(str.encode('Z'))
                sleep(0.05)
                COMPort.write(str.encode('\r'))
                ResponseString = COMPort.readline()
                print("Response String: " + str(ResponseString))
                COMPort.close()
                if 'Base' in str(ResponseString):
                    print("Basestation found on port " + str(Port))
                    PortName = Port
            finally:
                pass
    ReturnList.append(str(PortName))
    ReturnList.append(str(ResponseString))
    return ReturnList
