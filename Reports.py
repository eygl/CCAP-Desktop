""" Generate reports.  XL, Text and raw data formats
Version Date        Changes

"""
import os.path
import stat
from datetime import datetime
from xlwt import Workbook, Borders, Alignment, XFStyle
import ComSubs
import webbrowser
import platform
import subprocess

# pylint: disable-msg=C0103
# pylint: disable-msg=C0301

DosePatternsRev = {'2/day': '12', '1/day': '24', '1/2 days': '48', '1/3 days': '72', '1/week': '168', '1/month': '720',
                   '1/3 months': '2160'}


def DataSubStringParse(RecievedSubstring):
    """pull out values from string separated by ','s and put into list"""
    ReturnSubList = []
    FirstDelimLoc = RecievedSubstring.find(',')
    ReturnSubList.append(RecievedSubstring[0:FirstDelimLoc])
    SecondDelimLoc = RecievedSubstring.find(',', FirstDelimLoc + 1)
    ReturnSubList.append(RecievedSubstring[FirstDelimLoc + 1:SecondDelimLoc])
    ThirdDelimLoc = RecievedSubstring.find(',', SecondDelimLoc + 1)
    ReturnSubList.append(RecievedSubstring[SecondDelimLoc + 1:ThirdDelimLoc])
    ReturnSubList.append(RecievedSubstring[ThirdDelimLoc + 1:])
    return ReturnSubList


def DataStringParse(RecievedString):
    """pull out values from string separated by ';'s and put into list"""
    ReturnList = []
    MoreDelimiters = True
    if not RecievedString == None:
        FirstDelimLoc = RecievedString.find(';')
        ReturnList.append(DataSubStringParse(RecievedString[0:FirstDelimLoc]))
        while MoreDelimiters:
            SecondDelimLoc = RecievedString.find(';', FirstDelimLoc + 1)
            if SecondDelimLoc < 0:
                MoreDelimiters = False
            else:
                ReturnList.append(DataSubStringParse(RecievedString[FirstDelimLoc + 1:SecondDelimLoc]))
            FirstDelimLoc = SecondDelimLoc
    return ReturnList


def GetReportDestination(ReportRootDirectory):
    """ generate folder based date and sub folder on sequential numbering"""
    # first create folders using today's date
    t = datetime.now()
    # if path doesn't exsist create it
    if not os.path.isdir(ReportRootDirectory + t.strftime("\%Y-%m-%d")):
        os.mkdir(ReportRootDirectory + t.strftime("\%Y-%m-%d"))
    # look in folder for folders with numbers, find highest number, add one,
    # thats our folder for the reports
    i = 0
    while os.path.isdir(ReportRootDirectory + t.strftime("\%Y-%m-%d\\") + str(i)):
        i += 1
    SaveDirectory = ReportRootDirectory + t.strftime("\%Y-%m-%d\\") + str(i)
    os.mkdir(SaveDirectory)
    return SaveDirectory


def GenerateReport(SaveDirectory, ReportType, Info, Data):
    """ send data and patient info to the selected report type generator """
    if ReportType == 'Excel':
        GenerateXLReport(SaveDirectory, Info, Data)
    else:  # Text
        GenerateTextReport(SaveDirectory, Info, Data)


def GenerateXLReport(Directory, Info, Data):
    """  create XL report """
    t = datetime.now()
    DateTimeNowStr = t.strftime("%Y-%m-%d %H:%M")

    wb = Workbook()
    ws = []

    N = 0  # number of controls

    if Info['Patient'] == '':
        ws.append(wb.add_sheet('Patient'))
    else:
        ws.append(wb.add_sheet(Info['Patient']))

    borders = Borders()
    borders.left = Borders.THIN
    borders.right = Borders.THIN
    borders.top = Borders.THIN
    borders.bottom = Borders.THIN
    aligncenter = Alignment()
    aligncenter.horz = Alignment.HORZ_CENTER
    BorderStyle = XFStyle()
    BorderStyle.borders = borders
    BorderStyle.alignment = aligncenter

    TableHeaderStyle = XFStyle()
    TableHeaderStyle.alignment = aligncenter

    alignleft = Alignment()
    alignleft.horz = Alignment.HORZ_LEFT
    LeftDateStyle = XFStyle()
    LeftDateStyle.num_format_str = 'M/D/YY'
    LeftDateStyle.alignment = alignleft

    LeftStyle = XFStyle()
    LeftStyle.alignment = alignleft

    alignright = Alignment()
    alignright.horz = Alignment.HORZ_RIGHT
    RightStyle = XFStyle()
    RightStyle.alignment = alignright

    # column 'A'
    ws[N].col(0).width = 4000
    ws[N].write(0, 0, 'cCap Compliance Report')
    ws[N].write(1, 0, 'Report Date:')
    ws[N].write(2, 0, "Facility:")
    ws[N].write(3, 0, "Doctor:")
    ws[N].write(4, 0, 'Client:')
    ws[N].write(5, 0, 'Patient:')
    ws[N].write(7, 0, "Treatment:")
    ws[N].write(8, 0, 'Dose Frequency')
    ws[N].write(9, 0, 'Dose Count:')
    ws[N].write(10, 0, 'Compliance Rate:')
    ws[N].write(12, 0, 'Time Points:')
    ws[N].write(13, 0, 'Month')

    # column 'B'
    ws[N].col(1).width = 4000
    ws[N].write(1, 1, DateTimeNowStr)
    ws[N].write(2, 1, Info['Facility'])
    ws[N].write(3, 1, Info['Doctor'])
    ws[N].write(4, 1, Info['Client'])
    ws[N].write(5, 1, Info['Patient'])
    ws[N].write(7, 1, Info['Treatment'])
    ws[N].write(8, 1, Info['DosePattern'])
    ws[N].write(9, 1, Info['DoseCount'])

    # calculate compliance
    DataList = DataStringParse(Data)
    DosesTaken = len(DataList)
    for DataSubList in DataList:
        if int(DataSubList[0]) == 0 and int(DataSubList[1]) == 0 and int(DataSubList[2]) == 0 and int(
                DataSubList[3]) == 0:
            DosesTaken -= 1

    if int(Info['DoseCount']) == 0:
        ws[N].write(10, 1, 'No Data')
    else:
        ws[N].write(10, 1, '{0:.1f} %'.format(100 * DosesTaken / int(Info['DoseCount'])))
    ws[N].write(13, 1, 'Day')

    ws[N].col(2).width = 4000
    ws[N].write(13, 2, 'Hour')
    ws[N].col(3).width = 4000
    ws[N].write(13, 3, 'Minute')

    i = 0
    DataList = DataStringParse(Data)
    for DataSubList in DataList:
        if int(DataSubList[2]) > 12:  # convert to AM/PM
            DataSubList[2] = str(int(DataSubList[2]) - 12) + 'PM'
        else:
            DataSubList[2] = DataSubList[2] + 'AM'
        ws[N].write(14 + i, 0, str(DataSubList[0]))
        ws[N].write(14 + i, 1, str(DataSubList[1]))
        ws[N].write(14 + i, 2, str(DataSubList[2]))
        ws[N].write(14 + i, 3, str(DataSubList[3]))
        i = i + 1

    i = 0
    if len(Directory) > 0:
        Directory += "/"
    # Filename = Directory + "/" + Info['Patient'] + '-'  + Info['Treatment'] + t.strftime("-%Y-%m-%d") + '.xls'
    Filename = Directory + "Report " + Info['Patient'] + '-' + Info['Treatment'] + t.strftime("-%Y-%m-%d") + '.xls'
    while os.path.exists(Filename):
        i = i + 1
        Filename = Directory + "Report " + Info['Patient'] + ' ' + Info['Treatment'] + ' ' + t.strftime("-%Y-%m-%d-") + str(
            i) + '.xls'

    wb.save(Filename)
    os.chmod(Filename, stat.S_IREAD)
    webbrowser.open(Filename)


def GenerateTextReport(Directory, Info, Data):
    """ create TXT report """
    t = datetime.now()
    DateTimeNowStr = t.strftime("%Y-%m-%d %H:%M")

    i = 0
    if len(Directory) > 0:
        Directory += "/"
    # Filename = Directory + "/" + Info['Patient'] + ' ' + Info['Treatment'] + ' ' + t.strftime("-%Y-%m-%d-") + '.txt'
    Filename = Directory + "Report " + Info['Patient'] + '-' + Info['Treatment'] + t.strftime("-%Y-%m-%d") + '.txt'
    while os.path.exists(Filename):
        i = i + 1
        Filename = Directory + "Report " + Info['Patient'] + ' ' + Info['Treatment'] + ' ' + t.strftime("-%Y-%m-%d-") + str(
            i) + '.txt'

    f = open(Filename, 'w')
    f.write('cCap COMPLIANCE REPORT')
    f.write('\n--------------------------------------------------------')
    f.write('\nReport Date:    ' + DateTimeNowStr)
    f.write('\nFacility:       ' + Info['Facility'])
    f.write('\nDoctor:         ' + Info['Doctor'])
    f.write('\nClient:         ' + Info['Client'])
    f.write('\nPatient:        ' + Info['Patient'])

    f.write('\n--------------------------------------------------------')
    f.write('\nTreatment:      ' + Info['Treatment'])
    f.write('\nDose Frequency: ' + Info['DosePattern'])
    f.write('\nDose Count:     ' + str(Info['DoseCount']))

    Doses = int(Info['DoseCount'])
    if 'Week' in Info['DosePattern']:
        DoseLengthUnits = ' Weeks'
        DoseLengthMultiplier = int(DosePatternsRev[Info['DosePattern']]) / 168
    elif 'Month' in Info['DosePattern']:
        DoseLengthUnits = ' Months'
        DoseLengthMultiplier = int(DosePatternsRev[Info['DosePattern']]) / 720
    else:
        DoseLengthUnits = ' Days'
        DoseLengthMultiplier = float(DosePatternsRev[Info['DosePattern']]) / 24
        if Doses % 2 != 0 and DoseLengthMultiplier == 12:
            Doses += 1
    f.write('\nDose Period:    ' + str(int(Doses * DoseLengthMultiplier)) + DoseLengthUnits)

    # calculate compliance
    DataList = DataStringParse(Data)
    DosesTaken = len(DataList)
    for DataSubList in DataList:
        if int(DataSubList[0]) == 0 and int(DataSubList[1]) == 0 and int(DataSubList[2]) == 0 and int(
                DataSubList[3]) == 0:
            DosesTaken -= 1
    if int(Info['DoseCount']) == 0:
        f.write('\nCompliance Rate:No Data')
    else:
        f.write('\nCompliance Rate:{0:.2f} %'.format(100 * DosesTaken / int(Info['DoseCount'])))

    f.write('\n')
    f.write('\nTime Points:')
    f.write('\nMon  Day  Hr   Min\n')

    for DataSubList in DataList:
        if int(DataSubList[0]) == 0 and int(DataSubList[1]) == 0 and int(DataSubList[2]) == 0 and int(
                DataSubList[3]) == 0:
            f.write('Missed Dose\n')
        else:
            if int(DataSubList[2]) > 12:  # convert to AM/PM
                DataSubList[2] = str(int(DataSubList[2]) - 12) + 'PM'
            else:
                DataSubList[2] = DataSubList[2] + 'AM'
            f.write(str(DataSubList[0]) + '   ' + str(DataSubList[1]) + '    ' + str(DataSubList[2]) + '   ' + str(
                DataSubList[3]) + '\n')

    f.close()
    #os.chmod(Filename, stat.S_IREAD)
    if platform.system() == "Darwin":
        subprocess.call(['open', '-a', 'TextEdit', Filename])
    elif platform.system() == "Windows":
        webbrowser.open(Filename)
    else:
        webbrowser.open(Filename)
