"""
general purpose com subroutines
"""
# pylint: disable-msg=C0103

import os
import serial
import sys
import platform
import glob
import os.path
from serial import Serial
import serial.tools.list_ports


def RxdStrParse(RecievedString):
    """pull out values from string separated by :'s and put into list"""
    print("Parsing: " + RecievedString)
    ReturnList = []
    MoreDelimiters = True
    FirstDelimLoc = RecievedString.find(':')
    ReturnList.append(RecievedString[0:FirstDelimLoc])
    while MoreDelimiters:
        SecondDelimLoc = RecievedString.find(':', FirstDelimLoc+1)
        if not (SecondDelimLoc > 0):
            MoreDelimiters = False
        ReturnList.append(RecievedString[FirstDelimLoc+1:SecondDelimLoc])
        FirstDelimLoc = SecondDelimLoc
    return ReturnList

def PortScan():
    """scan for available ports. return a list of ints of open ports"""
    available = []
    if platform.system() == "Windows":
        # Scan for available ports.
        for i in range(256):
            try:
                s = serial.Serial(str(i))
                available.append(int(i))
                s.close()
            except serial.SerialException:
                pass
        return available
    elif platform.system() == "Darwin": # Mac
        for dev in glob.glob('/dev/tty.usb*'):
            try:
                port = Serial(dev)
            except:
                pass
            else:
                available.append(dev)
    else:# Assume Linux of Unix
        from serial.tools.list_ports_posix import comports
        if os.path.exists('/dev/serial/by-id'):
            entries = os.listdir('/dev/serial/by-id')
            dirs = [os.readlink(os.path.join('/dev/serial/by-id', x)) for x in entries]
            # available.extend([os.path.normpath(os.path.join('/dev/serial/by-id', x)) for x in dirs])
        for dev in glob.glob('/dev/ttyUSB*'):
            try:
                port = Serial(dev)
            except:
                pass
            else:
                available.append(dev)
    return available
