from win32com.client import *
from win32com.client.connect import *
import os
import sys
import subprocess
import time
import threading

print('start!')


application = win32com.client.DispatchEx("CANoe.Application")
ver = application.Version
print('Loaded CANoe version ',
    ver.major, '.',
    ver.minor, '.',
    ver.Build, '...')#, sep,''
Measurement = application.Measurement.Running
print(Measurement)

application.Visible = 1

#print('Loaded version ',
            #ver.major, '.',
            #ver.minor, '.',
#            ver.Build, '...')
