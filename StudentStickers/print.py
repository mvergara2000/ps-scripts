
import os
import win32com.client as win32
hn=os.environ["USERNAME"]
labelCom = win32.Dispatch('Dymo.DymoAddIn')

#print(labelCom.GetDymoPrinters())

