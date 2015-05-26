#Copyright(c) 2015, David Mans, Konrad Sobon
# @arch_laboratory, http://archi-lab.net, http://neoarchaic.net

import clr
import sys
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

import System
from System import Array
from System.Collections.Generic import *

pyt_path = r'C:\Program Files (x86)\IronPython 2.7\Lib'
sys.path.append(pyt_path)

import os.path
import os

appDataPath = os.getenv('APPDATA')
bbPath = appDataPath + r"\Dynamo\0.8\packages\Bumblebee\extra"
if bbPath not in sys.path:
	sys.path.Add(bbPath)

import bumblebee as bb

clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel

System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo("en-US")
from System.Runtime.InteropServices import Marshal

#The inputs to this node will be stored as a list in the IN variable.
dataEnteringNode = IN

filePath = IN[0]
runMe = IN[1]
sheetName = IN[2]
searchValues = IN[3]

def SetUp(xlApp):
	# supress updates and warning pop ups
	xlApp.Visible = False
	xlApp.DisplayAlerts = False
	xlApp.ScreenUpdating = False
	return xlApp

def ExitExcel(filePath, xlApp, wb, ws):
	# clean up before exiting excel, if any COM object remains
	# unreleased then excel crashes on open following time
	def CleanUp(_list):
		if isinstance(_list, list):
			for i in _list:
				Marshal.ReleaseComObject(i)
		else:
			Marshal.ReleaseComObject(_list)
		return None
	
	wb.SaveAs(str(filePath))
	xlApp.ActiveWorkbook.Close(False)
	xlApp.ScreenUpdating = True
	CleanUp([ws,wb,xlApp])
	return None

if runMe:
	message = None
	xlApp = SetUp(Excel.ApplicationClass())
	
	if os.path.isfile(str(filePath)):
		xlApp.Workbooks.open(str(filePath))
		wb = xlApp.ActiveWorkbook
		ws = xlApp.Sheets(sheetName)
		
		originX = ws.UsedRange.Row
		originY = ws.UsedRange.Column
		boundX = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row
		boundY = ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column
		
		dataOut = []
		xlAfter = ws.Cells(originY, originX)
		xlLookIn = -4163
		xlLookAt = "&H2"
		xlSearchOrder = "&H1"
		xlSearchDirection = 1
		xlMatchCase = False
		xlMatchByte = False
		xlSearchFormat = False
		if isinstance(searchValues, list):
			for key in searchValues:
				cellAddress = ws.Cells.Find(key, xlAfter, xlLookIn, xlLookAt, xlSearchOrder, xlSearchDirection, xlMatchCase, xlMatchByte, xlSearchFormat).Address(False, False)
				addressX = xlApp.Range(cellAddress).Row
				addressY = xlApp.Range(cellAddress).Column
				row = ws.Range[ws.Cells(addressX, originY), ws.Cells(addressX, boundY)].Value2
				dataOut.append(row)
		else:
			cellAddress = ws.Cells.Find(searchValues, xlAfter, xlLookIn, xlLookAt, xlSearchOrder, xlSearchDirection, xlMatchCase, xlMatchByte, xlSearchFormat).Address(False, False)
			addressX = xlApp.Range(cellAddress).Row
			addressY = xlApp.Range(cellAddress).Column
			row = ws.Range[ws.Cells(addressX, originY), ws.Cells(addressX, boundY)].Value2
			dataOut = row
		ExitExcel(filePath, xlApp, wb, ws)
	else:
		message = "Specified File doesn't exist."
else:
	message = "Run Me is set to False. Please set \nto True if you wish to write data \nto Excel."

#Assign your output to the OUT variable
if message == None:
	OUT = dataOut
else:
	OUT = '\n'.join('{:^35}'.format(s) for s in message.split('\n'))
