#Copyright(c) 2015, David Mans, Konrad Sobon
# @arch_laboratory, http://archi-lab.net, http://neoarchaic.net

import clr
import sys
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

import System
from System import Array
from System.Collections.Generic import *

clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel

System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo("en-US")
from System.Runtime.InteropServices import Marshal

pyt_path = r'C:\Program Files (x86)\IronPython 2.7\Lib'
sys.path.append(pyt_path)
import os.path

#The inputs to this node will be stored as a list in the IN variable.
dataEnteringNode = IN

filePath = IN[0]
runMe = IN[1]
sheetName = IN[2]
byColumn = IN[3]
data = IN[4]

def LiveStream():
	try:
		xlApp = Marshal.GetActiveObject("Excel.Application")
		xlApp.Visible = True
		xlApp.DisplayAlerts = False
		return xlApp
	except:
		return None

def WriteData(ws, data, byColumn):

	def FillData(x=None, y=None, x1=None, y1=None, ws=None, data=None):
		if y1 != None:
			ws.Cells[x+1, y+1] = data[x1][y1]
		else:
			ws.Cells[x+1, y+1] = data[x1]
		return ws
	if any(isinstance(item, list) for item in data):
		for i, valueX in enumerate(data):
			for j, valueY in enumerate(valueX):
				if byColumn:
					FillData(j,i,i,j, ws, data)
				else:
					FillData(i,j,i,j, ws, data)
	else:
		for i, valueX in enumerate(data):
			if byColumn:
				FillData(i,0,i,None, ws, data)
			else:
				FillData(0,i,i,None, ws, data)
	return ws

if runMe:
	message = None
	if LiveStream() == None:
		xlApp = Excel.ApplicationClass()
		xlApp.Visible = False
		xlApp.DisplayAlerts = False
		xlApp.ScreenUpdating = False
		if os.path.isfile(str(filePath)):
			xlApp.Workbooks.open(str(filePath))
			wb = xlApp.ActiveWorkbook
			ws = xlApp.Sheets(sheetName)
			ws.Cells.ClearContents()
			ws.Cells.Clear()
		else:
			wb = xlApp.Workbooks.Add()
			ws = wb.Worksheets[1]
			ws.Name = sheetName
		WriteData(ws, data, byColumn)
		wb.SaveAs(str(filePath))
		xlApp.ActiveWorkbook.Close(False)
		xlApp.ScreenUpdating = True
		Marshal.ReleaseComObject(ws)
		Marshal.ReleaseComObject(wb)
		Marshal.ReleaseComObject(xlApp)
	else:
		message = "Close currently running Excel \nsession."
else:
	message = "Run Me is set to False. Please set \nto True if you wish to write data \nto Excel."

#Assign your output to the OUT variable
if message == None:
	OUT = "Success!"
else:
	OUT = '\n'.join('{:^35}'.format(s) for s in message.split('\n'))
