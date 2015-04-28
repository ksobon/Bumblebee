#Copyright(c) 2015, David Mans, Konrad Sobon
# @arch_laboratory, http://archi-lab.net, http://neoarchaic.net

import clr
import sys

pyt_path = r'C:\Program Files (x86)\IronPython 2.7\Lib'
sys.path.append(pyt_path)

import os
import os.path
appDataPath = os.getenv('APPDATA')
bbPath = appDataPath + r"\Dynamo\0.7\packages\Bumblebee\extra"
if bbPath not in sys.path:
	sys.path.Add(bbPath)

import System
from System import Array
from System.Collections.Generic import *

clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel

System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo("en-US")
from System.Runtime.InteropServices import Marshal

import bumblebee as bb
import string
import re

#The inputs to this node will be stored as a list in the IN variable.
dataEnteringNode = IN

filePath = IN[0]
runMe = IN[1]
sheetName = IN[2]
cellFill = IN[3]
borderStyle = IN[4]
cellRange = IN[5]

def LiveStream():
	try:
		xlApp = Marshal.GetActiveObject("Excel.Application")
		xlApp.Visible = True
		xlApp.DisplayAlerts = False
		return xlApp
	except:
		return None

def ParseFillSettings(cellFill):
	paramList = cellFill.split("~")
	patternType = bb.GetPatternType(paramList[0])
	# if pattern type is supplied then it needs a background color
	# to be set so if no background color supplied it will be assigned 
	# a default value of white
	if paramList[1] == "xlNone" and paramList[0] == "xlNone":
		backColor = -4142
	elif paramList[1] == "xlNone" and paramList[0] != "xlNone":
		backColor = bb.RGBToRGBLong((255,255,255))
	else:
		bColors = paramList[1].split(",")
		backColor = bb.RGBToRGBLong((int(bColors[2]), int(bColors[1]), int(bColors[0])))
	if paramList[2] == "xlNone" and paramList[0] == "xlNone":
		foreColor = -4142
	elif paramList[2] == "xlNone" and paramList[0] != "xlNone":
		foreColor = bb.RGBToRGBLong((0,0,0))
	else:
		fColors = paramList[2].split(",")
		foreColor = bb.RGBToRGBLong((int(fColors[2]), int(fColors[1]), int(fColors[0])))
	return [patternType, backColor, foreColor]
			
def FormatCells1(cellFill=None, borderStyle=None, ws=None):

	def FormatData1(x=None, y=None, x1=None, y1=None, cellFill=None, borderStyle=None, ws=None):
		if cellFill != None:
			fillSettings = cellFill[x1][y1]
			patternType = ParseFillSettings(fillSettings)[0]
			backColor = ParseFillSettings(fillSettings)[1]
			foreColor = ParseFillSettings(fillSettings)[2]
			ws.Cells[x+1, y+1].Interior.Pattern = patternType
			ws.Cells[x+1, y+1].Interior.PatternColor = foreColor
			ws.Cells[x+1, y+1].Interior.Color = backColor
		if borderStyle != None:
			ws.Cells[x+1, y+1].BorderAround(bb.GetLineStyle(borderStyle[x1][y1]))
		return ws
					
	for i in range(0, len(cellFill), 1):
		for j in range(0, len(cellFill[0]), 1):
			FormatData1(i, j, i, j, cellFill, borderStyle, ws)
	return ws

def FormatCells(origin=None, extent=None, cellFill=None, borderStyle=None, ws=None):
	if cellFill != None:
		patternType = ParseFillSettings(cellFill)[0]
		backColor = ParseFillSettings(cellFill)[1]
		foreColor = ParseFillSettings(cellFill)[2]
		ws.Range[origin, extent].Interior.Pattern = patternType
		ws.Range[origin, extent].Interior.PatternColor = foreColor
		ws.Range[origin, extent].Interior.Color = backColor
	if borderStyle != None:
		ws.Range[origin, extent].BorderAround(bb.GetLineStyle(borderStyle))
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
			
			if cellRange != None and not any(isinstance(item, list) for item in cellFill):
				origin = ws.Cells(bb.xlRange(cellRange)[1], bb.xlRange(cellRange)[0])
				extent = ws.Cells(bb.xlRange(cellRange)[3], bb.xlRange(cellRange)[2])
				FormatCells(origin, extent, cellFill, borderStyle, ws)
				Marshal.ReleaseComObject(extent)
				Marshal.ReleaseComObject(origin)
			elif cellRange == None and not any(isinstance(item, list) for item in cellFill):
				origin = ws.Cells(ws.UsedRange.Row, ws.UsedRange.Column)
				extent = ws.Cells(ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row, ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column)
				FormatCells(origin, extent, cellFill, borderStyle, ws)
				Marshal.ReleaseComObject(extent)
				Marshal.ReleaseComObject(origin)
			elif cellRange == None and any(isinstance(item, list) for item in cellFill):
				FormatCells1(cellFill, borderStyle, ws)

			wb.SaveAs(str(filePath))
			xlApp.ActiveWorkbook.Close(False)
			
			xlApp.ScreenUpdating = True
			Marshal.ReleaseComObject(ws)
			Marshal.ReleaseComObject(wb)
			Marshal.ReleaseComObject(xlApp)
		else:
			message = "No file exists. Please Use Write Data Node to create file first before formatting it."
	else:
		message = "Close currently running Excel \nsession."
else:
	message = "Run Me is set to False. Please set \nto True if you wish to write data \nto Excel."

#Assign your output to the OUT variable
if message == None:
	OUT = "Success!"
else:
	OUT = '\n'.join('{:^35}'.format(s) for s in message.split('\n'))
