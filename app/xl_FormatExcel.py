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
fillStyle = IN[3]
textStyle = IN[4]
borderStyle = IN[5]
cellRange = IN[6]

def LiveStream():
	try:
		xlApp = Marshal.GetActiveObject("Excel.Application")
		xlApp.Visible = True
		xlApp.DisplayAlerts = False
		return xlApp
	except:
		return None
			
def FormatCells1(fillStyle=None, textStyle=None, borderStyle=None, ws=None):

	def FormatData1(x=None, y=None, x1=None, y1=None, fillStyle=None, textStyle=None, borderStyle=None, ws=None):
		if fillStyle != None:
			fillSettings = fillStyle[x1][y1]
			patternType = bb.ParseFillSettings(fillSettings)[0]
			backColor = bb.ParseFillSettings(fillSettings)[1]
			foreColor = bb.ParseFillSettings(fillSettings)[2]
			ws.Cells[x+1, y+1].Interior.Pattern = patternType
			ws.Cells[x+1, y+1].Interior.PatternColor = foreColor
			ws.Cells[x+1, y+1].Interior.Color = backColor
		if borderStyle != None:
			ws.Cells[x+1, y+1].BorderAround(bb.GetLineStyle(borderStyle[x1][y1]))
		if textStyle != None:
			textSettings = textStyle[x1][y1]
			ws.Cells[x+1, y+1].Font.Name = bb.ParseTextStyle(textSettings)[0]
			ws.Cells[x+1, y+1].Font.Size = bb.ParseTextStyle(textSettings)[1]
			ws.Cells[x+1, y+1].Font.Color = bb.ParseTextStyle(textSettings)[2] 
			ws.Cells[x+1, y+1].HorizontalAlignment = bb.ParseTextStyle(textSettings)[3]
			ws.Cells[x+1, y+1].VerticalAlignment = bb.ParseTextStyle(textSettings)[4]
			ws.Cells[x+1, y+1].Font.Bold = bb.ParseTextStyle(textSettings)[5]
			ws.Cells[x+1, y+1].Font.Italic = bb.ParseTextStyle(textSettings)[6]
			ws.Cells[x+1, y+1].Font.Underline = bb.ParseTextStyle(textSettings)[7]
			ws.Cells[x+1, y+1].Font.Strikethrough = bb.ParseTextStyle(textSettings)[8]
		return ws

	if fillStyle != None:				
		for i in range(0, len(fillStyle), 1):
			for j in range(0, len(fillStyle[0]), 1):
				FormatData1(i, j, i, j, fillStyle, textStyle, borderStyle, ws)
	elif textStyle != None:
		for i in range(0, len(textStyle), 1):
			for j in range(0, len(textStyle[0]), 1):
				FormatData1(i, j, i, j, fillStyle, textStyle, borderStyle, ws)
	else:
		for i in range(0, len(borderStyle), 1):
			for j in range(0, len(borderStyle[0]), 1):
				FormatData1(i, j, i, j, fillStyle, textStyle, borderStyle, ws)
	return ws

def FormatCells(origin=None, extent=None, fillStyle=None, textStyle=None, borderStyle=None, ws=None):
	if fillStyle != None:
		patternType = bb.ParseFillSettings(fillStyle)[0]
		backColor = bb.ParseFillSettings(fillStyle)[1]
		foreColor = bb.ParseFillSettings(fillStyle)[2]
		ws.Range[origin, extent].Interior.Pattern = patternType
		ws.Range[origin, extent].Interior.PatternColor = foreColor
		ws.Range[origin, extent].Interior.Color = backColor
	if borderStyle != None:
		ws.Range[origin, extent].BorderAround(bb.GetLineStyle(borderStyle))
	if textStyle != None:
		ws.Range[origin, extent].Font.Name = bb.ParseTextStyle(textStyle)[0]
		ws.Range[origin, extent].Font.Size = bb.ParseTextStyle(textStyle)[1]
		ws.Range[origin, extent].Font.Color = bb.ParseTextStyle(textStyle)[2] 
		ws.Range[origin, extent].HorizontalAlignment = bb.ParseTextStyle(textStyle)[3]
		ws.Range[origin, extent].VerticalAlignment = bb.ParseTextStyle(textStyle)[4]
		ws.Range[origin, extent].Font.Bold = bb.ParseTextStyle(textStyle)[5]
		ws.Range[origin, extent].Font.Italic = bb.ParseTextStyle(textStyle)[6]
		ws.Range[origin, extent].Font.Underline = bb.ParseTextStyle(textStyle)[7]
		ws.Range[origin, extent].Font.Strikethrough = bb.ParseTextStyle(textStyle)[8]
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
			
			if cellRange != None and not bb.CheckInputs(fillStyle, textStyle, borderStyle):
				#message = "cellRange + no nested inp"
				origin = ws.Cells(bb.xlRange(cellRange)[1], bb.xlRange(cellRange)[0])
				extent = ws.Cells(bb.xlRange(cellRange)[3], bb.xlRange(cellRange)[2])
				FormatCells(origin, extent, fillStyle, textStyle, borderStyle, ws)
				Marshal.ReleaseComObject(extent)
				Marshal.ReleaseComObject(origin)
			elif cellRange == None and not bb.CheckInputs(fillStyle, textStyle, borderStyle):
				message = "cellRange is None but no nested inp"
				origin = ws.Cells(ws.UsedRange.Row, ws.UsedRange.Column)
				extent = ws.Cells(ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row, ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column)
				FormatCells(origin, extent, fillStyle, textStyle, borderStyle, ws)
				Marshal.ReleaseComObject(extent)
				Marshal.ReleaseComObject(origin)
			elif cellRange == None and bb.CheckInputs(fillStyle, textStyle, borderStyle):
				#message = "cellRange is None and inputs are nested"
				FormatCells1(fillStyle, textStyle, borderStyle, ws)
			else:
				message = "Range and Style List cannot be combined. Please either use Range with single item Styles or styles as list input"

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
