#Copyright(c) 2015, David Mans, Konrad Sobon
# @arch_laboratory, http://archi-lab.net, http://neoarchaic.net

import clr
import sys

pyt_path = r'C:\Program Files (x86)\IronPython 2.7\Lib'
sys.path.append(pyt_path)

import os
import os.path
appDataPath = os.getenv('APPDATA')
bbPath = appDataPath + r"\Dynamo\0.8\packages\Bumblebee\extra"
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
cellRange = IN[3]
graphicStyle = IN[4]

def LiveStream():
	try:
		xlApp = Marshal.GetActiveObject("Excel.Application")
		xlApp.Visible = True
		xlApp.DisplayAlerts = False
		return xlApp
	except:
		return None
			
def FormatCells1(ws=None, graphicStyle=None):

	def FormatData1(x=None, y=None, x1=None, y1=None, ws=None, graphicStyle=None):
		if graphicStyle.fillStyle != None:
			fillSettings = fillStyle[x1][y1]
			patternType = bb.ParseFillSettings(fillSettings)[0]
			backColor = bb.ParseFillSettings(fillSettings)[1]
			foreColor = bb.ParseFillSettings(fillSettings)[2]
			ws.Cells[x+1, y+1].Interior.Pattern = patternType
			ws.Cells[x+1, y+1].Interior.PatternColor = foreColor
			ws.Cells[x+1, y+1].Interior.Color = backColor
		if graphicStyle.textStyle != None:
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
		if graphicStyle.borderStyle != None:
			borderSettings = borderStyle[x1][y1]
			lineStyle = bb.ParseBorderStyle(borderSettings)[0]
			lineWeight = bb.ParseBorderStyle(borderSettings)[1]
			lineColor = bb.ParseBorderStyle(borderSettings)[2]
			ws.Cells[x+1, y+1].BorderAround(lineStyle, lineWeight, lineColor)
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

def FormatCells(origin=None, extent=None, ws=None, graphicStyle=None):
	if graphicStyle.fillStyle != None:
		fillStyle = graphicStyle.fillStyle
		ws.Range[origin, extent].Interior.Pattern = fillStyle.PatternType()
		ws.Range[origin, extent].Interior.PatternColor = fillStyle.PatternColor()
		ws.Range[origin, extent].Interior.Color = fillStyle.BackgroundColor()
	if graphicStyle.textStyle != None:
		textStyle = graphicStyle.textStyle
		ws.Range[origin, extent].Font.Name = textStyle.Name()
		ws.Range[origin, extent].Font.Size = textStyle.Size()
		ws.Range[origin, extent].Font.Color = textStyle.Color()
		ws.Range[origin, extent].HorizontalAlignment = textStyle.HorizontalAlign()
		ws.Range[origin, extent].VerticalAlignment = textStyle.VerticalAlign()
		ws.Range[origin, extent].Font.Bold = textStyle.Bold()
		ws.Range[origin, extent].Font.Italic = textStyle.Italic()
		ws.Range[origin, extent].Font.Underline = textStyle.Underline()
		ws.Range[origin, extent].Font.Strikethrough = textStyle.Strikethrough()
	if graphicStyle.borderStyle != None:
		borderStyle = graphicStyle.borderStyle
		ws.Range[origin, extent].BorderAround(borderStyle.LineType(), borderStyle.Weight(), borderStyle.Color())
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
			
			if cellRange != None and not isinstance(graphicStyle, list):
				origin = ws.Cells(bb.xlRange(cellRange)[1], bb.xlRange(cellRange)[0])
				extent = ws.Cells(bb.xlRange(cellRange)[3], bb.xlRange(cellRange)[2])
				FormatCells(origin, extent, ws, graphicStyle)
				Marshal.ReleaseComObject(extent)
				Marshal.ReleaseComObject(origin)
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
