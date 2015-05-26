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

#The inputs to this node will be stored as a list in the IN variable.
dataEnteringNode = IN

filePath = IN[0]
runMe = IN[1]
sheetName = IN[2]
searchValues = IN[3]

if runMe:
	objExcel = Excel.ApplicationClass() 
	objExcel.Visible = False
	objExcel.DisplayAlerts = False
	objExcel.screenUpdating = False
	objExcel.Workbooks.open(str(filePath))
	
	excelWorkbook = objExcel.ActiveWorkbook
	excelSheet = objExcel.Sheets(sheetName)
	
	origin = None
	if origin == None:
		originX = excelSheet.UsedRange.Row
		originY = excelSheet.UsedRange.Column
	bound = None
	if bound == None:
		boundX = excelSheet.UsedRange.Rows(excelSheet.UsedRange.Rows.Count).Row
		boundY = excelSheet.UsedRange.Columns(excelSheet.UsedRange.Columns.Count).Column
	# test search function
	dataOut = []
	xlAfter = excelSheet.Cells(originY, originX)
	xlLookIn = -4163
	xlLookAt = "&H2"
	xlSearchOrder = "&H1"
	xlSearchDirection = 1
	xlMatchCase = False
	xlMatchByte = False
	xlSearchFormat = False
	if isinstance(searchValues, list):
		for key in searchValues:
			cellAddress = excelSheet.Cells.Find(key, xlAfter, xlLookIn, xlLookAt, xlSearchOrder, xlSearchDirection, xlMatchCase, xlMatchByte, xlSearchFormat).Address(False, False)
			addressX = objExcel.Range(cellAddress).Row
			addressY = objExcel.Range(cellAddress).Column
			row = excelSheet.Range[excelSheet.Cells(addressX, originY), excelSheet.Cells(addressX, boundY)].Value2
			dataOut.append(row)
	else:
		cellAddress = excelSheet.Cells.Find(searchValues, xlAfter, xlLookIn, xlLookAt, xlSearchOrder, xlSearchDirection, xlMatchCase, xlMatchByte, xlSearchFormat).Address(False, False)
		addressX = objExcel.Range(cellAddress).Row
		addressY = objExcel.Range(cellAddress).Column
		row = excelSheet.Range[excelSheet.Cells(addressX, originY), excelSheet.Cells(addressX, boundY)].Value2
		dataOut = row
	
	objExcel.ActiveWorkbook.Close(False)
	objExcel.screenUpdating = True

else:
	dataOut = "Set RunMe to True."

#Assign your output to the OUT variable
OUT = dataOut
