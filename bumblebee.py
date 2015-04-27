"""
Copyright(c) 2015, Konrad Sobon
@arch_laboratory, http://archi-lab.net

Copyright (c) 2015, David Mans
http://neoarchaic.net

Excel and Dynamo interop library

"""
import clr
import sys

pyt_path = r'C:\Program Files (x86)\IronPython 2.7\Lib'
sys.path.append(pyt_path)

import System
from System import Array
from System.Collections.Generic import *

clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel

System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo("en-US")
from System.Runtime.InteropServices import Marshal

import string
import re

""" Misc Functions """
def ProcessList(_func, _list):
    return map( lambda x: ProcessList(_func, x) if type(x)==list else _func(x), _list )

""" Misc Definitions"""

def ConvertNumber(num):
    letters = ''
    while num:
	    mod = num % 26
	    num = num // 26
	    letters += chr(mod + 64)
    return ''.join(reversed(letters))

def ConvertChar(char):
    number =- 25
    for l in char:
	    if not l in string.ascii_letters:
		    return False
	    number += ord(l.upper()) - 64 + 25
    return int(number)

def CellIndex(cellAddress):
    match = re.match(r"([a-z]+)([0-9]+)", cellAddress, re.I)
    if match:
        addressItems = match.groups()
        row = ConvertChar(addressItems[0])
        column = int(addressItems[1])
    return [row, column]

def xlRange(address):
    originAddress = address.split(":")[0]
    extentAddress = address.split(":")[1]
    originRow = int(CellIndex(originAddress)[0])
    originCol = int(CellIndex(originAddress)[1])
    extentRow = int(CellIndex(extentAddress)[0])
    extentCol = int(CellIndex(extentAddress)[1])
    return [originRow, originCol, extentRow, extentCol]

def GetLineStyle(key):
    keys = ["Continuous", "Dash", "DashDot", "DashDotDot", "RoundDot", "SquareDotMSO", "LongDash", 
            "DoubleXL", "NoneXL"]
    values = [1, -4115, 4, 5, -4118, -4118, -4115, -4119, -4142]
    d = dict()
    for i in range(len(keys)):
        d[keys[i]] = values[i]
    if key in d:
	return d[key]
    else:
	return None

def GetPatternType(key):
    keys = ["xlCheckerBoard", "xlCrissCross", "xlDarkDiagonalDown", "xlGrey16", "xlGray25", 
	    "xlGray50", "xlGray75", "xlGray8", "xlGrid", "xlDarkHorizontal", 
	    "xlLightDiagonalDown", "xlLightHorizontal", "xlLightDiagonalUp", "xlLightVertical", "xlNone", 
	    "xlSemiGray75", "xlSolid", "xlDarkDiagonalUp", "xlDarkVertical"]
    values = [9, 16, -4121, 17, -4124, 
		-4124, -4126, 18, 15, -4128,
		13, 11, 14, 12, -4142, 
		10, 1, -4162, -4166]
    d = dict()
    for i in range(len(keys)):
	    d[keys[i]] = values[i]
    if key in d:
	    return d[key]
    else:
	    return None

def RGBToRGBLong(rgb):
    strValue = '%02x%02x%02x' % rgb
    iValue = int(strValue, 16)
    return iValue
