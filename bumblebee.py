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
