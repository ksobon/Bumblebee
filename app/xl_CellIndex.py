#Copyright(c) 2015, David Mans, Konrad Sobon
# @arch_laboratory, http://archi-lab.net, http://neoarchaic.net

import clr
import sys

pyt_path = r'C:\Program Files (x86)\IronPython 2.7\Lib'
sys.path.append(pyt_path)

import os
appDataPath = os.getenv('APPDATA')
bbPath = appDataPath + r"\Dynamo\0.7\packages\Bumblebee\extra"
if bbPath not in sys.path:
	sys.path.Add(bbPath)

import bumblebee as bb
import string
import re

#The inputs to this node will be stored as a list in the IN variable.
dataEnteringNode = IN

cellAddress = str(IN[0])

match = re.match(r"([a-z]+)([0-9]+)", cellAddress, re.I)
if match:
    addressItems = match.groups()

row = bb.ConvertChar(addressItems[0])
column = int(addressItems[1])
	
#Assign your output to the OUT variable
OUT = row, column
