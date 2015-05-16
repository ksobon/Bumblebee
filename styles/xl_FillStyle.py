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

import bumblebee as bb

#The inputs to this node will be stored as a list in the IN variable.
dataEnteringNode = IN

patternType = IN[0]
backgroundColor = IN[1]
patternColor = IN[2]
opacity = IN[3]
bevelType = IN[4]

fillStyle = bb.BBFillStyle()

if patternType != None:
	fillStyle.patternType = patternType
if backgroundColor != None:
	fillStyle.backgroundColor = backgroundColor
if patternColor != None:
	fillStyle.patternColor = patternColor
if opacity != None:
	fillStyle.opacity = opacity
if bevelType != None:
	fillStyle.bevelType = bevelType

#Assign your output to the OUT variable
OUT = fillStyle
