#Copyright(c) 2015, David Mans, Konrad Sobon
# @arch_laboratory, http://archi-lab.net, http://neoarchaic.net

#The inputs to this node will be stored as a list in the IN variable.
dataEnteringNode = IN

name = IN[0]
size = IN[1]
color = IN[2]
hAlign = IN[3]
vAlign = IN[4]
bold = IN[5]
italic = IN[6]
underline = IN[7]
strikethrough = IN[8]

if name == None:
	name = "Arial"
if size == None:
	size = str(8)
else:
	size = str(IN[1])
if color == None:
	color = ",".join([str(0), str(0), str(0)])
else:
	color = ",".join([str(color.Red), str(color.Green), str(color.Blue)])
if hAlign == None:
	hAlign = "Left"
if vAlign == None:
	vAlign = "Center"
if bold == None:
	bold = "false"
else:
	bold = str(IN[5])
if italic == None:
	italic = "false"
else:
	italic = str(IN[6])
if underline == None:
	underline = "false"
else:
	underline = str(IN[7])
if strikethrough == None:
	strikethrough = "false"
else:
	strikethrough = str(IN[8])

textStyle = "~".join([name, size, color, hAlign, vAlign, bold, italic, underline, strikethrough])

#Assign your output to the OUT variable
OUT = textStyle
