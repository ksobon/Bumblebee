#Copyright(c) 2015, David Mans, Konrad Sobon
# @arch_laboratory, http://archi-lab.net, http://neoarchaic.net

#The inputs to this node will be stored as a list in the IN variable.
dataEnteringNode = IN

patternType = IN[0]
backColor = IN[1]
foreColor = IN[2]
opacity = IN[3]
bevelType = IN[4]

if backColor != None:
	bcolor = ",".join([str(backColor.Red), str(backColor.Green), str(backColor.Blue)])
else:
	bcolor = "xlNone"
if foreColor != None:
	fcolor = ",".join([str(foreColor.Red), str(foreColor.Green), str(foreColor.Blue)])
	if backColor == None:
		bcolor = ",".join([str(255), str(255), str(255)])
	else:
		bcolor = ",".join([str(backColor.Red), str(backColor.Green), str(backColor.Blue)])
else:
	fcolor = "xlNone"
cellFill = "~".join([patternType, bcolor, fcolor, str(opacity), bevelType])

#Assign your output to the OUT variable
OUT = cellFill
