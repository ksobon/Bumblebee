#Copyright(c) 2015, David Mans, Konrad Sobon
# @arch_laboratory, http://archi-lab.net, http://neoarchaic.net

#The inputs to this node will be stored as a list in the IN variable.
dataEnteringNode = IN

lineType = IN[0]
weight = IN[1]
color = IN[2]

if lineType == None:
	lineType = "NoneXL"
if weight == None:
	weight = "Thin"
if color == None:
	color = ",".join([str(0), str(0), str(0)])
else:
	color =  ",".join([str(color.Red), str(color.Green), str(color.Blue)])

lineStyle = "~".join([lineType, weight, color])

#Assign your output to the OUT variable
OUT = lineStyle
