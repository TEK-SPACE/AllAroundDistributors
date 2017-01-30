<%
' ###################################################
' ## BarChart                                       #
' ## Class to draw bar charts in ASP applications   #
' ## Author : Anton Bawab                           #
' ## First written : March 27th 2000                #
' ###################################################

' ###################################################
' ## Include this file in your ASP script           #
' ## assign the properties                          #
' ## then use the Draw method                       #
' ###################################################

Class BarChart

Private mchartBGcolor
Private mchartTitle
Private mchartWidth
Private mchartValueArray
Private mchartLabelsArray
Private mchartColorArray
Private mchartViewDataType
Private mchartBarHeight
Private mchartBorder
Private mchartTextColor
Private mchartCounter ' general counter
Private mchartMaxValue
Private mchartFactor
Private mchartTotalValues
Private mchartMinValue

Public Property LET chartBGcolor(strColor)
	mchartBGcolor = strColor

	'code validation
		IF LEN(mchartBGcolor) <> 7 THEN
			ERR.Number = vbObjectError + 1000
			ERR.Description = "Color string provided unequal to 7 characters"
			Response.Write Err.Number & vbCRLF & ERR.Description
			ERR.Clear

			EXIT Property
		END IF
END Property

Public Property LET chartTitle(strTitle)
	mchartTitle = strTitle
END Property

Public Property LET chartWidth(intWidth)
	mchartWidth = intWidth
END Property

Public Property LET chartValueArray(arrValues)
	
	mchartValueArray = arrValues

	IF NOT isArray(mchartValueArray) THEN
		ERR.Number = vbObjectError + 1001
		ERR.Description = "Values passed are not an array"
		Response.Write Err.Number & vbCRLF & ERR.Description
		EXIT Property
		ERR.Clear
		ERR.Number = vbObjectError + 1002
		ERR.Description "Number of values passed does not match labels"
		Response.Write Err.Number & vbCRLF & ERR.Description
		ERR.Clear
		EXIT Property
	END IF

END Property

Public Property LET chartLabelsArray(arrLabels)
	
	mchartLabelsArray = arrLabels

	IF NOT isArray(mchartLabelsArray) THEN
		ERR.Number = vbObjectError + 1001
		ERR.Description = "Label values passed are not an array"
		Response.Write Err.Number & vbCRLF & ERR.Description
		EXIT Property
		ERR.Clear
	ELSEIF UBOUND(mchartValueArray) <> UBOUND(mchartLabelsArray) THEN
		ERR.Number = vbObjectError + 1002
		ERR.Description = "Number of values passed does not match labels"
		Response.Write Err.Number & vbCRLF & ERR.Description
		ERR.Clear	
		EXIT Property
	END IF
END Property

Public Property LET chartColorArray(arrColors)
	Dim tempNumOfColors, I
	
	mchartColorArray = arrColors

	IF NOT isArray(mchartColorArray) THEN
		ERR.Number =  vbObjectError + 1001
		ERR.Description = "Color values passed are not an array"
		Response.Write Err.Number & vbCRLF & ERR.Description
		EXIT Property
		ERR.Clear
	END IF

	' match the number of the colors to the number of elements to draw
	IF UBOUND(mchartColorArray) < UBOUND(mchartValueArray) THEN
		tempNumOfColors = UBOUND(mchartColorArray) 'Get the number of colors provided

		REDIM PRESERVE mchartColorArray(UBOUND(mchartValueArray))

		' Cycling the values through the array
		For I = tempNumOfColors+1 to UBOUND(mchartColorArray)
			mchartColorArray(I) = mchartColorArray((I mod (tempNumOfColors+1)))
		NEXT

	END IF
END Property

Public Property LET chartViewDataType(strProp)
	mchartViewDataType = UCASE(strProp)

	IF (mchartViewDataType <> "N") AND (mchartViewDataType <> "P") AND (mchartViewDataType <> "V") THEN
		mchartViewDataType = "V"
	END IF

END Property

Public Property LET chartBarHeight(intBarHeight)
	mchartBarHeight = intBarHeight

	IF NOT ISNumeric(mchartBarHeight) THEN
		ERR.Number =  vbObjectError + 1003
		ERR.Description "chartBarHeight property can only accept numerical values"
		Response.Write Err.Number & vbCRLF & ERR.Description
		EXIT Property
		ERR.Clear
	END IF
END Property

Public Property LET chartBorder(intBorder)
	mchartBorder = intBorder

	IF NOT ISNumeric(mchartBorder) THEN
		ERR.Number =  vbObjectError + 1003
		ERR.Description "chartBorder property can only accept numerical values"
		Response.Write Err.Number & vbCRLF & ERR.Description
		EXIT Property
		ERR.Clear
	END IF
END Property

Public Property LET chartTextColor(strColor)
	mchartTextColor = strColor

	IF LEN(mchartTextColor) <> 7 THEN
		ERR.Number =  vbObjectError + 1000
		ERR.Description = "Color string provided less than 7 characters"
		Response.Write Err.Number & vbCRLF & ERR.Description
		ERR.Clear
		EXIT Property
	END IF
END Property


Private Property LET chartMaxValue(intValue)
	mchartMaxValue = intValue
END Property

Private Property LET chartMinValue(intValue)
	mchartMinValue = intValue
END Property

Private Property LET chartTotalValues(intValue)
	mchartTotalValues = intValue
END Property

Public Property GET chartMaxValue
	chartMaxValue = mchartMaxValue
END Property

Public Property GET chartMinValue
	chartMinValue = mchartMinValue
END Property

Public Property GET chartTotalValues
	chartTotalValues = mchartTotalValues
END Property


Private Function MakeChart()
Dim F

' getting the hieghest and lowest values within the array
' and calculating the total of the values
mchartMinValue = 0
mchartMaxValue = 0
mchartTotalValues = 0
For each F in mchartValueArray
		IF F > mchartMaxValue THEN
			mchartMaxValue = F
		END IF

		IF mchartMinValue = 0 THEN
			mchartMinValue = F
		ELSEIF F < mchartMinValue THEN
			mchartMinValue = F
			' Response.Write mchartMinValue
		END IF

	mchartTotalValues = mchartTotalValues + F
	' getting the total of the values in the array
NEXT

chartMaxValue = mchartMaxValue
chartMinValue = mchartMinValue
chartTotalValues = mchartTotalValues

' Determining the factor to use for resizing the values to fit
' within the given width
IF mchartMaxValue > (mchartWidth-20) THEN
	' getting the factor
	mchartFactor = mchartMaxValue / (mchartWidth-20)
	'Response.Write("Factor of : " & mchartFactor & "<BR>")

	' changing the values of all the entries within the array
	For mchartCounter = 0 to UBOUND(mchartValueArray)
		mchartValueArray(mchartCounter) = CINT(mchartValueArray(mchartCounter) / mchartFactor)
	NEXT
END IF

' Modifying the chartLabelsArray to reflect the setting required
SELECT CASE mchartViewDataType
	Case "V" ' display the value
		For mchartCounter = 0 to UBOUND(mchartValueArray)
			mchartLabelsArray(mchartCounter) = mchartLabelsArray(mchartCounter) & "-" & mchartValueArray(mchartCounter)
		NEXT

	Case "P" ' display the percentage
		For mchartCounter = 0 to UBOUND(mchartValueArray)
			mchartLabelsArray(mchartCounter) = mchartLabelsArray(mchartCounter) & "-" & ((mchartValueArray(mchartCounter) / mchartTotalValues) * 100) & "%"
		NEXT
END SELECT

MakeChart = "<table width=""" & mchartWidth & """ border=""" & mchartBorder & """>"
MakeChart = MakeChart & "<tr><td bgcolor=""" & mchartBGcolor & """>"

MakeChart = MakeChart & "<table width=""100%"" border=""0"" cellpadding=""1"" cellspacing=""1""><tr>"
MakeChart = MakeChart & "<th colspan=""2""><b><font face=""Arial, Tahoma, Verdana"" color=""" & mchartTextColor & """ size=""1"">"
MakeChart = MakeChart & "<u><b>" & mchartTitle & "</b></u></font></th></tr>"

FOR mchartCounter = 0 to UBOUND(mchartValueArray)
	MakeChart = MakeChart & "<tr><td valign=""middle"" align=""left"">"
	MakeChart = MakeChart & "<font face=""Arial, Tahoma, Verdana"" color=""" & mchartTextColor & """ size=""1"">"
	MakeChart = MakeChart & mchartLabelsArray(mchartCounter) & "</font></td>"
	MakeChart = MakeChart & "<td valign=""middle"" align=""left"">"
	MakeChart = MakeChart & "<table border=""0"" cellpadding=""1"" cellspacing=""0"">"
	MakeChart = MakeChart & "<tr height=""" & mchartBarHeight & """>"
	MakeChart = MakeChart & "<td width=""" & mchartValueArray(mchartCounter) & """ bgcolor=""" & mchartColorArray(mchartCounter) & """>"	
	MakeChart = MakeChart & "</td></tr></table>"
	MakeChart = MakeChart & "</td></tr>"
NEXT

MakeChart = MakeChart & "</table>"
MakeChart = MakeChart & "</tr></td></table>"
MakeChart = MakeChart & vbCRLF & "<!--Chart created with BarChartClass by Anton Bawab © 2000-->"
END Function

Public SUB Draw()
	CheckProps()
	Response.Write MakeChart()
END SUB

Private Function CheckProps()

		IF ISEMPTY(mchartBGcolor) THEN chartBGcolor = "#FFFFFF"

		IF ISEMPTY(mchartColorArray) THEN chartColorArray = Array("#990000" , "#009900" , "#000099")

		IF ISEMPTY(mchartTitle) THEN chartTitle = "Chart Title"

		IF ISEMPTY(mchartViewDataType) THEN chartViewDataType = "V"

		IF ISEMPTY(mchartBarHeight) Then mchartBarHeight = 15

		IF ISEMPTY(mchartBorder) THEN mchartBorder = 0

		IF ISEMPTY(mchartTextColor) THEN	mchartTextColor = "#000000"

END FUNCTION
END CLASS
%>