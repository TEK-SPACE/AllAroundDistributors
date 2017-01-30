<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' USA - 2005
' http://www.comersus.com 
' Details: misc CC validate functions
%>

<%

Function ValidateCreditCard(ByRef pStrNumber, ByRef pStrType)

' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT
' WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES
' OF MERCHANTABILITY AND/OR FITNESS FOR A  PARTICULAR
' PURPOSE.

' Copyright 2000 - 2001. All rights reserved.
' Lewis Moten
' http://www.lewismoten.com
' email: lewis@moten.com

	
	ValidateCreditCard = False ' Initialize negative results
	
	' Clean Credit Card Number (Removes dashes and spaces)
	pStrNumber = ParseDigits(pStrNumber)		
	
	' Validate number with LUHN Formula
	If Not LUHN(pStrNumber) Then Exit Function
	
	' Apply rules based on type of card
	Select Case pStrType
		Case "A"
			Select Case Left(pStrNumber, 2)
				Case "34", "37"
					' Do Nothing
				Case Else
					Exit Function
			End Select
			If Not Len(pStrNumber) = 15 Then Exit Function
		Case "D"
			Select Case Left(pStrNumber, 2)
				Case "36" , "38"
					' Do Nothing
				Case "30"
					Select Case Left(pStrNumber, 3)
						Case "300", "301", "302", "303", "304", "305"
							' Do Nothing
						Case Else
							Exit Function
					End Select
				Case Else
					Exit Function
			End Select
			If Not Len(pStrNumber) = 14 Then Exit Function
		Case "Discover"
			If Not Left(pStrNumber, 4) = "6011" Then Exit Function
			If Not Len(pStrNumber) = 16 Then Exit Function
		Case "JCB"
			If Left(pStrNumber, 1) = "3" And Len(pStrNumber) = 16 Then
				' Do Nothing
			ElseIf Left(pStrNumber, 14) = "2131" And Len(pStrNumber) = 15 Then
				' Do Nothing
			ElseIf Left(pStrNumber, 14) = "1800" And Len(pStrNumber) = 15 Then
				' Do Nothing
			Else
				Exit Function
			End If
		Case "M"
			Select Case Left(pStrNumber, 2)
				Case "51", "52", "53", "54", "55"
					' Do Nothing
				Case Else
					Exit Function
			End Select
			If Not Len(pStrNumber) = 16 Then Exit Function
		Case "V"
			If Not Left(pStrNumber, 1) = "4" Then Exit Function
			If Not (Len(pStrNumber) = 13 Or Len(pStrNumber) = 16) Then Exit Function
		Case Else
			' Unknown Card Type
			Exit Function
	End Select
	
	' We got this far so the number passed all the rules!
	ValidateCreditCard = True
	
End Function
' ------------------------------------------------------------------------------
Function LUHN(ByRef pStrDigits)
	
	Dim lLngMaxPosition
	Dim lLngPosition
	Dim lLngSum				' Sum of all positions
	Dim lLngDigit			' Current digit in specified position
	
	' Initialize
	lLngMaxPosition = Len(pStrDigits)
	lLngSum = 0
	
	' Read from right to left
	For lLngPosition = lLngMaxPosition To 1 Step -1
		
		' If we are working with an even digit (from the right)
		If (lLngMaxPosition - lLngPosition) Mod 2 = 0 Then
		
			lLngSum = lLngSum + CInt(Mid(pStrDigits, lLngPosition, 1))
		
		Else
		
			' Double the digit
			lLngDigit = CInt(Mid(pStrDigits, lLngPosition, 1)) * 2
			
			' shortcut adding sum of digits
			If lLngDigit > 9 Then lLngDigit = lLngDigit - 9
			
			lLngSum = lLngSum + lLngDigit
			
		End If
	Next

	' A mod 10 check must not return any remainders
	LUHN = lLngSum Mod 10 = 0

End Function
' ------------------------------------------------------------------------------
Function ParseDigits(ByRef pStrData)
	
	' Strip all the numbers from a string
	' (cleans up dashes and spaces)
	
	Dim lLngMaxPosition
	Dim lLngPosition
	
	lLngMaxPosition = Len(pStrData)
	
	For lLngPosition = 1 To lLngMaxPosition
		If IsNumeric(Mid(pStrData, lLngPosition, 1)) Then
			ParseDigits = ParseDigits & Mid(pStrData, lLngPosition, 1)
		End If
	Next

End Function
%>