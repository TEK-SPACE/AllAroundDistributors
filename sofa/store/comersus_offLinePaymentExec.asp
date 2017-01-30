<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: takes credit card data for off-line payments, verify exp and number 
%>

<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/screenMessages.asp"--> 
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/encryption.asp" -->
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/sendMail.asp"--> 
<!--#include file="../includes/miscFunctions.asp"--> 
<!--#include file="../includes/preCharge.asp"--> 
<!--#include file="../includes/maxMind.asp"--> 

<% 

on error resume next

dim mySQL, conntemp, rsTemp

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanyLogo	 	= getSettingKey("pCompanyLogo")
pOrderPrefix	 	= getSettingKey("pOrderPrefix")
pEncryptionPassword	= getSettingKey("pEncryptionPassword")
pEncryptionMethod 	= getSettingKey("pEncryptionMethod")
pSendPlainText		= getSettingKey("pSendPlainText")
pEmailSender		= getSettingKey("pEmailSender")
pEmailAdmin		= getSettingKey("pEmailAdmin")
pSmtpServer		= getSettingKey("pSmtpServer")
pEmailComponent		= getSettingKey("pEmailComponent")
pDebugEmail		= getSettingKey("pDebugEmail")

pPreChargeMerchant	= getSettingKey("pPreChargeMerchant")
pMaxMindLicenseKey 	= getSettingKey("pMaxMindLicenseKey")
pMaxMindScoreApproved	= getSettingKey("pMaxMindScoreApproved")
pMaxMindScoreAlert  	= getSettingKey("pMaxMindScoreAlert")

pIdCustomer     	= getSessionVariable("idCustomer",0)
pUserIp	 		= getUserInput(request.ServerVariables("REMOTE_HOST"),20)
pIdOrder		= getUserInput(request.form("idOrder"),20)

' extract real idOrder (without prefix)
pRealIdOrder		= removePrefix(pIdOrder,pOrderPrefix)

pCardType		= getUserInput(request.form("cardType"),8)
pCardNumber		= getUserInput(request.form("cardNumber"),20)
pExpirationMonth	= getUserInput(request.form("expMonth"),2)
pExpirationYear		= getUserInput(request.form("expYear"),4)
pExpiration		= pExpirationMonth & "/" & pExpirationYear 
pSeqCode		= getUserInput(request.form("seqCode"),4)
pNameOnCard		= getUserInput(request.form("nameOnCard"),50)
pStatementAddress	= getUserInput(request.form("statementAddress"),300)

pResponse=""

' send MaxMind request
if pMaxMindLicenseKey <> "0" then
	pScore = 10
	if MaxMindApproved(pIdCustomer, pUserIp, pScore) = 0 then
		response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(717,"This transaction has been rejected. Please verify your information and try again. Score : ") & pScore)
	end if
end if ' maxmind 

' send precharge request
if pPreChargeMerchant<>"0" then
	if preChargeApproved(pCardNumber, pExpirationMonth, pExpirationYear, pResponse)=0 then 
	 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(513,"This transaction has been rejected. Please verify your information and try again. Response details: ") &pResponse)
	end if
end if

pObs="Statement address: " & pStatementAddress & " " & "Name on card: "  & pNameOnCard

' validates missing fields
if trim(pRealIdOrder)="" then
   response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(514,"Missing order #")) 
end if

' validates expiration
if DateDiff("d", Month(Now)&"/"&Year(now), pExpirationMonth&"/"&pExpirationYear)<0 then
   response.redirect "comersus_message.asp?message="&Server.UrlEncode(getMsg(515,"Credit card expired"))
end if

if not ValidateCreditCard(pCardNumber, pCardType) then
	response.redirect "comersus_message.asp?message="&Server.UrlEncode(getMsg(516,"Invalid card number"))
end if

' encrypt CC data
pECardNumber=EnCrypt(pCardNumber, pEncryptionPassword)

' save credit card info
mySQL="INSERT INTO creditCards (idOrder, cardType, cardNumber, expiration, seqCode, obs) VALUES (" &pRealIdOrder& ",'" &pCardType& "','" &pECardNumber& "','" &pExpiration& "','" &pSeqCode& "','" &pObs& "')"

call updateDatabase(mySQL, rstemp, "optOffLinePaymentExec")

if pSendPlainText="-1" then
 ' send plain text email with CC data (try not to use this option for security reasons)
 call sendmail (pCompany, pEmailSender, pEmailAdmin, "CC payment, order: #"&pOrderPrefix&pIdorder, "Card Type:"&VBCrlf&pCardType&VBCrlf&"Name: " &pNameOnCard&Vbcrlf& "Adress:" &pStatementAddress&Vbcrlf& "Card Type: " &pCardType&Vbcrlf& "Card Number: " &pCardNumber&VBcrlf& "Expiration:"&pExpiration)
else
 ' send encrypted email
 call sendmail (pCompany, pEmailSender, pEmailAdmin, "CC payment, order: #"&pOrderPrefix&pIdorder, "Card Type: "&VBCrlf&pCardType&VBCrlf&"Card Number: "&pECardNumber&VBcrlf&VBcrlf&"Expiration:"&pExpiration)
end if

call closeDb()

' go to confirmation page
response.redirect "comersus_offLinePaymentConfirmation.asp?idOrder="&pIdOrder

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
