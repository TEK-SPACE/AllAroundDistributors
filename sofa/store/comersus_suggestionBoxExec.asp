<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: send suggestion box
%>
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/screenMessages.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/sendMail.asp"-->

<%
on error resume next

dim mySQL, connTemp, rsTemp

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCompany	 	= getSettingKey("pCompany")

pEmailSender		= getSettingKey("pEmailSender")
pEmailAdmin		= getSettingKey("pEmailAdmin")
pSmtpServer		= getSettingKey("pSmtpServer")
pEmailComponent		= getSettingKey("pEmailComponent")
pDebugEmail		= getSettingKey("pDebugEmail")
pVerificationCodeEnabled= getSettingKey("pVerificationCodeEnabled")

' user form data

pSuggestion		= formatForDb(getUserInput(request.form("suggestion"),1000))
pIdProduct 		= getUserInput(request.form("idProduct"),10)
pSku			= getUserInput(request.form("sku"),20)
pDescription		= getUserInput(request.form("description"),100)
pVerificationCode	= getUserInput(request.form("verificationCode"),8)

pIdCustomer     	= getSessionVariable("idCustomer",0)
pRealBoxCode    	= getSessionVariable("boxCode","")

if pIdCustomer > 0 Then
	mySQL="SELECT name, lastName FROM customers WHERE idCustomer=" &pIdCustomer
	call getFromDatabase(mySQL, rstemp, "suggBoxExec")
	if not rstemp.eof then
	 pName = rstemp("name")& " "&rstemp("lastName")
	end if
end if

if pVerificationCodeEnabled="-1" then
 if pRealBoxCode<>pVerificationCode then
  response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(762,"Invalid"))
 end if
end if

if pSuggestion="" then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(731,"You have to..."))
end if

if pIdCustomer > 0 Then
 call sendmail (pCompany, pEmailSender, pEmailAdmin,"New suggestion for item "&pSku,"Customer: " & pName &vbcrlf& "Date:" & Date() &vbcrlf& "Item: ("&pSku&") " & pDescription &vbCrlf& "Item Id: "&pIdProduct& vbcrlf&"Suggestion: " & vbcrlf&pSuggestion &VBcrlf&VBcrlf)
else
 call sendmail (pCompany, pEmailSender, pEmailAdmin,"New suggestion for item "&pSku,"Date:" & Date() &vbcrlf& "Item: ("&pSku&") " & pDescription &vbCrlf& "Item Id: "&pIdProduct& vbcrlf&"Suggestion: " & vbcrlf&pSuggestion &VBcrlf&VBcrlf)
end if

response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(292,"Thanks, your message was sent."))

call closeDb()
%>
