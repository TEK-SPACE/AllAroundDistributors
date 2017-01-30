<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/sendMail.asp"-->
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->   
<!--#include file="../includes/databaseFunctions.asp"-->  
<!--#include file="../includes/stringFunctions.asp"-->   

<%
on error resume next

dim mySQL, connTemp, rstemp

' get settings 
pDefaultLanguage 	= getSettingKey("pDefaultLanguage")
pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")

pEncryptionPassword 	= getSettingKey("pEncryptionPassword")
pEmailSender		= getSettingKey("pEmailSender")
pEmailAdmin		= getSettingKey("pEmailAdmin")
pSmtpServer		= getSettingKey("pSmtpServer")
pEmailComponent		= getSettingKey("pEmailComponent")
pDebugEmail		= getSettingKey("pDebugEmail")

pError			= getQueryString(request.querystring("error"),200)
%>

<!--#include file="header.asp"-->
<!--#include file="./includes/settings.asp"-->

<br><b>BackOffice Error</b><br>

<br>Error description: <%=pError%>

<!--#include file="footer.asp"-->