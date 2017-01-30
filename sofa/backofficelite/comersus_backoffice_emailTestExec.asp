<%
' Comersus Diagnostics
' e-commerce ASP Open Source
' CopyRight Rodrigo S. Alhadeff, Comersus
' September 2002
' http://www.comersus.com 
' Diagnostics for Comersus store
%>
<!--#include file="comersus_backoffice_adminVerify.asp"-->    
<!--#include file="../includes/sendmail.asp"--> 
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"-->  
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->   

<% 
on error resume next 

dim connTemp, rsTemp

' get settings 
pDefaultLanguage 	= getSettingKey("pDefaultLanguage")
pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanyLogo	 	= getSettingKey("pCompanyLogo")

pEmailSender		= getSettingKey("pEmailSender")
pEmailAdmin		= getSettingKey("pEmailAdmin")
pSmtpServer		= getSettingKey("pSmtpServer")
pEmailComponent		= getSettingKey("pEmailComponent")
pDebugEmail		= getSettingKey("pDebugEmail")

%>
<!--#include file="header.asp"--> 
<b>Email Test</b><br><br>
Trying to send email...

<%
on error resume next

pSmtpServer	= request.form("pSmtpServer")
pEmailFrom	= request.form("pEmailFrom")
pEmailTo	= request.form("pEmailTo")

pEmailComponent="PersitsASPMail" 

call sendMail(pEmailFrom, pEmailFrom, pEmailTo, "Comersus e-mail diagnostics", "This configuration works!"&Vbcrlf&"Component:"&pEmailComponent&Vbcrlf&"SMTP Server:"&pSmtpServer&Vbcrlf&"Email From:"&pEmailFrom)

pEmailComponent="Jmail"

call sendMail(pEmailFrom, pEmailFrom, pEmailTo, "Comersus e-mail diagnostics", "This configuration works!"&Vbcrlf&"Component:"&pEmailComponent&Vbcrlf&"SMTP Server:"&pSmtpServer&Vbcrlf&"Email From:"&pEmailFrom)

pEmailComponent="ServerObjectsASPMail1"

call sendMail(pEmailFrom, pEmailFrom, pEmailTo, "Comersus e-mail diagnostics", "This configuration works!"&Vbcrlf&"Component:"&pEmailComponent&Vbcrlf&"SMTP Server:"&pSmtpServer&Vbcrlf&"Email From:"&pEmailFrom)

pEmailComponent="ServerObjectsASPMail2"

call sendMail(pEmailFrom, pEmailFrom, pEmailTo, "Comersus e-mail diagnostics", "This configuration works!"&Vbcrlf&"Component:"&pEmailComponent&Vbcrlf&"SMTP Server:"&pSmtpServer&Vbcrlf&"Email From:"&pEmailFrom)

pEmailComponent="CDONTS"

call sendMail(pEmailFrom, pEmailFrom, pEmailTo, "Comersus e-mail diagnostics", "This configuration works!"&Vbcrlf&"Component:"&pEmailComponent&Vbcrlf&"SMTP Server:"&pSmtpServer&Vbcrlf&"Email From:"&pEmailFrom)

pEmailComponent="CDO"

call sendMail(pEmailFrom, pEmailFrom, pEmailTo, "Comersus e-mail diagnostics", "This configuration works!"&Vbcrlf&"Component:"&pEmailComponent&Vbcrlf&"SMTP Server:"&pSmtpServer&Vbcrlf&"Email From:"&pEmailFrom)

pEmailComponent="OCXMAIL"

call sendMail(pEmailFrom, pEmailFrom, pEmailTo, "Comersus e-mail diagnostics", "This configuration works!"&Vbcrlf&"Component:"&pEmailComponent&Vbcrlf&"SMTP Server:"&pSmtpServer&Vbcrlf&"Email From:"&pEmailFrom)

pEmailComponent="DUNDASMAIL"

call sendMail(pEmailFrom, pEmailFrom, pEmailTo, "Comersus e-mail diagnostics", "This configuration works!"&Vbcrlf&"Component:"&pEmailComponent&Vbcrlf&"SMTP Server:"&pSmtpServer&Vbcrlf&"Email From:"&pEmailFrom)

pEmailComponent="ASPSMARTMAIL"

call sendMail(pEmailFrom, pEmailFrom, pEmailTo, "Comersus e-mail diagnostics", "This configuration works!"&Vbcrlf&"Component:"&pEmailComponent&Vbcrlf&"SMTP Server:"&pSmtpServer&Vbcrlf&"Email From:"&pEmailFrom)

pEmailComponent="BambooSMTP"

call sendMail(pEmailFrom, pEmailFrom, pEmailTo, "Comersus e-mail diagnostics", "This configuration works!"&Vbcrlf&"Component:"&pEmailComponent&Vbcrlf&"SMTP Server:"&pSmtpServer&Vbcrlf&"Email From:"&pEmailFrom)

%> <br><br>
Test finished.
Check your email account. If you have not received any e-mail in the following minutes try again with different SMTP and email sender
<%call closeDb()%>
<!--#include file="footer.asp"--> 