<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"-->  
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->   
<!--#include file="../includes/stringFunctions.asp"-->   
<!--#include file="../includes/encryption.asp"--> 
<!--#include file="../store/comersus_optDes.asp"--> 
<!--#include file="includes/settings.asp" --> 

<%
on error resume next

dim conntemp, rstemp, mysql

' get settings 
pEncryptionPassword 	= getSettingKey("pEncryptionPassword")
pEncryptionMethod 	= getSettingKey("pEncryptionMethod")

if pBackOfficeDemoMode=-1 then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Function disabled in demo mode")
end if

pPassword	= getLoginField(request.form("password"),10)
pPasswordVerify	= getLoginField(request.form("passwordVerify"),10)

if pPassword="" then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Please complete password field")
end if

if len(pPassword)<6 then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Please enter a password with more than 5 chars")
end if

if pPassword<>pPasswordVerify then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Incorrect password verification. Please try again")
end if

' encrypt
pPassword	= EnCrypt(pPassword,pEncryptionPassword)

' update user record (all since Lite is not multiadmin)
mysql="UPDATE admins SET adminPassword='" &pPassword&"'"

call updateDatabase(mySQL, rstemp, "comersus_backoffice_changePasswordExec.asp")  	

call closeDb()

response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Password updated")
%>