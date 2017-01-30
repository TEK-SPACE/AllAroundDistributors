<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/settings.asp"-->    
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/encryption.asp"--> 
<!--#include file="../store/comersus_optDes.asp"--> 

<%
on error resume next

dim mySQL, conntemp, rstemp, pemail, ppassword

if request("adminName")="" or request("adminPassword")="" then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Please enter user and password")
end if

' get settings 
pEncryptionPassword 	= getSettingKey("pEncryptionPassword")
pEncryptionMethod 	= getSettingKey("pEncryptionMethod")

' form parameters
pAdminName	= getLoginField(request("adminName"),12)
pAdminPassword	= Cstr(EnCrypt(getLoginField(request("adminPassword"),15), pEncryptionPassword))

' authenticated and charge session
mySQL="SELECT adminName, adminLevel, idAdmin FROM admins WHERE adminName='" &pAdminName& "' AND adminPassword='" &pAdminPassword& "'"
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_login.asp") 

if rstemp.eof then
  session("loginFailed")=-1
  response.redirect "comersus_backoffice_invalidLogin.asp"
end if

session("admin") = 1

call closeDb()

response.redirect "comersus_backoffice_menu.asp?lastLogin="& pLastLogin
%>