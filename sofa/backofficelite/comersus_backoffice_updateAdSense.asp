<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"-->   
<!--#include file="../includes/stringFunctions.asp"-->   
<!--#include file="../includes/screenMessages.asp"-->   
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->  
<!--#include file="../includes/settings.asp"-->    

<% 
on error resume next

dim mySQL, conntemp, rstemp, rstemp2

' get settings 
pRunInstallationWizard 	= getSettingKey("pRunInstallationWizard")

if pRunInstallationWizard<>"-1" then 
 response.redirect "comersus_backoffice_message.asp?message="&Server.UrlEncode("Sorry, the Installation Wizard is disabled")
end if

pAdSenseClient=getUserInput(request.form("pAdSenseClient"),30)

if pAdSenseClient<>"" then
 pAdSenseClient=replace(pAdSenseClient,"ca-","")
 mySQL = "UPDATE settings SET settingValue='" & pAdSenseClient & "' WHERE settingKey='pAdSenseClient'"
 call updateDatabase(mySQL, rsTemp, "comersus_backoffice_updateAdSense.asp")
end if

pAdSenseType=getUserInput(request.querystring("pAdSenseType"),30)

if pAdSenseType<>"" then
 mySQL = "UPDATE settings SET settingValue='" & pAdSenseType & "' WHERE settingKey='pAdSenseType'"
 call updateDatabase(mySQL, rsTemp, "comersus_backoffice_updateAdSense.asp")
end if

call closeDb()
 
response.redirect "comersus_backoffice_message.asp?message="&Server.UrlEncode("AdSense Configuration Updated. You can close this window.")

%>

