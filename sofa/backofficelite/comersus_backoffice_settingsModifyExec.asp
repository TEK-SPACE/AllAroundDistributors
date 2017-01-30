<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"--> 
<!--#include file="../includes/settings.asp"-->    
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="./includes/settings.asp"-->  

<% 
on error resume next

dim mySQL, conntemp, rstemp

if pBackOfficeDemoMode=-1 then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Function disabled in demo mode")
end if

for each field in request.form 

 pSettingValue=request.form(field)
 pSettingValue=replace(pSettingValue,"'","''")  
 
 if isNumeric(field) then
  mySQL="UPDATE settings SET settingValue='" &pSettingValue& "' WHERE idSetting="& field              
  call updateDatabase(mySQL, rstemp, "comersus_backoffice_settingsModifyExec.asp") 
 end if  
	
next 

call closeDb()

response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Settings updated")

%>