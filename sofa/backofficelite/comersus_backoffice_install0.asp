<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="../includes/settings.asp"-->    
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 

<% 
on error resume next 

dim mySQL, conntemp, rstemp

pRunInstallationWizard 	= getSettingKey("pRunInstallationWizard")

if pRunInstallationWizard<>"-1" then 
 response.redirect "comersus_backoffice_message.asp?message="&Server.UrlEncode("Sorry, the installation wizard is disabled")
end if

' try to determine if there are permissions defined
pTime=Time()

mySQL1="INSERT INTO settings (settingKey, settingValue) VALUES ('Wizard Test','"&pTime&"') "
call updateDatabase(mySQL1, rstemp, "comersus_backoffice_install.asp") 

' retrieve
mySQL1="SELECT * FROM settings WHERE settingKey='Wizard Test'"
call updateDatabase(mySQL1, rstemp, "comersus_backoffice_install.asp") 

pPermissionsSet=-1

if rstemp.eof then
 pPermissionsSet=0
end if
%>
<!--#include file="header.asp"--> 
<!--#include file="./includes/settings.asp"-->

<br><font size="3"><b>Installation Wizard</b></font><br>

<%if pPermissionsSet=0 then%>
 <br>WARNING (!) It seems that you have not defined database folder permissions.<br>
 <br>If you are using a local computer <a href="comersus_backoffice_permissionsForLocal.asp" target="_blank">click here...</a>
 <br>If you are using a web hosting service <a href="comersus_backoffice_permissionsForHosting.asp" target="_blank">click here...</a>
 <br>If you want to obtain technical assistance <a href="http://www.comersus.org/forum" target="_blank">click here...</a>
 <br>The installation Wizard cannot continue until you solve the permission issue
<%else%>
 <br>Hello, this is your first execution of the BackOffice. 
 <br>With this installation Wizard you will be able to setup your Comersus store in 3 minutes. 
 <br><br>If you need assistance please use "Request Assistance" button in the header.
 <br><br><a href="comersus_backoffice_install1.asp"><img src="images/start.gif" border=0 alt="Click here to start the Wizard"></a>
<%end if%>

<%call closeDb()%>
<!--#include file="footer.asp"-->