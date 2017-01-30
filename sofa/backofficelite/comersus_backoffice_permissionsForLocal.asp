<%
' Comersus BackOffice Lite
' Free Management Utility
' Comersus Open Technologies LC
' April-2003
' http://www.comersus.com 
%>

<!--#include file="includes/settings.asp"-->
<!--#include file="../includes/settings.asp"-->    
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<% 
on error resume next 

dim mySQL, conntemp, rstemp

%>
<!--#include file="header.asp"--> 
<b>Define permissions for a local installation</b>
<br>
<br>Windows XP Pro<br>
1. Disable the "Use simple file sharing" setting<br>
Within the Windows Explorer [Start -> All Programs -> Accessories -> Windows Explorer],
open the Folder Options interface [Tools -> Folder Options] and then go the the [View] tab
and deselect the "Use simple file sharing" option. Select [OK] to save the change. Now you
will see the [Security] tab within the folder Properties interface... and the [Sharing] tab of the
root file system of a drive shows the available setting options, without warning about how
risky it is to share a root directory.<br>
2. Press right click over comersus/database folder, go to Security TAB, give all permissions
to Anonymus Internet User<br>
3. Go to sharing TAB, Share the database folder

<br><br>Windows 2000
<br>1. Right click on the database folder (usually c:\Inetpub\wwwRoot\comersus\database)
<br>2. Click Properties
<br>3. Click Security tab
<br>4. Add the IUSR_MACHINENAME user to the permissions list, and grant this
user full control. (where MACHINENAME is the hostname of the server)
<br><br>

<form method=post action=default.asp>
<input type=submit name=s value=">> Solved permissions issue? Try again">
</form>
<!--#include file="footer.asp"-->