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
<b>Define permissions for a hosting service</b>
<br>
<br>1. Login into your control panel
<br>2. Try to locate a link to manage file and folder permissions
<br>3. Assign all permissions to comersus/database folder and comersus.mdb file
<br><br>Note: some hosting services provide a /db folder with permissions pre-assigned. In that case, copy comersus.mdb there, create a DSN connection named "comersus" (using Control Panel) and edit comersus/includes/settings.asp with a text editor to comment current database connection string and uncomment DSN connection string (pDatabaseConnectionString = "DSN=comersus")
<br><br>

<form method=post action=default.asp>
<input type=submit name=s value=">> Solved permissions issue? Try again">
</form>

<!--#include file="footer.asp"-->