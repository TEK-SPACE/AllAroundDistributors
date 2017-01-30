<%
' Comersus Diagnostics
' e-commerce ASP Open Source
' CopyRight Rodrigo S. Alhadeff, Comersus
' September 2004
' http://www.comersus.com 
' Diagnostics for Comersus store
%>
<!--#include file="comersus_backoffice_adminVerify.asp"--> 
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"-->  
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/encryption.asp"-->  
<!--#include file="header.asp"--> 
<b>Basic security check</b><br>
<%
on error resume next

dim mySQL, conntemp, rstemp

' settings
pEncryptionPassword 	= getSettingKey("pEncryptionPassword")
pEncryptionMethod 	= getSettingKey("pEncryptionMethod")
pSendPlainText		= getSettingKey("pSendPlainText")

alertLevel		= 0
dim securityIssue(6)

if instr(pDatabaseConnectionString,"\database\comersus.mdb")>0 then
	securityIssue(1)="Change default Access database location and edit includes/settings.asp to specify new location"
	alertLevel = alertLevel + 2
elseif instr(pDatabaseConnectionString,"SQL Server")>0  then
	position=instr(pDatabaseConnectionString,"password=")
	if  (instr(pDatabaseConnectionString,"password=;")>0 or instr(pDatabaseConnectionString,"password")=0 or mid(pDatabaseConnectionString,position+9)="")then
		securityIssue(1)="Enter a complex password for SQL Server and update includes/settings.asp"
		alertLevel = alertLevel + 2
	end if	
elseif instr(pDatabaseConnectionString,"mySQL")>0  then
	position=instr(pDatabaseConnectionString,"Pwd=")
	if (mid(pDatabaseConnectionString,position+4)="" or instr(pDatabaseConnectionString,"Pwd")=0)then
		securityIssue(1)="Enter a complex password for mySQL and update includes/settings.asp"
		alertLevel = alertLevel + 2	
	end if		
end if

mySQL="SELECT adminpassword FROM admins WHERE adminName='Admin'"
call getFromDatabase(mySQL, rsTemp, "securityCheck")

if Decrypt(rsTemp("adminpassword"),pEncryptionPassword)="123456" then
	securityIssue(3)="Change default admin password"
	alertLevel = alertLevel + 2
end if

if pEncryptionPassword="HGSDYGDSLWREIUCJD938439402342" then
	securityIssue(4)="Change default encryption key and update all old stored passwords and encrypted info in your database"
	alertLevel = alertLevel + 2
end if

mySQL="SELECT count(idPayment) AS UseOffLinePayments FROM payments WHERE redirectionUrl like '%comersus_offLinePaymentForm.asp%'"
call getFromDatabase(mySQL, rsTemp, "securityCheck")
if rsTemp("UseOffLinePayments")>0 and pSendPlainText="-1" then
	securityIssue(5)="Change the setting pSendPlainText so credit card information is not send by plain email"
	alertLevel = alertLevel + 2
end if
%>
<br>Security level: <%=alertLevel%>. 
<%if alertLevel=0 then%>
<br>Basic test results are ok. Read Comersus documentation and online suggestions to increase your store security.
<%else%>
<br><br>Suggestions:<br>
<%for i=1 to 6
if securityIssue(i)<>"" then
	response.write "-" & securityIssue(i) & "<br>"
end if
next
end if%>
<br>
<!--#include file="footer.asp"--> 