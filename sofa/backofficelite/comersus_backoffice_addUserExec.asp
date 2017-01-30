<%
' Comersus BackOffice 
' e-commerce ASP Open Source
' CopyRight Rodrigo S. Alhadeff, Comersus
' May-2002
' http://www.comersus.com 
pPageLevel=Cint(0)
%>

<!--#include file="comersus_backoffice_adminVerify.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/settings.asp"-->  
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->  
<!--#include file="../includes/encryption.asp"--> 
<!--#include file="../store/comersus_optDes.asp"--> 

<% 
on error resume next

dim mySQL, conntemp, rstemp

' get settings 
pEncryptionPassword 	= getSettingKey("pEncryptionPassword")
pEncryptionMethod 	= getSettingKey("pEncryptionMethod")

pAdminName 		= request.querystring("adminName")
pAdminPassword	 	= EnCrypt(request.querystring("adminPassword"), pEncryptionPassword)
pAdminLevel	 	= request.querystring("adminLevel")

' check if there is another one with the same adminName

mySQL="SELECT idAdmin FROM admins WHERE lcase(adminName)=lcase('" &pAdminName&"')"

call getFromDatabase(mySQL, rstemp, "comersus_backoffice_adduserexec.asp") 

if not rstemp.eof then
 response.redirect "comersus_backoffice_message.asp?message="&Server.Urlencode("There is another user with that name.")
end if

' insert into db
mySQL="INSERT INTO admins (adminName, adminPassword, adminLevel) VALUES ('" &pAdminName& "','" &pAdminPassword& "'," &pAdminLevel&")"

call updateDatabase(mySQL, rstemp, "comersus_backoffice_adduserexec.asp") 

call closeDb()

response.redirect "comersus_backoffice_message.asp?message="&Server.Urlencode("Admin user added.")
%>