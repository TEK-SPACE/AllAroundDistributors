<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"--> 
 <!--#include file="../store/comersus_optDes.asp"--> 
<!--#include file="../includes/encryption.asp"-->
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->   
<!--#include file="../includes/settings.asp"-->   

<%  
on error resume next

dim conntemp, rstemp, mysql, etime, stime, psw, txt, strTemp, x

' get settings 
pEncryptionPassword 	= getSettingKey("pEncryptionPassword")
pEncryptionMethod 	= getSettingKey("pEncryptionMethod")

etime = 0
stime = 0

psw = pEncryptionPassword
txt = request.form("txt")

if txt="" then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Please enter an input string")
end if
%>

<!--#include file="header.asp"-->
<!--#include file="./includes/settings.asp"--> 

<br><b>Encryption & Decryption</b><br>
<br>Input string: <%=txt%>
<br>Encryption string: <%=psw%>
<%

stime 		= timer
strTemp 	= EnCrypt(txt, psw)
strTemp2 	= DeCrypt(txt, psw)
etime 		= cdbl(timer - stime)   

if strTemp2="" then
strTemp2="N/A"
end if
%>
<br>
<br>Encrypted text: <b><%=strTemp%></b>
<br>Decrypted text: <b><%=strTemp2%></b>
<br>
<br>Encryption took: <%=etime%> (&plusmn;55 msec)

<%call closeDb()%>
<!--#include file="footer.asp"--> 