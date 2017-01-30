<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"--> 
<!--#include file="header.asp"-->
<!--#include file="includes/settings.asp"--> 

<br><b>Encryption & Decryption</b><br>
<form method="post" action="comersus_backoffice_rc4Exec.asp">
String: <input type="text" name="txt"><br><br>               
<input type="submit" value="Encrypt/Decrypt">
</form>
<!--#include file="footer.asp"-->
