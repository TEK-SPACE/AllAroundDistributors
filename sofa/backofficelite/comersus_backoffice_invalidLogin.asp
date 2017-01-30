<%
' Comersus BackOffice Lite
' Free Management Utility
' Comersus Open Technologies LC
' Dic-2002
' http://www.comersus.com 
%>

<!--#include file="../includes/stringFunctions.asp"--> 
<!--#include file="header.asp"-->
<!--#include file="includes/settings.asp"--> 
<font size="3"><b>Login results</b></font>

<br><br>Your login information is incorrect

<%if session("loginFailed")=-1 then%>
<br>You can install Diagnostics and Tools to reset the login or you can <a href="http://www.comersus.com/contact.html" target="_blank">contact Comersus</a> to request a remote reset of admin password.
<%end if%>
<br><br>
<!--#include file="footer.asp"--> 