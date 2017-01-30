<%
' Comersus Diagnostics
' e-commerce ASP Open Source
' CopyRight Rodrigo S. Alhadeff, Comersus
' September 2004
' http://www.comersus.com 
%>
<!--#include file="comersus_backoffice_adminVerify.asp"--> 
<!--#include file="header.asp"--> 

<b>Session variables</b><br>
<br>Storing: <%session("message")="Test ok!"%> done 
<br>Retrieve: <%=session("message")%> 
<br>
<!--#include file="footer.asp"--> 