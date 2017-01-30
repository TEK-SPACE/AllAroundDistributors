<%
' Comersus Diagnostics
' e-commerce ASP Open Source
' CopyRight Rodrigo S. Alhadeff, Comersus
' September 2002
' http://www.comersus.com 
' Diagnostics for Comersus store
%>
<!--#include file="comersus_backoffice_adminVerify.asp"--> 
<!--#include file="header.asp"--> 
<b>ASP Execute test</b><br><br>
>>Starting execution test....<br><br>
Trying Active Server Pages Execution:<%="ok!"%> <br>
Using Server: <%=request.servervariables("SERVER_SOFTWARE")%> <br>
Path:<%=request.servervariables("PATH_INFO")%> <br>
Cookie: <%=request.servervariables("HTTP_COOKIE")%> <br><br>
<%=">>End of test"%>
<br>
<!--#include file="footer.asp"--> 