<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/sendMail.asp"-->
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->   
<!--#include file="../includes/stringFunctions.asp"-->   

<%
on error resume next

dim mySQL, connTemp, rstemp

pError			= getUserInput(request.querystring("error"),300)
%>

<!--#include file="header.asp"-->
<!--#include file="./includes/settings.asp"-->

<br><b>Error</b><br>

<%if instr(pError,"Wizard")<>0 then%>
 <br>WARNING (!) It seems that you have not defined database folder permissions.<br>
 <br>If you are using a <a href="comersus_backoffice_permissionsForLocal.asp">local computer click here...</a>
 <br>If you are using a <a href="comersus_backoffice_permissionsForHosting.asp">web hosting service click here...</a>
 <br>If you want to obtain <a href="http://www.comersus.org" target="_blank">technical assistance click here...</a>
 <br>The installation Wizard cannot continue until you solve the permission issue<br><br>
<%end if%>

Details: <%=pError%>
<!--#include file="footer.asp"-->