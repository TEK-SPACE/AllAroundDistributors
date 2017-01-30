<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"-->  
<!--#include file="header.asp"-->
<!--#include file="./includes/settings.asp"-->

<br><b>SQL Query</b><br>

<form method="post" action="comersus_backoffice_genericSqlexec.asp" name="gsql" target="_blank">
 Sentence<br>
 <textarea name="sqlsentence" cols="60" rows="5"><%=request.querystring("mySQL")%></textarea>    
 <br><br><input type="submit" name="Submit" value="Execute">
</form>
<!--#include file="footer.asp"--> 
