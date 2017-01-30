<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<%
' verifies if admin is logged, caso not send to login page
if Session("admin")=0 then
 response.redirect "comersus_backoffice_notLogged.asp"
end if
%>