<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"-->   
<!--#include file="../includes/settings.asp"-->   
<!--#include file="../includes/databaseFunctions.asp"-->   

<% 
on error resume next 

if pBackOfficeDemoMode=-1 then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Function disabled in demo mode")
end if

dim  mySQL, conntemp, rstemp

' form parameter
pIdPayment	= request.Querystring("idPayment")

if trim(pIdPayment)="" then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Please enter a valid payment ID")
end if

' delete from db
mySQL="DELETE FROM payments WHERE idPayment=" &pIdPayment

call updateDatabase(mySQL, rstemp, "comersus_backoffice_deletePaymentExec.asp") 

call closeDb()

response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Payment removed")
%>
