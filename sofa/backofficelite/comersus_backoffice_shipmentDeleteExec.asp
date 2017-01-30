<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/settings.asp"-->    

<% 
on error resume next 

dim mySQL, conntemp, rstemp

pIdShipment	= request.Querystring("idshipment")

if trim(pIdShipment)="" then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Enter a valid shipment")
end if

mySQL="DELETE FROM shipments WHERE idShipment="&pIdShipment

call updateDatabase(mySQL, rstemp2, "comersus_backoffice_shipmentDeleteExec.asp") 

response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Shipment deleted")	
call closeDb()
%>

