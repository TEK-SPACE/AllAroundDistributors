
<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"-->  
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/currencyFormat.asp"-->    
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/getSettingKey.asp"-->  
<!--#include file="../includes/stringFunctions.asp" --> 

<%
on error resume next

dim mySQL, conntemp, rstemp


pIdZone	= getUserInput(request("idZone"),100)

MySQL = "SELECT * FROM shipments WHERE idShippingZone = "& pIdZone	
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_shipmentZoneDeleteExec.asp") 

If not rstemp.eof Then
	response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Can't delete the zone, it has shipments related. Delete those shipments first.")
end if

' Delete Zone Contents
mySQL="DELETE FROM shippingZonesContents WHERE idShippingZone = " & pIdZone

call getFromDatabase(mySQL, rstemp, "comersus_backoffice_shipmentZoneDeleteExec.asp") 

' Delete Zone 

mySQL="DELETE FROM shippingZones WHERE idShippingZone = " & pIdZone

call getFromDatabase(mySQL, rstemp, "comersus_backoffice_shipmentZoneDeleteExec.asp") 

response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Zone deleted!")

%>

