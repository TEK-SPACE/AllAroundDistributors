<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"-->   
<!--#include file="../includes/settings.asp" -->  
<!--#include file="../includes/stringFunctions.asp"-->

<% 
on error resume next 

dim mySQL, conntemp, rstemp

pZoneDesc		= getUserInput(request.form("zoneDesc"),100)




' validate
if pZoneDesc="" then
 response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Please enter a valid zone description")	
end if

' insert shipment in to db
mySQL="INSERT INTO shippingZones(zoneName) VALUES ('"&pZoneDesc&"')"

call updateDatabase(mySQL, rstemp, "comersus_backoffice_addShipmentZoneExec.asp") 

mySQL="SELECT Max(idShippingZone) as MaxId FROM shippingZones WHERE zoneName = '"&pZoneDesc&"'"
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_listZonesContents.asp") 

if not rstemp.eof then
	response.redirect "comersus_backoffice_listZoneAndContent.asp?idZone="&cInt(rstemp("MaxId"))
else
	response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Error in addShipmentZoneExec. Zone wasn't added.")	
end if

call closeDb()


%>