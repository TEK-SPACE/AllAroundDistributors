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

pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")


'update shipping zone
pIdZone		= getUserInput(request.form("idZone"),50)
pZoneName	= getUserInput(request.form("zoneName"),50)

MySQL	= "UPDATE shippingZones SET zoneName='"&pZoneName&"' WHERE idShippingZone=" & pIdZone
call updateDatabase(mySQL, rstemp, "comersus_backoffice_listShipmentZones.asp") 
response.redirect "comersus_backoffice_listZoneAndContent.asp?IdZone="&pIdZone


%>

