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


pIdZone			= getUserInput(request("idZone"),10)

'update shipping zone
pZip = getUserInput(request("zip"),50)
pStateCode = getUserInput(request("statecode"),10)
pCountryCode = getUserInput(request("CountryCode"),10)

MySQL	= "INSERT INTO shippingZonesContents (idShippingZone, zip, stateCode, countryCode) VALUES ("&pIdZone&",'"&pZip&"', '"&pStateCode&"', '"&pCountryCode&"')"
call updateDatabase(mySQL, rstemp, "comersus_backoffice_modifyZonesContents.asp") 
response.redirect "comersus_backoffice_listZoneAndContent.asp?IdZone="&pIdZone


%>