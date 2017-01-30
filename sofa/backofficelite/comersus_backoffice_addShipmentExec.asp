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

<% 
on error resume next 

dim mySQL, conntemp, rstemp

pShipmentDesc		= request.form("shipmentDesc")
pPriceToAdd		= request.form("priceToAdd")
pPercentageToAdd	= request.form("percentageToAdd")
pShipmentTime		= request.form("shipmentTime")

pQuantityFrom		= request.form("quantityFrom")
pQuantityUntil		= request.form("quantityUntil")

pWeightFrom		= request.form("WeightFrom")
pWeightUntil		= request.form("WeightUntil")

pPriceFrom		= request.form("priceFrom")
pPriceUntil 		= request.form("priceUntil")

pIdZone 		= request.form("shippingZone")


pIdCustomerType		= request.form("customerType")


' validate
if pShipmentDesc="" then
 response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Please enter a valid shipment description")	
end if

' insert shipment in to db
mySQL="INSERT INTO shipments (idStore, shipmentDesc, priceToAdd, percentageToAdd, shipmentTime, quantityFrom, quantityUntil, WeightFrom, WeightUntil, priceFrom, priceUntil, idShippingZone, idCustomerType) VALUES ("&pIdStore&",'" &pshipmentDesc & "',"& ppriceToAdd &","& ppercentageToAdd &",'"&pshipmentTime&"',"&pquantityFrom&","&pquantityUntil&","&pWeightFrom&","&pWeightUntil&","&ppriceFrom& "," &ppriceUntil& ","&pIdZone&", " &pidCustomerType& ")"

call updateDatabase(mySQL, rstemp, "comersus_backoffice_addShipmentExec.asp") 

call closeDb()

response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Shipment added")	

%>