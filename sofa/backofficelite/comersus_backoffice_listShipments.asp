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
<!--#include file="../includes/miscFunctions.asp"-->  

<%
on error resume next

dim mySQL, conntemp, rstemp

pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")

' get all shipments
mySQL="SELECT distinct idShipment, shipmentDesc, priceToAdd, percentageToAdd, shipmentTime, zoneName FROM shipments, shippingZones WHERE shipments.idShippingZone = shippingZones.idShippingZone ORDER BY shipmentDesc"

call getFromDatabase(mySQL, rstemp, "comersus_backoffice_listShipments.asp") 

if  rstemp.eof then 
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("No shipments in the database")
end if  

%>

<!--#include file="header.asp"-->
<!--#include file="./includes/settings.asp"--> 

<br><b>Shipments</b><br><br>
<%
 do while not rstemp.eof

   pIdshipment		= rstemp("idshipment")
   pShipmentDesc	= rstemp("shipmentDesc")   
   pPriceToAdd		= rstemp("priceToAdd")
   pPercentageToAdd	= rstemp("percentageToAdd")
   pShipmentTime	= rstemp("shipmentTime")
   pZoneName		= rstemp("zoneName")   
   %> 
   <br> 
   <%=pShipmentDesc%> [<%=pCurrencySign & money(pPriceToAdd)%> <%if pPercentageToAdd<>0 then response.write " "&pPercentageToAdd&"%"%>]  <a href="comersus_backoffice_shipmentDeleteExec.asp?idShipment=<%=pIdShipment%>">Delete</a>
   <br>
   Zone:&nbsp;<%=pZoneName%><br>
   
   <%  
   count = count + 1
   rstemp.MoveNext
 loop
%> 

<!--#include file="footer.asp"-->
