<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' USA - 2006
' http://www.comersus.com 
' Details: off line shipment method functions
%>

<%
' Off Line Shipments

function createArrayShipments(pCartQuantity, pCartTotalWeight, pSubTotal)

' retrieve available shipment methods, load a string array with | as row separator

dim mySQL, rstemp, pFilterStateCode, pFilterCountryCode, pFilterZip

pIdCustomer 		= getSessionVariable("idCustomer",0)

pIdDbSession	 	= checkSessionData()
pIdDbSessionCart 	= checkDbSessionCartOpen()

pShippingDefaultFee		= getSettingKey("pShippingDefaultFee")
pShippingDefaultServiceType	= getSettingKey("pShippingDefaultServiceType")


createArrayShipments	=""

' get customer address

mySQL="SELECT zip, stateCode, state, countryCode, shippingZip, shippingStateCode, shippingState, shippingCountryCode FROM customers WHERE idCustomer="&pIdCustomer

call getFromDatabase(mySQL, rstemp, "checkOutOffLineShipment")    

if not rstemp.eof then

 pStateCode		= rstemp("stateCode")
 pState			= rstemp("state")
 pCountryCode		= rstemp("countryCode")
 pZip			= rstemp("zip")
 
 pShippingStateCode	= rstemp("shippingStateCode")
 pShippingState		= rstemp("shippingState")
 pShippingCountryCode   = rstemp("shippingCountryCode")
 pShippingZip		= rstemp("shippingZip")
 
else
 ' error, cannot find customer address
 response.redirect "comersus_supportError.asp?error="&Server.Urlencode("Cannot retrieve customer address in shippingMethodFunction")
end if

if pShippingCountryCode<>"" then
 ' use shipping codes
 pFilterStateCode	= pShippingStateCode
 pFilterCountryCode	= pShippingCountryCode
 pFilterZip		= pShippingZip
else
 ' use billing
 pFilterStateCode	= pStateCode
 pFilterCountryCode	= pCountryCode 
 pFilterZip		= pZip
end if

' if customer use anotherState, insert a dummy state code to simplify SQL sentence
if pFilterStateCode="" then
   pFilterStateCode="**"
end if

' get shipments depending on pStateCode, pCountryCode and pZip

mySQL="SELECT DISTINCT quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil, idShipment, shipmentDesc, priceToAdd, percentageToAdd, handlingPercentage, handlingFix FROM shipments, shippingZones,  shippingZonesContents WHERE (shipments.idShippingZone = shippingZones.idShippingZone) AND (shippingZonesContents.idShippingZone = shippingZones.idShippingZone)AND ((stateCode='" &pFilterStateCode& "') OR (stateCode ='') OR (stateCode IS NULL)) AND ((countryCode='"&pFilterCountryCode&"') OR (countryCode ='') OR (countryCode IS NULL)) AND ((zip='" &pFilterZip& "') OR (zip ='') OR (zip IS NULL)) AND (idCustomerType=" &pIdCustomerType& " OR idCustomerType IS NULL OR idCustomerType=0) AND idStore=" &pIdStore
call getFromDatabase(mySQL, rstemp, "selecthipment")    

if  rstemp.eof then 
 
 ' load default
 
 shipmentDesc	= pShippingDefaultServiceType
 shippingAmount	= pShippingDefaultFee
 handlingAmount	= 0
 createArrayShipments	= createArrayShipments &   shipmentDesc & "&" & shippingAmount & "&" & handlingAmount & "|"   	   	
	   
end if

' create string

do until rstemp.eof 

      ' insert if all rules are ok
      if pCartQuantity>=rstemp("quantityFrom") and pCartQuantity<=rstemp("quantityUntil") and pCartTotalWeight>=rstemp("weightFrom") and pCartTotalWeight<=rstemp("weightUntil") and pSubtotal>=Cdbl(rstemp("priceFrom")) and pSubtotal<=Cdbl(rstemp("priceUntil")) then	             
          
           ' shipping amount          
           shippingAmount 	= rstemp("priceToAdd") + (rstemp("percentageToAdd")*pSubTotal/100)                                                        
                    
          ' handling amount
           handlingAmount	= rstemp("handlingFix") + (rstemp("handlingPercentage")*pSubTotal/100)                    
	   
	   shipmentDesc		= rstemp("shipmentDesc")

	   createArrayShipments	= createArrayShipments &   shipmentDesc & "&" & shippingAmount & "&" & handlingAmount & "|"
	   
      end if ' rules end
      
   rstemp.movenext
loop

if createArrayShipments="" then
 
 ' load default
 shipmentDesc	= pShippingDefaultServiceType
 shippingAmount	= pShippingDefaultFee
 handlingAmount	= 0
 createArrayShipments	= createArrayShipments &   shipmentDesc & "&" & shippingAmount & "&" & handlingAmount & "|"   	   	
 
end if
 
' save shipment array in the database

mySQL="UPDATE dbSession set sessionData='" &createArrayShipments& "' WHERE idDbSession=" &pIdDbSession
call updateDatabase(mySQL, rstemp, "orderVerify")
 
end function

%>

