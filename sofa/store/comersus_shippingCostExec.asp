<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: resturn off line shipping cost
%>
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/screenMessages.asp" --> 
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/cartFunctions.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/currencyFormat.asp"-->
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/adSenseFunctions.asp"-->


<%
on error resume next

dim mySQL, connTemp, rsTemp


pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanyLogo	 	= getSettingKey("pCompanyLogo")
pMoneyDontRound	 	= getSettingKey("pMoneyDontRound")
pAuctions		= getSettingKey("pAuctions")
pListBestSellers 	= getSettingKey("pListBestSellers")
pNewsLetter 		= getSettingKey("pNewsLetter")
pPriceList 		= getSettingKey("pPriceList")
pStoreNews 		= getSettingKey("pStoreNews")

shipmentTotal		= Cdbl(0)

pIdDbSession	 	= checkSessionData()
pIdDbSessionCart 	= checkDbSessionCartOpen()

pIdCustomer     	= getSessionVariable("idCustomer",0)
pIdCustomerType 	= getSessionVariable("idCustomerType",1)

pShippingStateCode	= getUserInput(request("stateCode"),10)
pShippingState		= getUserInput(request("state"),50)
pShippingCountryCode    = getUserInput(request("countryCode"),10)
pShippingZip		= getUserInput(request("zip"),20)

if (pShippingState<>"" and pShippingStateCode<>"") then 
  response.redirect "comersus_message.asp?message="&Server.UrlEncode(getMsg(422,"You cannot enter state and a non listed state."))
end if

' calculate total price of the order, total weight and product total quantities
pSubTotal  			= Cdbl(calculateCartTotal(pIdDbSessionCart))
pCartTotalWeight 		= Cdbl(calculateCartWeight(pIdDbSessionCart))
pCartQuantity  			= Cdbl(calculateCartQuantity(pIdDbSessionCart))

pCartFreeShippingTotal		= Cdbl(calculateFreeShippingTotal(pIdDbSessionCart))
pCartFreeShippingWeight 	= Cdbl(calculateCartFreeShippingWeight(pIdDbSessionCart))
pCartFreeShippingQuantity	= Cdbl(calculateCartFreeShippingQuantity(pIdDbSessionCart))

' rest Free Shipping values to obtain free shipping
pSubTotal	 =pSubTotal-pCartFreeShippingTotal
pCartTotalWeight =pCartTotalWeight-pCartFreeShippingWeight
pCartQuantity	 =pCartQuantity-pCartFreeShippingQuantity


' use shipping codes
pFilterStateCode	= pShippingStateCode
pFilterCountryCode	= pShippingCountryCode
pFilterZip		= pShippingZip

' if customer use anotherState, insert a dummy state code to simplify SQL sentence
if pFilterStateCode="" then
   pFilterStateCode="**"
end if


mySQL="SELECT distinct quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil, idShipment, shipmentDesc, priceToAdd, percentageToAdd, handlingPercentage, handlingFix FROM shipments, shippingZones,  shippingZonesContents" & _
	  " WHERE (shipments.idShippingZone = shippingZones.idShippingZone) AND (shippingZonesContents.idShippingZone = shippingZones.idShippingZone)AND ((stateCode='" &pFilterStateCode& "') OR (stateCode ='')) AND ((countryCode='"&pFilterCountryCode&"') OR (countryCode ='')) AND ((zip='" &pFilterZip& "') OR (zip ='')) AND (idCustomerType=" &pIdCustomerType& " OR idCustomerType IS NULL OR idCustomerType=0) AND idStore=" &pIdStore
'response.write mysql
'response.end
call getFromDatabase(mySQL, rstemp, "selecthipment")    


if  rsTemp.eof then 
    response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(427,"Sorry, there are no service types loaded for your location. Please contact us to request a shipping quote."))
end if

' load shipment array
reDim shipmentsArray(200,7)
shipmentIndex = 0

do until rstemp.eof 

      ' insert if all rules are ok
      if pCartQuantity>=rstemp("quantityFrom") and pCartQuantity<=rstemp("quantityUntil") and pCartTotalWeight>=rstemp("weightFrom") and pCartTotalWeight<=rstemp("weightUntil") and pSubtotal>=Cdbl(rstemp("priceFrom")) and pSubtotal<=Cdbl(rstemp("priceUntil")) then
	   shipmentsArray(shipmentIndex,0)	= rstemp("idShipment")
	   shipmentsArray(shipmentIndex,1)	= rstemp("shipmentDesc")
	   shipmentsArray(shipmentIndex,2)	= Cdbl(rstemp("priceToAdd"))
	   shipmentsArray(shipmentIndex,3)	= rstemp("percentageToAdd")
           shipmentIndex			= shipmentIndex+1
      end if ' rules end
      
   rstemp.movenext
loop

if shipmentIndex = 0 then
   response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(427,"Sorry, there are no service types loaded for your location. Please contact us to request a shipping quote."))
end if

%>
<!--#include file="header.asp"--> 
<br><b><%=getMsg(425,"s cost")%></b><br><br>
<TABLE BORDER="0" CELLPADDING="0" WIDTH="500">
<TR>
 <TD bgcolor="#ff6600"><b><%=getMsg(428,"service type")%></b><br></TD>
 <TD bgcolor="#ff6600"><b><%=getMsg(429,"price")%></b><br></TD>	
</TR>
<%for f=0 to shipmentIndex-1%>
   <TR>
	<TD><%=shipmentsArray(f,1)%></TD>
	<TD> <% 
	      if shipmentsArray(f,2)>0 or shipmentsArray(f,3)>0 then 
                howMuch = shipmentsArray(f,2) + (shipmentsArray(f,3)*pSubTotal/100)                              
                response.write " " & pCurrencySign & money(howMuch)
              else
                response.write " " & pCurrencySign & money(0)
              end if%>
        </TD>		
  </TR>
<%next%>
<TR><TD>&nbsp;</TD><TD>&nbsp;</TD></TR>
<TR><TD>&nbsp;</TD><TD>&nbsp;</TD></TR>
<TR>
 <TD><a href="comersus_showCart.asp"><%=getMsg(430,"back to cart")%></a></TD>
 <TD>&nbsp;</TD>		
</TR>

<%if pIdCustomer<>"" then%>
 <TR>
  <TD><a href="comersus_shippingCostForm.asp"><%=getMsg(431,"try again")%></a></TD>
  <TD> &nbsp;</TD>				
 </TR>
<%end if%>

<TR><TD>&nbsp;</TD><TD>&nbsp;</TD></TR>

</table>
<!--#include file="footer.asp"-->
<%call closeDb()%>