<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
%>
 
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/dateFunctions.asp"--> 
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/screenMessages.asp"--> 
<!--#include file="../includes/cartFunctions.asp"-->
<!--#include file="../includes/miscFunctions.asp"-->
<!--#include file="../includes/stringFunctions.asp"--> 
<!--#include file="../includes/discountFunctions.asp"--> 
<!--#include file="../includes/encryption.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/paymentMethodFunctions.asp"--> 
<!--#include file="../includes/shippingMethodFunctions.asp"--> 
<!--#include file="../includes/adSenseFunctions.asp"-->
<!--#include file="comersus_optDes.asp"-->    
<!--#include file="comersus_optRealTimeUps.asp"-->  
<!--#include file="comersus_optRealTimeUsps.asp"-->  
<!--#include file="comersus_optRealTimeUspsInt.asp"-->  
<!--#include file="comersus_optRealTimeInterShipper.asp"-->  
<!--#include file="comersus_optRealTimeFedexShipment.asp"-->  
<!--#include file="comersus_optRealTimeCAShipment.asp"-->  

<% 

on error resume next

dim mySQL, connTemp, rsTemp, rsTemp2, rsTemp3

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pMoneyDontRound	 	= getSettingKey("pMoneyDontRound")
pCompany	 	= getSettingKey("pCompany")
pCompanySlogan 		= getSettingKey("pCompanySlogan")
pCompanyAddress		= getSettingKey("pCompanyAddress")
pCompanyZip		= getSettingKey("pCompanyZip")
pCompanyCity		= getSettingKey("pCompanyCity")
pCompanyStateCode	= getSettingKey("pCompanyStateCode")
pCompanyCountryCode	= getSettingKey("pCompanyCountryCode")
pCompanyPhone		= getSettingKey("pCompanyPhone")
pCompanyFax		= getSettingKey("pCompanyFax")
pChangeDecimalPoint	= getSettingKey("pChangeDecimalPoint")
pEncryptionMethod	= getSettingKey("pEncryptionMethod")
pOrderFieldName1	= getSettingKey("orderFieldName1")
pOrderFieldName2	= getSettingKey("orderFieldName2")
pOrderFieldName3	= getSettingKey("orderFieldName3")
pDateSwitch		= getSettingKey("pDateSwitch")
pDateFormat		= getSettingKey("pDateFormat")

pByPassShipping		= getSettingKey("pByPassShipping")
pHeaderKeywords 	= getSettingKey("pHeaderKeywords")
pMinimumPurchase  	= getSettingKey("pMinimumPurchase")

pAuctions		= getSettingKey("pAuctions")
pListBestSellers 	= getSettingKey("pListBestSellers")
pNewsLetter 		= getSettingKey("pNewsLetter")
pPriceList 		= getSettingKey("pPriceList")
pStoreNews 		= getSettingKey("pStoreNews")
pUseVatNumber		= getSettingKey("pUseVatNumber")
pShowNews		= getSettingKey("pShowNews")
pShowSearchBox		= getSettingKey("pShowSearchBox")

pUseShippingAddress	= getSettingKey("pUseShippingAddress")
pStoreClosed		= getSettingKey("pStoreClosed")
pRealTimeShipping	= getSettingKey("pRealTimeShipping")

pIdDbSession	 	= checkSessionData()
pIdDbSessionCart 	= checkDbSessionCartOpen()

pIdCustomer     	= getSessionVariable("idCustomer",0)
pIdCustomerType 	= getSessionVariable("idCustomerType",1)
pDiscountCode	 	= getSessionVariable("discountCode","")
pToken			= getSessionVariable("token","0")

pTime			= getUserInput(request.querystring("time"),12)

' if the store is closed it will not allow checkout
if pStoreClosed="-1" then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(391,"Sorry, your cannot checkout, the store is closed right now."))
end if

if countCartRows(pIdDbSessionCart)=0 then
  response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(392,"The cart is empty, you cannot checkout."))
end if

if Cdbl(calculateCartTotal(pIdDbSessionCart))<Cdbl(pMinimumPurchase) then  
  response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(393,"To checkout you have tu purchase at least: ")& pCurrencySign &money(pMinimumPurchase))
end if


if pIdCustomer=0 then
 response.redirect "comersus_checkoutCustomerForm.asp"
end if

' get customer information

mySQL="SELECT * FROM customers WHERE idCustomer=" &pIdCustomer
call getFromDatabase(mySQL, rstemp, "checkout2")

if not rstemp.eof then
 pName			= rstemp("name")
 pLastName		= rstemp("lastName")
 pCustomerCompany	= rstemp("customerCompany")
 pEmail			= rstemp("email")
 pPassword		= rstemp("password")
 pPhone			= rstemp("phone") 

 pShippingName		= rstemp("shippingName")
 pShippingLastName	= rstemp("shippingLastName")
 pAddress		= rstemp("address")
 pZip			= rstemp("zip")
 pStateCode		= rstemp("stateCode")
 pState			= rstemp("state")
 pCity			= rstemp("city")
 pCountryCode		= rstemp("countryCode")

 pShippingAddress	= rstemp("shippingAddress")
 pShippingZip		= rstemp("shippingZip")
 pShippingStateCode	= rstemp("shippingStateCode")
 pShippingState		= rstemp("shippingState")
 pShippingCity		= rstemp("shippingCity")
 pShippingCountryCode   = rstemp("shippingCountryCode")
end if

if pShippingCountryCode<>"" then
 pFilterCountryCode=pShippingCountryCode
else
 pFilterCountryCode=pCountryCode
end if


' calculate total price of the order, total weight and product total quantities
pSubTotal  		= Cdbl(calculateCartTotal(pIdDbSessionCart))
pCartTotalWeight 	= Cdbl(calculateCartWeight(pIdDbSessionCart))
pCartQuantity  		= Cdbl(calculateCartQuantity(pIdDbSessionCart))

' populate payment methods
pStringPaymentMethods	= createArrayPayments(pCartQuantity, pCartTotalWeight, pSubTotal)

if pToken<>"0" then

   ' get Express Checkout ID
   
   mySQL="SELECT idPayment, paymentDesc FROM payments WHERE redirectionUrl='comersus_gatewayPayPalExpress3.asp' AND idStore=" &pIdStore
   call getFromDatabase(mySQL, rstemp, "checkout2")
   
   pPaymentSurchargeAmount=0
   
   pStringPaymentMethods=rstemp("paymentDesc")&"&"&pPaymentSurchargeAmount&"&"&rstemp("idPayment")&"|"
   
end if

if pStringPaymentMethods="" then 
  response.redirect "comersus_supportError.asp?error="&Server.Urlencode("There are no payments available. Please contact us")
end if

pArrayPaymentMethods	= split(pStringPaymentMethods,"|")

' populate shipment methods

' calculate Free Shipping Items inside the cart
pCartFreeShippingTotal		= Cdbl(calculateFreeShippingTotal(pIdDbSessionCart))
pCartFreeShippingWeight 	= Cdbl(calculateCartFreeShippingWeight(pIdDbSessionCart))
pCartFreeShippingQuantity	= Cdbl(calculateCartFreeShippingQuantity(pIdDbSessionCart))

Select Case lcase(pRealTimeShipping)
   
   Case "ups"
     pStringShipmentMethods		= createArrayShipmentsUps(pCartQuantity-pCartFreeShippingQuantity, pCartTotalWeight-pCartFreeShippingWeight, pSubTotal-pCartFreeShippingTotal)
   
   Case "usps" 
   
    if pFilterCountryCode="US" then
     pStringShipmentMethods		= createArrayShipmentsUsps(pCartQuantity-pCartFreeShippingQuantity, pCartTotalWeight-pCartFreeShippingWeight, pSubTotal-pCartFreeShippingTotal)
    else
     pStringShipmentMethods		= createArrayShipmentsUspsInt(pCartQuantity-pCartFreeShippingQuantity, pCartTotalWeight-pCartFreeShippingWeight, pSubTotal-pCartFreeShippingTotal)
    end if
    
   Case "intershipper" 
     pStringShipmentMethods		= createArrayShipmentsIntershipper(pCartQuantity-pCartFreeShippingQuantity, pCartTotalWeight-pCartFreeShippingWeight, pSubTotal-pCartFreeShippingTotal)
   
   Case "fedex" 
     pStringShipmentMethods		= createArrayShipmentsFedex(pCartQuantity-pCartFreeShippingQuantity, pCartTotalWeight-pCartFreeShippingWeight, pSubTotal-pCartFreeShippingTotal)
   
   Case "ca" 
     pStringShipmentMethods		= createArrayShipmentsCA(pCartQuantity-pCartFreeShippingQuantity, pCartTotalWeight-pCartFreeShippingWeight, pSubTotal-pCartFreeShippingTotal)

   Case Else 
     pStringShipmentMethods		= createArrayShipments(pCartQuantity-pCartFreeShippingQuantity, pCartTotalWeight-pCartFreeShippingWeight, pSubTotal-pCartFreeShippingTotal)     
     
End Select

if pStringShipmentMethods="" then 
  response.redirect "comersus_supportError.asp?error="&Server.Urlencode("There are no shipments available. Please contact us")
end if

pArrayShipmentMethods	= split(pStringShipmentMethods,"|")

' retrieve discount

pDiscountDesc	=""
pDiscountAmount	=0

call getDiscount(pDiscountCode, pDiscountDesc, pDiscountAmount) 

pSubTotal=pSubTotal-pDiscountAmount

%> 
<!--#include file="header.asp"-->
<form method="post" action="comersus_checkOut3.asp" name="saveorder">
<table width="100%" border="0" cellspacing="0">
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4"> 
     <b><%=getMsg(58,"checkout")%></b>
     
     <%if pTime<>"" then%>
     <br><br><i><%=getMsg(681,"Reloaded")%></i>
     <%end if%>
    </td>
  </tr>
  <tr> 
    <td colspan="4"> 
      <hr>
    </td>
  </tr>
  <tr> 
    <td colspan="4">
    <%=getMsg(411,"billing")%><br><br>
    <b><%=getMsg(448,"name")%>: </b></b><%=pName&" "&pLastName%> 
    
    <%if pCustomerCompany<>"" then%>
     <br><b><%=getMsg(449,"cpany")%>: </b> </b><%=pCustomerCompany%>
    <%end if%>
    
    <br><b><%=getMsg(450,"phone")%>: </b></b> <%=pPhone%><br>
      
      <b><%=getMsg(451,"addr")%>: </b></b> <%=pAddress%>&nbsp;<%=pCity%>, <%=pState & pStateCode%>&nbsp;<%=pZip%>,  
      &nbsp;<%=getCountryName(pCountryCode)%><br>        
     
     <br><a href="comersus_customerModifyForm.asp?redirect=checkout">[<%=getMsg(185,"modify")%>]</a>     
      
     <%if pUseShippingAddress="-1" and pByPassShipping="0" then%>
     
       <br><br><%=getMsg(420,"shipping")%><br>
       
       <%if trim(pShippingAddress)="" or isNull(pShippingAddress) then%>
        <br><b><%=getMsg(471,"ship address")%>: </b></b> <%=getMsg(472,"same as bill")%></b>
       <%else%>
        <br><b><%=getMsg(448,"name")%>: </b></b><%=pShippingName&" "&pShippingLastName%> 
        <br><b><%=getMsg(471,"ship address")%>: </b> <%=pShippingAddress%>&nbsp;<%=pShippingCity%>,&nbsp;<%=pShippingState & pShippingStateCode%>&nbsp; <%=pShippingZip%>
        &nbsp;<%=pShippingCountryCode%>
      <%end if%>
      
      <br><br><a href="comersus_customerModifyShippingForm.asp">[<%=getMsg(185,"modify")%>]</a>     
       
     <%end if%>                
     
     </td>
  </tr>
  <tr> 
    <td colspan="4" height="20"> 
      <hr>
    </td>
  </tr>
  <tr> 
    <td width="10%"> <b><%=getMsg(453,"qty")%></b></td>
    <td width="50%"><b><%=getMsg(454,"item")%></b></td>
    <td width="30%"><b><%=getMsg(455,"variat")%></b></td>
    <td width="10%"> <b><%=getMsg(456,"price")%></b></td>
  </tr>
  <%
  ' iterate through cart
  mySQL="SELECT idCartRow, cartRows.idProduct, quantity, unitPrice, description, sku, deliveringTime, personalizationDesc FROM cartRows, products WHERE cartRows.idProduct=products.idProduct AND cartRows.idDbSessionCart="&pIdDbSessionCart

  call getFromDatabase(mySQL, rstemp, "checkout2")       

  do while not rstemp.eof
   pIdCartRow		= rstemp("idCartRow")
   pIdProduct		= rstemp("idProduct")
   pQuantity		= rstemp("quantity")
   pUnitPrice		= Cdbl(rstemp("unitPrice"))   
   pDescription		= rstemp("description")
   pSku			= rstemp("sku")
   pPersonalizationDesc	= rstemp("personalizationDesc")      
 %> 
  <tr> 
    <td width="10%"><%=pQuantity%></td>
    <td width="50%"><%=pSku & " "&pDescription%>
    <%if pPersonalizationDesc<>"" then response.write "&nbsp;(" &pPersonalizationDesc& ")"%>
    </td>
    <td width="30%"> 
    <%=getCartRowOptionals(pIdCartRow)%></td>
    <td width="10%">            
     <%=pCurrencySign &  money(getCartRowPrice(pIdCartRow)) %>      
    </td>
  </tr>
  <%rstemp.movenext
  loop%> 
  <tr> 
   <td colspan="4"> 
   <br>
   <a href="comersus_showCart.asp">[<%=getMsg(185,"modify")%>]</a>
   <hr>
   <br>     
    </td>
  </tr>  

  <tr> 
    <td colspan="2">&nbsp;</td>
    <td><b><%=getMsg(459,"discount")%></b></td>
    <td> - <%=pCurrencySign & money(pDiscountAmount)%>
    </td>
  </tr>  
  
  <tr> 
    <td colspan="2">&nbsp;</td>
    <td><b><%=getMsg(465,"subtotal")%></b></td>
    <td> <%=pCurrencySign & money(pSubTotal)%>
    </td>
  </tr>

 
 <tr> 
    <td colspan="4">&nbsp;</td>    
  </tr>
    
  <tr> 
    <td colspan="3">
    <br><b><%=getMsg(457,"payment")%></b> <%=pPaymentDesc%></td>
    <td> 
      <div align="left">
  <select name="idPaymentArray" size="1">
  <%for f=0 to uBound(pArrayPaymentMethods)-1%> 
  
  <%
  if instr(pArrayPaymentMethods(f),"&")=0 then
   response.redirect "comersus_supportError.asp?error="&Server.Urlencode("Invalid payment method selection array")
  end if
  pArrayOnePaymentMethod	=split(pArrayPaymentMethods(f),"&")
  pPaymentMethodDescription	=pArrayOnePaymentMethod(0)
  pPaymentMethodSurcharge	=pArrayOnePaymentMethod(1)
  if pPaymentMethodSurcharge>0 then
   pPaymentMethodDescription=pPaymentMethodDescription&" " & pCurrencySign & money(pPaymentMethodSurcharge)
  end if
  %>
          <option value="<%=f%>"> <%=pPaymentMethodDescription%></option>
  <%next%> 
 </select>              
      </div>
    </td>
  </tr>

  <%if pByPassShipping="0" then%>
  <tr> 
    <td colspan="3"><b><%=getMsg(458,"ship method")%></b> <%=pShipmentDesc%></td>
    <td> 
      <div align="left">
      
  <select name="idShipmentArray" size="1">
  <%for f=0 to uBound(pArrayShipmentMethods)-1%> 
  
  <%
  if instr(pArrayShipmentMethods(f),"&")=0 then
   response.redirect "comersus_supportError.asp?error="&Server.Urlencode("Invalid shipment method selection array")
  end if
  pArrayOneShipmentMethod	=split(pArrayShipmentMethods(f),"&")
  pShipmentMethodDescription	=pArrayOneShipmentMethod(0)
  pShipmentMethodAmount		=pArrayOneShipmentMethod(1)
  pShipmentMethodHandling	=pArrayOneShipmentMethod(2)
  if Cdbl(pShipmentMethodAmount)>0 then
   pShipmentMethodDescription=pShipmentMethodDescription&" " & pCurrencySign & money(pShipmentMethodAmount)
  end if
  
  if Cdbl(pShipmentMethodHandling)>0 then
   pShipmentMethodDescription=pShipmentMethodDescription&" handling " & pCurrencySign & money(pShipmentMethodHandling)
  end if
  
  %>
          <option value="<%=f%>"> <%=pShipmentMethodDescription%></option>
  <%next%> 
 </select>   
      
      </div>
    </td>
  </tr>
 <%end if%>
 
 <%if pUseVatNumber="-1" then%>
   <tr>
    <td colspan="3">
    <br><b><%=getMsg(733,"VAT")%></b> </td>
    <td> 
      <div align="left">
      <input type=text name="vatNumber" length=20 maxlength=20>
      </div>
    </td>
  </tr>  
 <%end if%>
 
  <%if pOrderFieldName1<>"" then%>
    <tr> 
      <td colspan=3> <b><%=pOrderFieldName1%></b></td>
      <td>         
       <input type=text name=user1>       
      </td>
    </tr>
    <%end if%>
    
    <%if pOrderFieldName2<>"" then%>
    <tr> 
      <td colspan=3> <b><%=pOrderFieldName2%></b></td>
      <td>         
       <input type=text name=user2>       
      </td>
    </tr>
    <%end if%>
    
    <%if pOrderFieldName3<>"" then%>
    <tr> 
      <td colspan=3> <b><%=pOrderFieldName3%></b> </td>
      <td >         
       <input type=text name=user3>       
      </td>
    </tr>
    <%end if%>
    
  <tr>
    <td colspan="3">
    <br><b><%=getMsg(444,"comments")%></b> </td>
    <td> 
      <div align="left">
      <textarea name="comments" cols="35" maxlength=500></textarea>
      </div>
    </td>
  </tr>  
  
  <tr>
    <td colspan="4">&nbsp;</td>
  </tr> 
  
  <tr> 
    <td colspan="4"> 
      <div align="center"> 
        <input type="submit" name="Save" value="<%=getMsg(421,"continue")%>">
      </div>
    </td>
  </tr>
</table>
</form>
<br><br>
<!--#include file="footer.asp"-->
<%call closeDb()%>