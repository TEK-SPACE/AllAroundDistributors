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
<!--#include file="../includes/taxFunctions.asp"--> 
<!--#include file="../includes/paymentMethodFunctions.asp"--> 
<!--#include file="../includes/shippingMethodFunctions.asp"--> 
<!--#include file="../includes/getFastTax.asp"--> 
<!--#include file="../includes/adSenseFunctions.asp"-->
<!--#include file="comersus_optDes.asp"-->    

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
pByPassShipping		= getSettingKey("pByPassShipping")
pBonusPointsPerPrice	= getSettingKey("pBonusPointsPerPrice")
pOffLineCardsAccepted	= getSettingKey("pOffLineCardsAccepted")
pHeaderKeywords 	= getSettingKey("pHeaderKeywords")
pRealTimeTaxSOLicense	= getSettingKey("pRealTimeTaxSOLicense")

pAuctions		= getSettingKey("pAuctions")
pListBestSellers 	= getSettingKey("pListBestSellers")
pNewsLetter 		= getSettingKey("pNewsLetter")
pPriceList 		= getSettingKey("pPriceList")
pStoreNews 		= getSettingKey("pStoreNews")
pUseShippingAddress	= getSettingKey("pUseShippingAddress")
pShippingTaxable	= getSettingKey("pShippingTaxable")
pTaxIncluded		= getSettingKey("pTaxIncluded")
pShowNews		= getSettingKey("pShowNews")
pShowSearchBox		= getSettingKey("pShowSearchBox")
pCalculateTaxIncluded	= getSettingKey("pCalculateTaxIncluded")

pPaymentSurcharge	= 0
pBonusPoints		= 0

pIdDbSession	 	= checkSessionData()
pIdDbSessionCart 	= checkDbSessionCartOpen()

pIdCustomer     	= getSessionVariable("idCustomer",0)
pIdCustomerType 	= getSessionVariable("idCustomerType",1)
pDiscountCode 		= getSessionVariable("discountCode","")
pToken			= getSessionVariable("token",0)
pBonusPointsUsed	= getSessionVariable("bonusPointsUsed",0)

' get shipping and payment selection
pIdShipmentArray	= getUserInput(request.form("idShipmentArray"),20)
pIdPaymentArray		= getUserInput(request.form("idPaymentArray"),20)
pComments		= getUserInput(request.form("comments"),500)
pVatNumber		= getUserInput(request.form("vatNumber"),30)
pUser1			= getUserInput(request.form("user1"),50)
pUser2			= getUserInput(request.form("user2"),50)
pUser3			= getUserInput(request.form("user3"),50)

' verify required datak, if something is missing, redirect to checkout1 step
if (pIdShipmentArray="" or isNumeric(pIdShipmentArray)=false) and pByPassShipping="0" then
 response.redirect "comersus_supportError.asp?error="&Server.Urlencode("Missing or invalid Shipment array data in checkout3") 
end if 

if (pIdPaymentArray="" or isNumeric(pIdPaymentArray)=false) and pBonusPointsUsed=0 then
 response.redirect "comersus_supportError.asp?error="&Server.Urlencode("Missing or invalid Shipment array data in checkout3") 
end if

' if cart has no items cancel the order
if countCartRows(pIdDbSessionCart)=0 then 
 response.redirect "comersus_supportError.asp?error="&Server.Urlencode("No items in cart at checkout3 step for customer with email "&pEmail)
end if

' calculate total price of the order, total weight and product total quantities
pSubTotal  		= Cdbl(calculateCartTotal(pIdDbSessionCart))
pCartTotalWeight 	= Cdbl(calculateCartWeight(pIdDbSessionCart))
pCartQuantity  		= Cdbl(calculateCartQuantity(pIdDbSessionCart))

pStringPaymentMethods	= createArrayPayments(pCartQuantity, pCartTotalWeight, pSubTotal)

if pStringPaymentMethods="" then 
   response.redirect "comersus_supportError.asp?error="&Server.Urlencode("There are no payments available. Please contact us")
end if

' populate payments
pArrayPaymentMethods	= split(pStringPaymentMethods,"|")
pArrayOnePayment	= split(pArrayPaymentMethods(pIdPaymentArray),"&")
pPaymentDescription	= pArrayOnePayment(0)
pPaymentSurcharge	= pArrayOnePayment(1)
pIdPayment		= pArrayOnePayment(2)

' populate shipment methods from dbSessionCart

if pByPassShipping="0" then

 mySQL="SELECT sessionData FROM dbSession WHERE idDbSession=" &pIdDbSession
 call getFromDatabase(mySQL, rstemp, "checkout3")
 
 if rstemp.eof then
  response.redirect "comersus_supportError.asp?error="&Server.Urlencode("Cannot get dbSession shipment data at checkout 3 for "&pIdDbSession)
 end if
 
 pStringShipmentMethods		=rstemp("sessionData") 

 ' verify shipment string array
 
 if pStringShipmentMethods="" then 
   response.redirect "comersus_supportError.asp?error="&Server.Urlencode("There are no shipments available. Please contact us")
 end if


 if instr(pStringShipmentMethods,"|")=0 then  
  response.redirect "comersus_checkout2.asp?time="&Time
 end if

 pArrayShipmentMethods	= split(pStringShipmentMethods,"|")
 pArrayOneShipment	= split(pArrayShipmentMethods(pIdShipmentArray),"&")

 ' verify shipment string array
 if uBound(pArrayOneShipment)<>2 then  
  response.redirect "comersus_checkout2.asp?time="&Time
 end if

 pShipmentDescription	= pArrayOneShipment(0)
 pShipmentCharge	= Cdbl(pArrayOneShipment(1))+Cdbl(pArrayOneShipment(2))

else
 pShipmentDescription	="N/A"
 pShipmentCharge	=0
end if

' get customer information

mySQL="SELECT * FROM customers WHERE idCustomer=" &pIdCustomer
call getFromDatabase(mySQL, rstemp, "orderVerify")

if not rstemp.eof then
 pName			= rstemp("name")
 pLastName		= rstemp("lastName")
 pCustomerCompany	= rstemp("customerCompany")
 pEmail			= rstemp("email")
 pPassword		= rstemp("password")
 pPhone			= rstemp("phone") 

 pAddress		= rstemp("address")
 pZip			= rstemp("zip")
 pStateCode		= rstemp("stateCode")
 pState			= rstemp("state")
 pCity			= rstemp("city")
 pCountryCode		= rstemp("countryCode")

 pShippingName		= rstemp("shippingName")
 pShippingLastName	= rstemp("shippingLastName")
 pShippingAddress	= rstemp("shippingAddress")
 pShippingZip		= rstemp("shippingZip")
 pShippingStateCode	= rstemp("shippingStateCode")
 pShippingState		= rstemp("shippingState")
 pShippingCity		= rstemp("shippingCity")
 pShippingCountryCode   = rstemp("shippingCountryCode")
end if


pDiscountDesc	=""
pDiscountAmount	=0

call getDiscount(pDiscountCode, pDiscountDesc, pDiscountAmount) 

' bonus points


if pRealTimeTaxSOLicense="" or pRealTimeTaxSOLicense="0" then
 pTaxAmount		= calculateTax(pShipmentCharge, pSubTotal, pVatNumber)
else
 pTaxAmount		= pSubTotal*Cdbl(getLiveSalesTax("90210", pRealTimeTaxSOLicense)) 
end if

pTotal 			= pSubTotal + pTaxAmount + pPaymentSurcharge + pShipmentCharge - pDiscountAmount

if pPaymentSurcharge>0 then
 pPaymentDetails=pPaymentDescription & " " &getMsg(726,"surcharge")& " " & pCurrencySign &money(pPaymentSurcharge) 
else
 pPaymentDetails=pPaymentDescription
end if

pShipmentDetails=pShipmentDescription& " " &pCurrencySign & money(pShipmentCharge) 

' save order data in dbSession
pSessionData=pShipmentDetails&"||"&pPaymentDetails&"||"&pComments&"||"&pBonusPoints&"||"&pVatNumber&"||"&pUser1&"||"&pUser2&"||"&pUser3&"||"& pDiscountCode &"||"& pDiscountAmount &"||"& pTaxAmount &"||"& pTotal &"||" &pIdPayment

mySQL="UPDATE dbSession set sessionData='" &pSessionData& "' WHERE idDbSession=" &pIdDbSession
call updateDatabase(mySQL, rstemp, "orderVerify")

pBonusPointsAvailable=getBonusPoints(pIdCustomer)

%> 
<!--#include file="header.asp"-->
<form method="post" action="comersus_checkOutSaveOrder.asp" name="saveorder">
<table width="100%" border="0" cellspacing="0">
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4"> 
      <div align="center"><b><font size="4"><%=pCompany%></font></b><br>
        <font size="1"><%=pCompanyAddress%><br>
        <%=pCompanyCity&", "&pCompanyStateCode&" "& pCompanyZip &" "& pCompanyCountryCode%></font><br><br>
        <hr>
        </div>
    </td>
  </tr>
  <tr> 
    <td colspan="2">&nbsp;</td>
    <td colspan="2"><b><%=getMsg(447,"date")%></b>: <%=formatDate(fixDate(Date))%></td>
  </tr>
  <tr> 
    <td colspan="4"> 
      <hr>
    </td>
  </tr>
  <tr> 
    <td colspan="4"><b><%=getMsg(448,"name")%>: </b><%=pName&" "&pLastName%> - <b><%=getMsg(449,"cpany")%>: </b> </b><%=pCustomerCompany%> <br>
      <b><%=getMsg(450,"phone")%>: </b></b> <%=pPhone%><br>
      
      <b><%=getMsg(451,"addr")%>: </b></b> <%=pAddress%>&nbsp;<%=pCity%>, <%=pState & pStateCode%>&nbsp;<%=pZip%> 
      &nbsp;<%=pCountryCode%><br>
      
     <%if pUseShippingAddress="-1" and pShippingAddress<>"" then%>
       <%if trim(pShippingAddress)="" then%>
        <b><%=getMsg(471,"ship address")%>: </b></b> <%=getMsg(472,"same as bill")%></b>
       <%else%>
        <b><%=getMsg(471,"ship address")%>: </b> (<%=pShippingName&" "&pShippingLastName%>) <%=pShippingAddress%>&nbsp;<%=pShippingCity%>,&nbsp;<%=pShippingState & pShippingStateCode%>&nbsp; <%=pShippingZip%>
        &nbsp;<%=pShippingCountryCode%>
      <%end if%>
     <%end if%> 
     
      <br><b><%=getMsg(452,"comments")%></b>: <%=pComments%>
      
      <%if pOrderFieldName1<>"" then%>
       <br><b><%=pOrderFieldName1%>: </b><%=pUser1%>
      <%end if%>
      <%if pOrderFieldName2<>"" then%>
       <br><b><%=pOrderFieldName2%>: </b><%=pUser2%>
      <%end if%>
      <%if pOrderFieldName3<>"" then%>
       <br><b><%=pOrderFieldName3%>: </b><%=pUser3%>
      <%end if%>
      
      </td>
  </tr>
     
  <tr> 
    <td colspan="4" height="20"> 
      <hr>
  </td>  
  </tr>
  <tr> 
    <td width="40"> <b><%=getMsg(453,"qty")%></b></td>
    <td width="189"><b><%=getMsg(454,"item")%></b></td>
    <td width="148"><b><%=getMsg(455,"variat")%></b></td>
    <td width="76"> <b><%=getMsg(456,"price")%></b></td>
  </tr>
  <%
  ' iterate through cart
  mySQL="SELECT idCartRow, cartRows.idProduct, quantity, unitPrice, description, sku, deliveringTime, personalizationDesc FROM cartRows, products WHERE cartRows.idProduct=products.idProduct AND cartRows.idDbSessionCart="&pIdDbSessionCart

  call getFromDatabase(mySQL, rstemp, "orderVerify")       

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
    <td width="40"><%=pQuantity%></td>
    <td width="189"><%=pSku & " "&pDescription%>
    <%if pPersonalizationDesc<>"" then response.write "&nbsp;(" &pPersonalizationDesc& ")"%>
    </td>
    <td width="148"> 
    <%=getCartRowOptionals(pIdCartRow)%></td>
    <td width="76">            
     <%=pCurrencySign &  money(getCartRowPrice(pIdCartRow)) %> 
    </td>
  </tr>
  <%rstemp.movenext
  loop%> 
  <tr> 
    <td colspan="4"> 
      <hr>
    </td>
  </tr>
  
 <%if pCalculateTaxIncluded="-1" then%>
 <%
 pTaxIncludedPercentage=getTaxIncludedPercentage()
 if pTaxIncludedPercentage<0 then pTaxIncludedPercentage=pTaxIncludedPercentage*-1
 pTaxIncludedFee=(pTaxIncludedPercentage*pSubTotal)/(100+pTaxIncludedPercentage)
 %>
  <tr> 
    <td colspan="3"><b><%=getMsg(473,"tax")%> Included</b></td>
    <td width="76"> 
      <div align="left"> <%=pCurrencySign & money(pTaxIncludedFee)%></div>
    </td>
  </tr>
 <%end if%>
 
  <%if pDiscountCode<>"" then%>   
  <tr> 
    <td colspan="3"><b><%=getMsg(459,"disc")%></b> 
	<%=pDiscountDesc%></td>
    <td width="76"> 
      <div align="left">- <%=pCurrencySign & money(pDiscountAmount)%></div>
    </td>
  </tr>
<%end if%>  

  <tr> 
    <td colspan="3"><b><%=getMsg(457,"payment")%></b> <%=pPaymentDescription%></td>
    <td width="76"> 
      <div align="left"><%if pPaymentSurcharge>0 then response.write pCurrencySign & money(pPaymentSurcharge)%></div>
    </td>
  </tr>
  
  <%if pByPassShipping="0" then%>
  <tr> 
    <td colspan="3"><b><%=getMsg(458,"ship method")%></b> <%=pShipmentDescription%>
    <%if Cdbl(pArrayOneShipment(2))>0 then%>
   + <%=getMsg(441,"handling")%>
    <%end if%>
    </td>
    <td width="76"> 
      <div align="left"><%=pCurrencySign & money(pShipmentCharge)%></div>
    </td>
  </tr>
 <%end if%>
 
  <tr> 
    <td colspan="2">&nbsp;</td>
    <td width="148"> 
      <div align="right"><b><%=getMsg(465,"subtotal")%></b></div>
    </td>
    <td width="76"> 
      <div align="left"><%=pCurrencySign & money(pSubTotal-pDiscountAmount+pPaymentSurcharge+pShipmentCharge) %>
        &nbsp;</div>
    </td>
  </tr>
 
   <tr> 
    <td colspan="2">&nbsp;</td>
    <td width="148"> 
      <div align="right"><b><%=getMsg(473,"tax")%></b></div>
    </td>         
    <td width="76"> 
    <%if pCalculateTaxIncluded<>"-1" then%>    
      <div align="left"><%=pCurrencySign & money(pTaxAmount)%>&nbsp;</div>
    <%end if%>
    </td>
  </tr>
 
  <tr> 
    <td colspan="2">&nbsp;</td>
    <td width="148"> 
      <div align="right"><b><%=getMsg(466,"total")%></b></div>
    </td>
    <td width="76"> 
      <div align="left"><%=pCurrencySign & money(pTotal) %> 
        &nbsp;</div>
    </td>
  </tr>
  
  <%if isOffLinePayment(pIdPayment) and pToken="0" then%>
  <tr>
    <td colspan="4">
    <table>
    <tr>
    <td><%=getMsg(512,"type")%></td>
    <td>
    
       <select name="cardType">
        <%if instr(pOffLineCardsAccepted,"VI")<>0 then%>
          <option value="V">Visa 
        <%end if%>  
        <%if instr(pOffLineCardsAccepted,"MA")<>0 then%>
          <option value="M">MasterCard      
        <%end if%>                 
        <%if instr(pOffLineCardsAccepted,"AM")<>0 then%>
          <option value="A">American Express           
        <%end if%>                 
        <%if instr(pOffLineCardsAccepted,"DIN")<>0 then%>  
          <option value="D">Diners Club
        <%end if%>                 
        <%if instr(pOffLineCardsAccepted,"DIS")<>0 then%>
          <option value="Discover">Discover
        <%end if%>                 
        <%if instr(pOffLineCardsAccepted,"JC")<>0 then%>
          <option value="JCB">JCB
        <%end if%>  
       </select>
         
    </td>
    </tr>
    <tr>
     <td><%=getMsg(504,"number")%></td>
     <%if pStoreFrontDemoMode="-1" then%> 
      <td><input type=text name=cardNumber size=12 value="4300000000000"></td>
     <%else%>
      <td><input type=text name=cardNumber size=12></td>
     <%end if%>
    </tr>
    <tr><td><%=getMsg(505,"expi")%></td><td>
    
    <%=getMsg(506,"month")%> 
        <select name="expMonth">
          <option value="1">1 
          <option value="2">2 
          <option value="3">3 
          <option value="4">4 
          <option value="5">5 
          <option value="6">6 
          <option value="7">7 
          <option value="8">8 
          <option value="9">9 
          <option value="10">10 
          <option value="11">11 
          <option value="12" selected>12 
        </select>
        <%=getMsg(507,"year")%> 
        <select name="ExpYear">                                                  
          <option value="2008">2008           
          <option value="2009">2009           
          <option value="2010">2010         
          <option value="2011">2011         
          <option value="2012">2012           
          <option value="2013">2013           
          <option value="2014">2014           
          <option value="2015">2015           
          <option value="2016">2016           
          <option value="2017">2017          
          <option value="2018">2018 
          <option value="2019">2019          
          <option value="2020">2020          
          <option value="2021" selected>2021                  
        </select>
    
    </td></tr>
    <tr>
     <td><%=getMsg(508,"CVV2")%></td>
     
     <%if pStoreFrontDemoMode="-1" then%> 
      <td><input type=text name=cvv2 size=4 value="123"></td>
     <%else%>
      <td><input type=text name=cvv2 size=4></td>
     <%end if%>
     
    </tr>
    </table>
    </td>
  </tr>  
  
  <%end if%>
  
  <tr>
    <td colspan="4">&nbsp;</td>
  </tr>  
  
   
<%if isBonusPayment(pIdPayment) and pBonusPointsAvailable >= pTotal then%> 
    <tr>
    	<td colspan="2" align="center">
    	    	<input type="submit" name="Bonus" value="<%=getMsg(757,"order and pay with bonus")%>" >
	</td>
	<td colspan="2"> 
   	   <div align="center"> 
        	<input type="submit" name="Cancel" value="<%=getMsg(468,"cancel")%>">
     	 </div>      
    </td>
   
    </tr>
<%else%>    
  <tr> 
    <td colspan="2"> 
      <div align="center"> 
        <input type="submit" name="Save" value=">> <%=getMsg(467,"confirm")%> <<">
      </div>
    </td>
    <td colspan="2"> 
      <div align="center"> 
        <input type="submit" name="Cancel" value="<%=getMsg(468,"cancel")%>">
      </div>      
    </td>
   <tr>    
<%end if%>

    <td colspan="4"><br><i>* <%=getMsg(469,"you have")%> <a href="comersus_conditions.asp" target="_blank"><%=getMsg(470,"terms")%></a></i></td>
    </tr>
  </tr>
</table>
</form>
<br><br>
<!--#include file="footer.asp"-->
<%call closeDb()%>