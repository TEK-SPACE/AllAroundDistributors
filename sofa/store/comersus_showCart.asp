<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: show cart contents
%>

<!--#include file="../includes/screenMessages.asp"-->
<!--#include file="../includes/cartFunctions.asp"--> 
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->
<!--#include file="../includes/miscFunctions.asp"-->
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/discountFunctions.asp"--> 
<!--#include file="../includes/customerTracking.asp"--> 
<!--#include file="../includes/adSenseFunctions.asp"-->


<%

on error resume next

dim mySQL, connTemp, rsTemp, rsTemp2, rsTemp7, pDefaultLanguage, pStoreFrontDemoMode, pCurrencySign, pDecimalSign, pMoneyDontRound, pCompany, pCompanyLogo, pHeaderKeywords, pAuctions, pListBestSellers, pNewsLetter, pPriceList, pStoreNews, pOneStepCheckout, pRealTimeShipping, pAllowNewCustomer, pByPassShipping, total, totalDeliveringTime, pIdDbSession, pIdDbSessionCart, pIdCustomer, pIdCustomerType, pLanguage, pCustomerName, pHeaderCartItems, pAffiliatesStoreFront, pIdCartRow, pIdProduct, pQuantity, pUnitPrice,  pDescription, pSku, pDeliveringTime, pPersonalizationDesc, pRental, pItHasFreeProduct, pDetails, pOptionGroupsTotal, pRowPrice, pBonusPoints

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanySlogan 		= getSettingKey("pCompanySlogan")
pAboutUsLink 		= getSettingKey("pAboutUsLink")
pHeaderKeywords 	= getSettingKey("pHeaderKeywords")
pPayPalExpressCheckout	= getSettingKey("pPayPalExpressCheckout")
pShowSearchBox 		= getSettingKey("pShowSearchBox")
pAdSenseClient 		= getSettingKey("pAdSenseClient")
pMoneyDontRound	 	= getSettingKey("pMoneyDontRound")
pByPassShipping		= getSettingKey("pByPassShipping")
pTaxIncluded		= getSettingKey("pTaxIncluded")
pRealTimeShipping	= getSettingKey("pRealTimeShipping")
pAllowNewCustomer	= getSettingKey("pAllowNewCustomer")
pForgotPassword 	= getSettingKey("pForgotPassword")

pAffiliatesStoreFront   = getSettingKey("pAffiliatesStoreFront")
pSuppliersList		= getSettingKey("pSuppliersList")
pAuctions		= getSettingKey("pAuctions")
pNewsLetter 		= getSettingKey("pNewsLetter")
pShowNews		= getSettingKey("pShowNews")

total			= Cdbl(0)
totalDeliveringTime	= Cdbl(0)

pIdDbSession	 	= checkSessionData()
pIdDbSessionCart 	= checkDbSessionCartOpen()

pIdCustomer	 	= getSessionVariable("idCustomer",0)
pIdCustomerType  	= getSessionVariable("idCustomerType",1)
pDiscountCode 		= getSessionVariable("discountCode","")

' check if the cart is empty
if countCartRows(pIdDbSessionCart)=0 then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(40,"emtpy"))       
end if

call customerTracking("comersus_showCart.asp", request.querystring)

pSubTotal		= calculateCartTotal(pIdDbSessionCart)

%>
<!--#include file="header.asp"--> 
<SCRIPT LANGUAGE="JavaScript">
function validateNumber(field)
{
  var val = field.value;
  if(!/^\d*$/.test(val)||val==0)
  {
      alert("<%=getMsg(41,"values >0")%>");
      field.focus();
      field.select();
  }
}
</script>

<br><%=getMsg(42,"Your cart contains")%>
  <form method="post" action="comersus_cartRecalculate.asp" name="recalculate">
    <table border="0">
      <tr> 
        <td bgcolor="#CCCCCC"><%=getMsg(43,"Qty")%></td>        
        <td bgcolor="#CCCCCC"><%=getMsg(44,"Item")%></td>
        <td bgcolor="#CCCCCC">&nbsp;</td>
        <td bgcolor="#CCCCCC"><%=getMsg(45,"Opt")%></td>
        <td bgcolor="#CCCCCC"><%=getMsg(46,"Price")%></td>
        <td bgcolor="#CCCCCC"><%=getMsg(47,"Act")%></td>
      </tr>
      <%    
' get all products in the cart
mySQL="SELECT idCartRow, cartRows.idProduct, quantity, unitPrice, description, sku, deliveringTime, personalizationDesc, rental, smallImageUrl FROM cartRows, products WHERE cartRows.idProduct=products.idProduct AND cartRows.idDbSessionCart="&pIdDbSessionCart

call getFromDatabase(mySQL, rstemp, "showCart") 

do while not rstemp.eof
 pIdCartRow		= rstemp("idCartRow")
 pIdProduct		= rstemp("idProduct")
 pQuantity		= rstemp("quantity")
 pUnitPrice		= Cdbl(rstemp("unitPrice")) 
 pDescription		= rstemp("description")
 pSku			= rstemp("sku")
 pDeliveringTime	= rstemp("deliveringTime")
 pPersonalizationDesc	= rstemp("personalizationDesc")
 pRental		= rstemp("rental")
 pSmallImageUrl		= rstemp("smallImageUrl")
   
 %> 
      <tr> 
        <td><input type="text" name="<%=pIdCartRow%>" size="2" value="<%=pQuantity%>" onBlur="validateNumber(this)"></td>                
        <td><a href="comersus_viewItem.asp?idProduct=<%=pIdProduct%>"><%=pSku%> - <%=pDescription%>
        <%if pPersonalizationDesc<>"" then response.write "&nbsp;(" &pPersonalizationDesc& ")"%></a>
        </td>
        <td>
        <%if pSmallImageUrl<>"" then%>
         <img src="catalog/<%=pSmallImageUrl%>">
        <%else%>
         <img src="catalog/imageNa_sm.gif">
        <%end if%>
        </td>
        <td><%=getCartRowOptionals(pIdCartRow)%></td>        
        <td><%=pCurrencySign &  money(getCartRowPrice(pIdCartRow))%></td>
        <td>  
          <a href="comersus_cartRemove.asp?idCartRow=<%=pIdCartRow%>&idDbSessionCart=<%=pIdDbSessionCart%>"><img src="images/buttonRemoveSmall.gif" border=0 alt="<%=getMsg(48,"Remove")%>">
          </a></td>
      </tr>            
  <%        
  if Cint(pDeliveringTime)>totalDeliveringTime then
     totalDeliveringTime =  Cint(pDeliveringTime)
  end if
  
rstemp.moveNext
loop
%> 

</table>
    
    <br><input type="image" name="Recalculate" src="images/buttonRecalculate.gif" alt="<%=getMsg(49,"Rec")%>">
    </form>

<%if pDiscountCode<>"" then%>
 <%
 pDiscountDesc	=""
 pDiscountTotal	=0
 
 call getDiscount(pDiscountCode, pDiscountDesc, pDiscountTotal) 
 pSubTotal=pSubTotal-pDiscountTotal
 
 %>
 <br><br><%=getMsg(443,"discount")%>: <%=pDiscountDesc%> - <%=pCurrencySign & money(pDiscountTotal)%> [<a href="comersus_resetDiscountCode.asp"><%=getMsg(48,"remove")%></a>]
<%else%>
 <%if pStoreFrontDemoMode="-1" then%>
  <br><form method=post action=comersus_applyDiscountCode.asp><%=getMsg(443,"discount")%> <input type=text name="discountCode" size=12 value="123456"> <input type=submit value="Apply"></form>  
 <%else%>
  <br><form method=post action=comersus_applyDiscountCode.asp><%=getMsg(443,"discount")%> <input type=text name="discountCode" size=12> <input type=submit value="Apply"></form>
 <%end if%>
<%end if%>

    <br>
    <b> <%=getMsg(50,"Total")%> <%=pCurrencySign & money(pSubTotal) %></b>     
    <%if pTaxIncluded<>"-1" then%>
     <i> <%=getMsg(51,"Tx inc")%></i> <br>    
    <%else%>
     <i> <%=getMsg(52,"Tax n/incl")%></i> <br>    
    <%end if%>
    
    <%=getMsg(53,"Availability")%>&nbsp;<%
    if totalDeliveringTime>0 and totalDeliveringTime<999 then 
      response.write totalDeliveringTime & " " & getMsg(54," days")
    end if
    if totalDeliveringTime=0 then
      response.write getMsg(56,"same day")
    end if    
    ' use 999 for undetermined
    if totalDeliveringTime>998 then
      response.write getMsg(55,"Und")
    end if
    %>
    <br>
    <br>      
   
   <a href="comersus_index.asp"><%=getMsg(57,"Keep sh")%></a>
   
   <%if ucase(pRealTimeShipping)="NONE" and pByPassShipping="0" then%>
    <br><a href="comersus_shippingCostForm.asp"><%=getMsg(59,"Get sh cost")%></a>
   <%end if%>

<%if pIdCustomer<>0 then%>
 
    <form method=post action="comersus_checkOut2.asp">
     <input type=submit name=checkout value=">> <%=getMsg(58,"Checkout")%>">
    </form>
   
   <%if pPayPalExpressCheckout="-1" then%>	
	 <br><a href="comersus_gatewayPayPalExpress1.asp">
	 <img src="https://www.paypal.com/en_US/i/btn/btn_xpressCheckoutsm.gif" align="left" style="margin-right:7px;" border=0>
	 </a><br><br>
   <%end if%>  
	
<%else%>   

 <br><br><b><%=getMsg(58,"Checkout")%></b>
 <br><br>

 <%if pAllowNewCustomer="-1" and (pTeleSignCustomerId="0" or pTeleSignCustomerId="") then%>
  <br><%=getMsg(395,"f time")%>&nbsp;<a href="comersus_checkOutCustomerForm.asp"><%=getMsg(396,"register")%></a>
 <%end if%>

 <br><br><%=getMsg(397,"returning customer")%>
 <form method="post" name="auth" action="comersus_customerAuthenticateExec.asp">
    <input type="hidden" name="redirect" value="checkout">
        <table border="0" align="left">
          <tr> 
            <td><%=getMsg(398,"email")%> </td>
            <td> <%if pStoreFrontDemoMode="-1" then%> 
              <input type="text" name="email" value="test@comersus.com">
          <%else%>
             <input type="text" name="email" value="">
          <%end if%>
        </td>       
            <td rowspan="3"> 
                <%if pChargebackProtectionMerchant<>"0" and pChargebackProtectionMerchant<>"" then%>
                    <a target="_blank" href="http://www.chargebackprotection.org/frontend/sealVerify.asp?idMerchant=<%=pChargebackProtectionMerchant%>&rDate=<%=pChargebackProtectionRegDate%>"><img src="http://www.chargebackprotection.org/frontend/images/chargebackProtectionSeal.gif" alt="Click here to verify this seal" border=0 width="100" height="100"></a> 
                <%end if%>                    
                <%if pSecureStoreGraph="-1" then%> 
                    <a href="http://www.comersus.com/secure.html" target="_blank"><img src="images/secureStore.gif" width="61" height="111" alt="Comersus Secure Store" border="0"></a> 
                <%end if%> 
            </td>
      </tr>
      
      <tr> 
            <td><%=getMsg(399,"pass")%></td>
            <td> <%if pStoreFrontDemoMode="-1" then%> 
              <input type="password" name="password" value="123456">            
           <%else%>
            <input type="password" name="password" value="">
          <%end if%>          
          <%if pForgotPassword="-1" then%>
            <br><a href="comersus_optForgotPasswordForm.asp"><i><%=getMsg(400,"forgot")%></i></a>
          <%end if%>                    
        </td>         
      </tr>
      <tr> 
            <td><input type="submit" name="Submit" value="<%=getMsg(58,"Checkout")%>"></td>
            <td>&nbsp;</td>        
      </tr>      

	<%if pPayPalExpressCheckout="-1" then%>
	 <tr>
	 <td colspan=2><br><br><a href="comersus_gatewayPayPalExpress1.asp"><img src="https://www.paypal.com/en_US/i/btn/btn_xpressCheckoutsm.gif" align="left" style="margin-right:7px;" border=0></a>	 
	 </td> 	
	 </tr>
	<%end if%>  
	      
      <tr> 
            <td colspan=2>&nbsp;</td>        
      </tr>
    </table>
</form>  
<%end if%>


<br><br> 
<!--#include file="footer.asp"-->
<%call closeDb()%>
