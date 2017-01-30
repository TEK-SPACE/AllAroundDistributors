<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: PayPal Gateway secure form

%>
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/screenMessages.asp"-->
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/encryption.asp"--> 
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/adSenseFunctions.asp"-->
<!--#include file="comersus_optDes.asp"-->    
<%

on error resume next

dim connTemp, rsTemp

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanySlogan	 	= getSettingKey("pCompanySlogan")

pMoneyDontRound	 	= getSettingKey("pMoneyDontRound")
pAuctions		= getSettingKey("pAuctions")
pListBestSellers 	= getSettingKey("pListBestSellers")
pNewsLetter 		= getSettingKey("pNewsLetter")
pPriceList 		= getSettingKey("pPriceList")
pStoreNews 		= getSettingKey("pStoreNews")
pEncryptionPassword 	= getSettingKey("pEncryptionPassword")
pUseEncryptedTotal	= getSettingKey("pUseEncryptedTotal")
pStoreLocation		= getSettingKey("pStoreLocation")
pEncryptionMethod	= getSettingKey("pEncryptionMethod")
pShowNews		= getSettingKey("pShowNews")
pPayPalMerchantEmail	= getSettingKey("pPayPalMerchantEmail")

pIdCustomer     	= getSessionVariable("idCustomer",0)
pIdCustomerType 	= getSessionVariable("idCustomerType",1)

pAmount			= 0

if pUseEncryptedTotal="-1" then
    pAmount=money(DeCrypt(getUserInput(request.querystring("ordertotal"),30),pEncryptionPassword))
else
    pAmount=money(getUserInput(request.querystring("orderTotal"),30))
end if

pAmount			= formatNumberForDb(pAmount)  

%>

<!--#include file="header.asp"--> 
<br><b><%=getMsg(676,"Paypal payment")%></b><br> 
 
<%if pStoreFrontDemoMode = "-1" then %>
 <form action="https://www.sandbox.paypal.com/cgi-bin/webscr" method="post">
<%else%>
 <form action="https://www.paypal.com/cgi-bin/webscr" method="post">
<%end if%>

<%if pCurrencySign="$" then%>
 <input type="hidden" name="currency_code" 	value="USD">
<%else%>
 <input type="hidden" name="currency_code" 	value="GBP">
<%end if%>

 <input type="hidden" name="idOrder" 		value="<%=getUserInput(request.querystring("idOrder"),10)%>">
 <input type="hidden" name="invoice" 		value="<%=getUserInput(request.querystring("idOrder"),10)%>">
 <input type="hidden" name="cmd" 		value="_xclick">
 <input type="hidden" name="bn" 		value="comersus">
 <input type="hidden" name="business" 		value="<%=pPayPalMerchantEmail%>"> 
 <input type="hidden" name="item_name" 		value="<%=pCompany & " items"%>">
 <input type="hidden" name="item_number" 	value="<%=getUserInput(request.querystring("orderDetails"),200)%>">
 <input type="hidden" name="amount" 		value="<%=pAmount%>">
 <input type="hidden" name="return" 		value="http://<%=pStoreLocation%>/store/comersus_gatewayPayPalOrderConfirmation.asp">
 <br><%=getMsg(677,"click 2 pay")%><br>
 
<br><br> <input type="image" src="http://images.paypal.com/images/x-click-but01.gif" name="submit" alt="Make payments with PayPal - it’s fast, free and secure!">

 
</form>
<br>
<!--#include file="footer.asp"-->
<%call closeDb()%>  