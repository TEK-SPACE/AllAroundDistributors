<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
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
<!--#include file="../includes/discountFunctions.asp"--> 


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
pUseVatNumber		= getSettingKey("pUseVatNumber")

pOrderFieldName1	= getSettingKey("orderFieldName1")
pOrderFieldName2	= getSettingKey("orderFieldName2")
pOrderFieldName3	= getSettingKey("orderFieldName3")

pIdCustomer     	= getSessionVariable("idCustomer",0)
pIdCustomerType 	= getSessionVariable("idCustomerType",1)

pIdDbSession	 	= checkSessionData()
pIdDbSessionCart 	= checkDbSessionCartOpen()

' get user inputs
pDiscountCode  		= getUserInput(request.form("discountCode"),10)

' calculate total price of the order, total weight and product total quantities
pSubTotal  		= Cdbl(calculateCartTotal(pIdDbSessionCart))
pCartTotalWeight 	= Cdbl(calculateCartWeight(pIdDbSessionCart))
pCartQuantity  		= Cdbl(calculateCartQuantity(pIdDbSessionCart))

' bounce invalid discount codes
if pDiscountCode="" then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(570,"Please enter"))  
end if

pDiscountDesc	=""
pDiscountTotal	=0

call getDiscount(pDiscountCode, pDiscountDesc, pDiscountTotal) 

if pDiscountTotal=0 then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(pDiscountDesc)  
end if

session("discountCode")=pDiscountCode

call closeDb()

response.redirect "comersus_showCart.asp"

%>
