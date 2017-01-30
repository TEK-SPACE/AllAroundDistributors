<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com 
' Details: add donation item to the shopping cart
%>

<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->
<!--#include file="../includes/screenMessages.asp"-->
<!--#include file="../includes/cartFunctions.asp"--> 
<!--#include file="../includes/stringFunctions.asp"-->

<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/miscFunctions.asp"--> 

<% 

on error resume next

dim mySQL, connTemp, rsTemp, rsTemp2, pTotalQuantity, lineNumber, pProductHasOptionals, pIdCartRow, pRowPrice, f, pOptionDescrip, pPriceToAdd, pDefaultLanguage, pStoreFrontDemoMode, pCurrencySign, pDecimalSign, pMoneyDontRound, pCompany, pCompanyLogo, pHeaderKeywords, pAuctions, pListBestSellers, pNewsLetter, pPriceList, pStoreNews, pOneStepCheckout, pForceSelectOptionals, pMaxAddCartQuantity, pCartQuantityLimit, pUnderStockBehavior, pDateFormat, pCookiesDetection, pChangeDecimalPoint, pIdDbSession, pIdDbSessionCart, pIdCustomer, pIdCustomerType, pIdProduct, p1StepCheckout, pOptionsQuantity, pQuantity, pPersonalizationDesc, pFrom, pUntil, pIdProductFree, pStock, pHowManyOptionals, pHowManyOptionalsSelected, pThereAreSelectedOptionals, pUnitPrice, pUnitBToBPrice, pUnitWeight, pUnitCost, pSpecialPrice, pCartRowPrice, pNewQuantity, pOldQuantity

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pForceSelectOptionals	= getSettingKey("pForceSelectOptionals")
pMaxAddCartQuantity	= getSettingKey("pMaxAddCartQuantity")
pCartQuantityLimit	= getSettingKey("pCartQuantityLimit")
pUnderStockBehavior	= getSettingKey("pUnderStockBehavior")
pDateFormat 		= getSettingKey("pDateFormat")
pCookiesDetection	= getSettingKey("pCookiesDetection")
pChangeDecimalPoint	= getSettingKey("pChangeDecimalPoint")
pTemplateStore		= getSettingKey("pTemplateStore")

if pCookiesDetection="-1" and readCookie()=0 then 
 response.redirect "comersus_cookiesInformation.asp"
end if

pIdDbSession	 		= checkSessionData()
pIdDbSessionCart 		= checkDbSessionCartOpen()

pIdCustomer      		= getSessionVariable("idCustomer",0)
pIdCustomerType  		= getSessionVariable("idCustomerType",1)

if pIdDbSessionCart=0 then
 pIdDbSessionCart=createNewDbSessionCart() 
end if

' get data from viewitem form
pIdProduct	 	= getUserInput(request("idProduct"),10)
pUnitPrice 		= getUserInput(request("price"),10)
pQuantity	 	= 1

' if is not a donation, redirect to regular viewItem
if not isDonation(pIdProduct) then 
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(671,"You can only make donations with donation items."))
end if
     
if pChangeDecimalPoint="-1" then
 pUnitPrice	= formatNumberForDb(pUnitPrice)
end if

' if is not a number
if not isNumeric(pUnitPrice) then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(710,"Invalid number"))
end if
   
' insert new row line, is not in the cart   

mySQL="INSERT INTO cartRows (idProduct, quantity, unitPrice, unitCost, unitWeight, idDbSessionCart, personalizationDesc) VALUES (" &pIdProduct& "," &pQuantity& "," &pUnitPrice& ",0,0," &pIdDbSessionCart& ",'')"
call updateDatabase(mySQL, rstemp, "addItem") 
 
call closeDB()

response.redirect "comersus_goToShowCart.asp"
%>


