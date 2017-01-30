<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: confirmation for PayPal payments
%>

<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/screenMessages.asp"--> 
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/adSenseFunctions.asp"-->

<%

on error resume next

dim connTemp, rsTemp

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanyLogo	 	= getSettingKey("pCompanyLogo")
pEmailAdmin	 	= getSettingKey("pEmailAdmin")

pAuctions		= getSettingKey("pAuctions")
pListBestSellers 	= getSettingKey("pListBestSellers")
pNewsLetter 		= getSettingKey("pNewsLetter")
pPriceList 		= getSettingKey("pPriceList")
pStoreNews 		= getSettingKey("pStoreNews")
pOneStepCheckout	= getSettingKey("pOneStepCheckout")
pEncryptionPassword 	= getSettingKey("pEncryptionPassword")
pUseEncryptedTotal	= getSettingKey("pUseEncryptedTotal")
pStoreLocation		= getSettingKey("pStoreLocation")
%>

<!--#include file="header.asp"-->
<br><br><b><%=getMsg(676,"Paypal payment")%></b><br> 
<br><br><%=getMsg(678,"Thanks")%>
<br>Payment transaction details will be emailed soon.
<br><br>   
<%call closeDb()%>                   
<!--#include file="footer.asp"-->

