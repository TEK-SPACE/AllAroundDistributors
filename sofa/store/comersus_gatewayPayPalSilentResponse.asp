<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: Instant Payment Notification (IPN) for PayPal
%>

<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/stringFunctions.asp"--> 
<!--#include file="../includes/screenMessages.asp"--> 
<!--#include file="../includes/sendMail.asp"--> 
<!--#include file="../includes/miscFunctions.asp"--> 
<!--#include file="../includes/currencyFormat.asp"-->
<!--#include file="../includes/orderFunctions.asp"--> 
<!--#include file="../includes/itemFunctions.asp"-->

<%
on error resume next

dim pAllData, pIdOrder, pTxn_id, pPayment_status, objXMLHTTP, xml, pPayPalMerchantEmail, pPayPalOrderTotal, mySQL, conntemp, rstemp, rstemp2, rstemp3

' get settings 

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
pOneStepCheckout	= getSettingKey("pOneStepCheckout")
pEncryptionPassword 	= getSettingKey("pEncryptionPassword")
pUseEncryptedTotal	= getSettingKey("pUseEncryptedTotal")
pStoreLocation		= getSettingKey("pStoreLocation")
pEmailSender		= getSettingKey("pEmailSender")
pEmailAdmin		= getSettingKey("pEmailAdmin")
pSmtpServer		= getSettingKey("pSmtpServer")
pEmailComponent		= getSettingKey("pEmailComponent")
pDebugEmail		= getSettingKey("pDebugEmail")
pOrderPrefix		= getSettingKey("pOrderPrefix")
pSerialCodeOnlyOnce	= getSettingKey("pSerialCodeOnlyOnce")
pPayPalMerchantEmail	= getSettingKey("pPayPalMerchantEmail")

pAllData	 	 = request.form
pPayment_status  	 = getUserInput(request.form("payment_status"),50)
pReceiver_email 	 = getUserInput(request.form("receiver_email"),50)
pPayPalOrderTotal	 = getUserInput(request.form("payment_gross"),10)
pIdOrder 	 	 = getUserInput(request.form("invoice"),10)

' process only if Receiver Email is equal to Merchant Email
if pReceiver_email<>pPayPalMerchantEmail then
  response.redirect "comersus_supportError.asp?error="&Server.UrlEncode("Receiver email is different from Merchant email. Defined is: "&pPayPalMerchantEmail& ", the other is:" &pReceiver_email) 
end if 

' extract real idorder (without prefix)
pRealIdOrder 		= removePrefix(pIdOrder,pOrderPrefix)

if Not isNumeric(pRealIdOrder) then
 response.redirect "comersus_supportError.asp?error="&Server.UrlEncode("A numeric value expected for id Order") 
end if

mySQL="SELECT orderStatus, total FROM orders WHERE idOrder=" &pRealIdOrder

call updateDatabase(mySQL, rstemp, "gatewayPayPalResponse")
 
' if the order was not pending, then exit
if rstemp("orderStatus")<>1 then
   response.redirect "comersus_supportError.asp?error="&Server.UrlEncode("PayPal IPN, Order is not pending") 
end if

pOrderTotal=rstemp("total")
 
' post back to PayPal system to validate
pAllData	= pAllData & "&cmd=_notify-validate"

' create an xmlhttp object
response.buffer = true
set xml 	= Server.CreateObject("Microsoft.XMLHTTP")

' connect to PayPal and send all data                   
xml.open "POST", "https://www.paypal.com/cgi-bin/webscr", false	 
xml.send pAllData

' retrieve PayPal response
strRetval=xml.responseText

' check notification validation
if (xml.status <> 200) then
 
 ' HTTP error handling
 call sendmail (pCompany, pEmailSender, pEmailAdmin, "#"&pOrderPrefix&pIdorder, "PayPal IPN response: HTTP error - STR: "&strRetval)

elseif (strRetval = "VERIFIED") then

 ' process payment (compare order total and idOrder and payment_status="Completed") - invoice, payment_gross  

 if pPayPalOrderTotal=money(pOrderTotal) then
  pIdOrder 	 	 = getUserInput(request.form("invoice"),10)  
  call paymentApproved("PayPal", pIdOrder, pAuthCode)	   
 else
   ' different order total
   call sendmail (pCompany, pEmailSender, pEmailAdmin, "#"&pOrderPrefix&pIdorder, "PayPal IPN warning. Different order total $"&pPayPalOrderTotal&" - real order total: $"&pOrderTotal&" - STR: "&strRetval)     
 end if
   
elseif (strRetval = "INVALID") then
 
 ' possible fraud
 call sendmail (pCompany, pEmailSender, pEmailAdmin, "#"&pOrderPrefix&pIdorder, "PayPal IPN response: maybe fraud. "&strRetval)
 
else 
 ' error
 call sendmail (pCompany, pEmailSender, pEmailAdmin, "#"&pOrderPrefix&pIdorder, "PayPal IPN response: maybe error. "&strRetval)
end if

set xml = nothing

call closeDb()

%>


