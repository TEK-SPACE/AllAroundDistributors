<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
%>

<% 
' if it's a cancel request, go to cancel script
if request.form("Cancel")<>"" then
 response.redirect "comersus_checkOutCancelOrder.asp"
end if

%> 
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/screenMessages.asp"--> 
<!--#include file="../includes/encryption.asp"--> 
<!--#include file="../includes/cartFunctions.asp"--> 
<!--#include file="../includes/dateFunctions.asp"--> 
<!--#include file="../includes/stringFunctions.asp"--> 
<!--#include file="../includes/miscFunctions.asp"--> 
<!--#include file="../includes/discountFunctions.asp"--> 
<!--#include file="../includes/creditCardFunctions.asp"--> 
<!--#include file="../includes/paymentMethodFunctions.asp"--> 
<!--#include file="../includes/sendMail.asp"--> 
<!--#include file="../includes/chargebackprotection.asp"--> 
<!--#include file="../includes/fraudPrevention.asp"--> 
<!--#include file="../includes/customerTracking.asp"--> 
<!--#include file="../includes/preCharge.asp"--> 
<!--#include file="../includes/maxMind.asp"--> 
<!--#include file="../includes/orderFunctions.asp"--> 
<!--#include file="../includes/itemFunctions.asp"-->
<!--#include file="comersus_optDes.asp"-->    

<% 

on error resume next

dim mySQL, connTemp, rsTemp, pTotal, pTaxAmount, rsTemp3

' get settings 
pDefaultLanguage 		= getSettingKey("pDefaultLanguage")
pStoreFrontDemoMode 		= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 		= getSettingKey("pCurrencySign")
pDecimalSign	 		= getSettingKey("pDecimalSign")
pMoneyDontRound	 		= getSettingKey("pMoneyDontRound")
pCompany	 		= getSettingKey("pCompany")

pEmailSender			= getSettingKey("pEmailSender")
pEmailAdmin			= getSettingKey("pEmailAdmin")
pSmtpServer			= getSettingKey("pSmtpServer")
pEmailComponent			= getSettingKey("pEmailComponent")
pDebugEmail			= getSettingKey("pDebugEmail")
pEncryptionPassword		= getSettingKey("pEncryptionPassword")
pFraudPreventionMode		= getSettingKey("pFraudPreventionMode")
pOrderPrefix			= getSettingKey("pOrderPrefix")
pUseEncryptedTotal		= getSettingKey("pUseEncryptedTotal")
pDisableSaveOrderEmail		= getSettingKey("pDisableSaveOrderEmail")
pDateFormat			= getSettingKey("pDateFormat")
pEncryptionMethod		= getSettingKey("pEncryptionMethod")
pBonusPointsPerPrice		= getSettingKey("pBonusPointsPerPrice")

pChargebackProtectionMerchant   = getSettingKey("pChargebackProtectionMerchant")
pChargebackProtectionPassword   = getSettingKey("pChargebackProtectionPassword")
pChargebackProtectionRegDate    = getSettingKey("pChargebackProtectionRegDate")

pPreChargeMerchant		= getSettingKey("pPreChargeMerchant")
pMaxMindLicenseKey 		= getSettingKey("pMaxMindLicenseKey")
pMaxMindScoreApproved		= getSettingKey("pMaxMindScoreApproved")
pMaxMindScoreAlert  		= getSettingKey("pMaxMindScoreAlert")


pCardType			= getUserInput(request.form("cardType"),8)
pCardNumber			= getUserInput(request.form("cardNumber"),20)
pExpirationMonth		= getUserInput(request.form("expMonth"),2)
pExpirationYear			= getUserInput(request.form("expYear"),4)
pExpiration			= pExpirationMonth & "/" & pExpirationYear 
pSeqCode			= getUserInput(request.form("cvv2"),4)

pBonusPointButton		= getUserInput(request.form("Bonus"),10)

if pBonusPointsPerPrice="" or not isNumeric(pBonusPointsPerPrice) then  
  pBonusPointsPerPrice=0
end if
 
' IP request to server variables
pBrowserIp			= request.ServerVariables("REMOTE_HOST")

' get idDbSession in order to retrieve fields from the session
pIdDbSession			= getSessionVariable("idDbSession", 0)
pWishListIdCustomer		= getSessionVariable("wishListIdCustomer", 0)

if sessionLost() then 
 response.redirect "comersus_message.asp?message="&Server.UrlEncode(getMsg(477,"It seems that your session was lost due to inactivity. Please try again from store home. Sorry for the inconvenience."))
end if

pIdDbSessionCart 		= checkDbSessionCartOpen()

' check if the cart is empty 
if countCartRows(pIdDbSessionCart)=0 then 
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(478,"Your cart is empty, you cannot save the order."))
end if

call customerTracking("comersus_checkoutSaveOrder.asp", request.querystring)

' get session variables
pIdCustomer     	= getSessionVariable("idCustomer",0)
pIdCustomerType     	= getSessionVariable("idCustomerType",1)
pIdAffiliate     	= getSessionVariable("idAffiliate",1)

' get all fields from dbSession 

mySQL="SELECT sessionData FROM dbSession WHERE idDbSession=" &pIdDbSession
call getFromDatabase(mySQL, rstemp, "orderVerify")
 
if rstemp.eof then
 response.redirect "comersus_supportError.asp?error="&Server.Urlencode("Cannot get dbSession data at saveOrder for "&pIdDbSession)
end if
 
pSessionData		=rstemp("sessionData") 
 
if len(pSessionData)=0 or pSessionData="*" then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(682,"It seems that you have reloaded or pressed twice the save order button. Please use the store navigation links."))
end if

pArraySessionData	=split(pSessionData,"||")
 
 ' reloaded verification? (a reload may cause sessionData to have different number of values)
if uBound(pArraySessionData)<>12 then 
  response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(682,"It seems that you have reloaded the saveorder script. Please use the store navigation links."))
end if
 
pShipmentDetails 	= pArraySessionData(0)
pPaymentDetails		= pArraySessionData(1)
pComments	 	= pArraySessionData(2)
pBonusPoints		= Cdbl(pArraySessionData(3))
pVatNumber 		= pArraySessionData(4)
pUser1 			= pArraySessionData(5)
pUser2 			= pArraySessionData(6)
pUser3 			= pArraySessionData(7)
pDiscountCode		= pArraySessionData(8)
pDiscountAmount		= pArraySessionData(9)
pTaxAmount		= pArraySessionData(10)
pTotal			= pArraySessionData(11)
pIdPayment		= pArraySessionData(12)


' check bonus points selection
pBonusPointsAvailable		=0
pBonusPointsAvailable		=getBonusPoints(pIdCustomer)
pBonusPoints			=0

' compare customer bonus points with order total
if cdbl(pBonusPointsAvailable) >= cdbl(pTotal) Then
 if pBonusPointButton <> "" then
 	' set bonus points used in this order
	pBonusPoints = cdbl(pTotal)
	pPaymentDetails = getMsg(158,"Bonus Points") 	
 end if
end if

if pTotal > 0 Then

 if isOffLinePayment(pIdPayment) then

 ' validates expiration
  if DateDiff("d", Month(Now)&"/"&Year(now), pExpirationMonth&"/"&pExpirationYear)<0 then
    response.redirect "comersus_message.asp?message="&Server.UrlEncode(getMsg(515,"Credit card expired")&". "&getMsg(615,"go back"))
  end if

  ' validates card
  if not ValidateCreditCard(pCardNumber, pCardType) then
	response.redirect "comersus_message.asp?message="&Server.UrlEncode(getMsg(516,"Invalid card number")&". "&getMsg(615,"go back"))
  end if
 
 end if
end if


' call fraud prevention

pFiltersReturned=rejectOrder()

if pFraudPreventionMode<>"none" and pFiltersReturned<>"" then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(716,"You cannot checkout ")&" " &pFiltersReturned)
end if

' send precharge request
if pPreChargeMerchant<>"0" and pTotal > 0 then
 if preChargeApproved(pCardNumber, pExpirationMonth, pExpirationYear, pResponse)=0 then 
  response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(513,"This transaction has been rejected. Please verify your information and try again. Response details: ") &pResponse)
 end if
end if

pPreChargeMsg=""
if pPreChargeMerchant="0" Then 
	pPreChargeMsg="PreCharge Not Enabled [Enable now at https://www.precharge.com/]"
end if
 
' send MaxMind request
if pMaxMindLicenseKey <> "0" then
 pScore = 10
 if MaxMindApproved(pIdCustomer, pBrowserIp, pScore) = 0 then
    response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(717,"This transaction has been rejected. Please verify your information and try again. Score : ") & pScore)
 end if
end if ' maxmind 

' get customer data

mySQL="SELECT * FROM customers WHERE idCustomer=" &pIdCustomer
call getFromDatabase(mySQL, rstemp, "saveOrder")

if not rstemp.eof then
 pName			= formatForDb(rstemp("name"))
 pLastName		= formatForDb(rstemp("lastName"))
 pCustomerCompany	= formatForDb(rstemp("customerCompany"))
 pEmail			= rstemp("email")
 pPassword		= rstemp("password")
 pPhone			= rstemp("phone") 

 pAddress		= formatForDb(rstemp("address"))
 pZip			= formatForDb(rstemp("zip"))
 pStateCode		= rstemp("stateCode")
 pState			= formatForDb(rstemp("state"))
 pCity			= formatForDb(rstemp("city"))
 pCountryCode		= rstemp("countryCode")

 pShippingName		= formatForDb(rstemp("shippingName"))
 pShippingLastName	= formatForDb(rstemp("shippingLastName"))
 pShippingAddress	= formatForDb(rstemp("shippingAddress"))
 pShippingZip		= formatForDb(rstemp("shippingZip"))
 pShippingStateCode	= rstemp("shippingStateCode")
 pShippingState		= formatForDb(rstemp("shippingState"))
 pShippingCity		= formatForDb(rstemp("shippingCity"))
 pShippingCountryCode   = rstemp("shippingCountryCode")
 
 if isNull(pCustomerCompany) then
  pCustomerCompany=""
 end if
 
end if

' call chargeback protection API

if pChargebackProtectionMerchant<>"0" and pChargebackProtectionMerchant<>"" then
    pChargeBackMSG = checkChargeProtection(pChargebackProtectionMerchant, pChargebackProtectionPassword, pName, pLastName, pCountryCode)
else    
    pChargeBackMSG = "Not enabled [Enable now http://www.chargebackprotection.org/register.html]" 
end if

' create date, fix 0 if day or month has one place
pOrderDate=fixDate(Date())

' compile order contents snapshot field (contents are still linked through cartRows)

mySQL="SELECT idCartRow, cartRows.idProduct, quantity, unitPrice, description, sku, deliveringTime, emailText, personalizationDesc FROM cartRows, products WHERE cartRows.idProduct=products.idProduct AND cartRows.idDbSessionCart="&pIdDbSessionCart    
  
call getFromDatabase(mySQL, rsTemp, "saveorder 4")    

  do while not rsTemp.eof
  
   pIdCartRow		= rsTemp("idCartRow")
   pIdProduct		= rsTemp("idProduct")
   pSku			= rsTemp("sku")      
   pQuantity		= rsTemp("quantity")
   pUnitPrice		= Cdbl(rsTemp("unitPrice"))   
   pDescription		= rsTemp("description")   
   pPersonalizationDesc	= rsTemp("personalizationDesc")      
   
   ' get optionals
   
    pOptionGroupsTotal=0
          
    ' get optionals for current cart row
  
    mySQL="SELECT optionDescrip, priceToAdd FROM cartRowsOptions WHERE idCartRow="&pIdCartRow	  	  	  	  
    
    call getFromDatabase(mySQL, rsTemp2, "saveorder 5")    
    
    pOptionsDescripCompound=""         
    
    do while not rsTemp2.eof
         
      pPriceToAdd			= Cdbl(rsTemp2("priceToAdd"))
      pOptionsDescripCompound		= pOptionsDescripCompound & rsTemp2("optionDescrip")&" "
  
      if pPriceToAdd>0 then
         pOptionsDescripCompound	= pOptionsDescripCompound &pCurrencySign & money(pPriceToAdd)
      end if	      
	  
      pOptionGroupsTotal		= pOptionGroupsTotal+ pPriceToAdd        
          
    rsTemp2.movenext
   loop                   
   
  pRowPrice = Cdbl(pQuantity * (pOptionGroupsTotal+pUnitPrice) )         
  
  ' compile order details `snapshot` (items are also saved into cartRows table)                
  pDetails	= pDetails & pQuantity& "x #" &pSku& "/" &pIdProduct & " " &pDescription 
  
  if pPersonalizationDesc<>"" then
    pDetails=pDetails&"(" &pPersonalizationDesc& ")"
  end if
  
  pDetails	= pDetails & " " & getMsg(479,"variations") &" " &pOptionsDescripCompound& "= " &pCurrencySign & money(pRowPrice) &Vbcrlf                
  rsTemp.movenext
loop

' compile discount and mark discount as used

if pDiscountCode<>"" then   

 pDiscountDetails = Cstr("")
 pDiscountDetails = pDiscountDetails& Vbcrlf & getMsg(481,"disc code") & ": " & pDiscountCode &" " & pCurrencySign&money(pDiscountAmount) 
 
 call markDiscountAsUsed(pDiscountCode)
 
end if ' discount code active
  
' get redirectionurl and email text from payment selected

if pBonusPointButton="" or pBonusPoints=0 then
 
 mySQL="SELECT redirectionUrl, emailText FROM payments WHERE idPayment=" &pIdPayment
 call getFromDatabase(mySQL, rsTemp, "saveorder 7")    

 if rsTemp.eof then 
  response.redirect "comersus_supportError.asp?error="&Server.Urlencode("Cannot locate selected payments in database.")
 end if

 ' load redirection and email text for payment method selected

 pRedirectionUrl	= rsTemp("redirectionUrl")
 pEmailText		= rsTemp("emailText")

end if

' if idAffiliate is not valid, set generic affiliate
if affiliateValid(pIdAffiliate)=0 then pIdAffiliate=1

' random number to locate inserted record without using specific DB engine methods 
pRandomNumber=randomNumber(99999)

' format fields for database
pTotalToSave		= replace(pTotal,",",".")
pTaxAmountToSave	= replace(pTaxAmount,",",".")
pDetails 		= formatForDb(pDetails)
pUser1 			= formatForDb(pUser1)
pUser2 			= formatForDb(pUser2)
pUser3 			= formatForDb(pUser3)

mySQL="INSERT INTO orders (orderDate, idCustomer, details, total, taxAmount, obs, address, zip, state, stateCode, city, countryCode, shippingName, shippingLastName, shippingAddress, shippingZip, shippingState, shippingStateCode, shippingCity, shippingCountryCode, shipmentDetails, paymentDetails, discountDetails, nroRan, orderStatus, idAffiliate, viewed, idCustomerType, digitalEmailText, browserIp, vatNumber, idStore, wishListIdCustomer, user1, user2, user3) VALUES ('" & pOrderDate  & "','" & Cstr(pIdCustomer)& "','" &pDetails &"'," &pTotalToSave& "," &pTaxAmountToSave& ",'" &pComments& "','" &pAddress & "','" &pZip& "','" &pState& "','" &pStateCode& "','" &pCity& "','" &pCountryCode& "','" &pShippingName & "','" &pShippingLastName & "','" &pShippingAddress & "','" &pShippingZip& "','" &pShippingState& "','" &pShippingStateCode& "','" &pShippingCity& "','" &pShippingCountryCode& "','" &pShipmentDetails& "','" &pPaymentDetails& "','" &pDiscountDetails& "'," &pRandomNumber& ",1," &pIdAffiliate& ",0," &pIdCustomerType& ",'','" &pBrowserIp& "','" &pVatNumber& "'," &pIdStore& "," &pWishListIdCustomer& ",'"& pUser1 &"','"& pUser2 &"','"& pUser3 &"')"

call updateDatabase(mySQL, rsTemp, "saveorder 9")    

' get id of the saved order 
mySQL1="SELECT idOrder FROM orders WHERE nroRan=" &pRandomNumber& " AND idCustomer=" &pIdCustomer& " AND orderDate='"&pOrderDate&"'"

call getFromDatabase(mySQL1, rsTemp, "saveorder 10")    
 
if rsTemp.eof then 
 response.redirect "comersus_supportError.asp?error="&Server.Urlencode("Error in saveorder 11. Cannot locate generated order with ran: "&pRandomNumber&", customer: "&pIdCustomer& " - Previous SQL: "&mySQL)
end if

pIdOrder=rsTemp("idOrder")

' save orderItems - assign idOrder to dbSessionCart and close the order

mySQL="UPDATE dbSessionCart SET idOrder=" &pIdOrder& ", cartOpen=0, idDbSession=NULL WHERE idDbSessionCart="&pIdDbSessionCart

call updateDatabase(mySQL, rsTemp, "saveorder 11")    

' delete dbSessionData inside dbSession
mySQL="UPDATE dbSession SET sessionData='*' WHERE idDbSession="&pIdDbSession
    
call updateDatabase(mySQL, rsTemp, "saveorder 12")    

if isOffLinePayment(pIdPayment) and pTotal > 0 then

 ' encrypt CC data
 pECardNumber=EnCrypt(pCardNumber, pEncryptionPassword)

 ' save credit card info
 mySQL="INSERT INTO creditCards (idOrder, cardType, cardNumber, expiration, seqCode) VALUES (" &pIdOrder& ",'" &pCardType& "','" &pECardNumber& "','" &pExpiration& "','" &pSeqCode& "')"
 call updateDatabase(mySQL, rstemp, "optOffLinePaymentExec")

end if

' compile email details text
customerEmail = Cstr("")
customerEmail = getMsg(482,"dear") &" "&pName&" "&pLastName&VBCrlf& getMsg(483,"thanks...") &VBCrlf&VBCrlf& getMsg(484,"order details")&VBCrlf&pDetails&VBCrlf&pPaymentDetails&VBCrlf&pShipmentDetails    

if pDiscountCode<>"" then
 customerEmail=customerEmail&VBCrlf&pDiscountDetails
end if

customerEmail=customerEmail&VBCrlf&getMsg(485,"taxes")&": "&pCurrencySign&money(pTaxAmount)&VBcrlf&"-------------------"&VBcrlf& getMsg(486,"total")&": " &" "&pCurrencySign& money(pTotal)&VBCrlf&VBCrlf

if trim(pEmailText)<>"" then 
 customerEmail=customerEmail &VBcrlf&VBcrlf& getMsg(487,"payment info")&": " & pEmailText
end if

' clear header cart variables
session("cartSubTotal") = 0
session("cartItems")	= 0

' clead discount code 
session("discountCode")	= ""

if pUseEncryptedTotal="-1" then
   pGatewayTotal = EnCrypt(pTotal, pEncryptionPassword)
else
   pGatewayTotal = pTotal
end if

pCustomerSubject = getMsg(488,"order at")&" " &pCompany& ", "&  " #" &pOrderPrefix&pIdorder

customerEmail	 =customerEmail&VBCrlf&VBCrlf&pCompany&VBCrlf&pEmailAdmin&VBcrlf&VBcrlf &getMsg(489,"payment confirm in other email")& VBcrlf&VBcrlf

if pDisableSaveOrderEmail<>"-1"  then
 ' send default customer email 
 call sendmail (pCompany, pEmailSender, pEmail,pCustomerSubject, customerEmail)
end if

' send email notification to the store admin
call sendmail (pCompany, pEmailSender, pEmailAdmin, "New order in your store #"&pOrderPrefix&pIdorder, "Details: "&VBCrlf&pDetails&VBCrlf& "Chargeback Abuse Verification: " & pChargeBackMSG & VBCrlf &pPreChargeMsg & VBCrlf &"Customer: "&pName&" "&pLastName&VBcrlf&"Email: "&pemail&VBCrlf&"Address: "&pAddress&VBCrlf&"Zip: "&pZip&VBCrlf&"State: "&pStateCode&pState&VBCrlf&"Country: "&pCountryCode&VBCrlf&"Phone #: "&pPhone&VBcrlf&"Comments: "&pComments&VBCrlf&"Payment: "&pPaymentDetails&VBCrlf&"Shipping: "&pShipmentDetails&VBCrlf&"Discounts: "&pDiscountDetails&VBcrlf&"Taxes: "&pCurrencySign&money(pTaxAmount)&VBcrlf&"Order total: "&pCurrencySign&money(pTotal)&VBcrlf&VBcrlf)

' mark order as paid if total is 0 
if pTotal=0 Then
  pSerialCodeOnlyOnce	= getSettingKey("pSerialCodeOnlyOnce")
  call paymentApproved(pCurrencySign&"0 order", pOrderPrefix&pIdorder, pCurrencySign&"0 order")	   
  response.redirect "comersus_defaultOrderConfirmation.asp?idOrder=" &pOrderPrefix&pIdOrder	
end if

' mark order as paid since bonus points were used
if pBonusPoints>0 then  
  
  session("usedBonusPoints")=-1
  pSerialCodeOnlyOnce	= getSettingKey("pSerialCodeOnlyOnce")
  call paymentApproved(getMsg(158,"Bonus Points"), pOrderPrefix&pIdorder, getMsg(158,"Bonus Points"))	   
  ' rest bonus points used to pay 
  call restBonusPoints(pBonusPoints, pIdCustomer)          	   
  response.redirect "comersus_defaultOrderConfirmation.asp?idOrder=" &pOrderPrefix&pIdOrder	
   
end if

' go to a payment form if is not offline payment
if trim(pRedirectionUrl)<>"" and pRedirectionUrl<>"comersus_offLinePaymentForm.asp" then   
   response.redirect pRedirectionUrl&"?idOrder="&pOrderPrefix&pIdorder& "&OrderTotal=" &pGatewayTotal& "&name=" &Server.UrlEncode(pName)& "&lastName=" &Server.UrlEncode(pLastName)& "&address=" &Server.UrlEncode(pAddress)& "&city=" &Server.UrlEncode(pCity)& "&state=" &Server.UrlEncode(pState&pStateCode)& "&zip=" &Server.UrlEncode(pZip)& "&country=" &Server.UrlEncode(pCountryCode)& "&phone=" &Server.UrlEncode(pPhone)& "&email=" &Server.UrlEncode(pEmail) & "&orderDetails=" &Server.UrlEncode(pDetails)&"&company="&Server.UrlEncode(pCustomerCompany)
else 
  ' redirect to default order confirmation
  response.redirect "comersus_defaultOrderConfirmation.asp?idOrder=" &pOrderPrefix&pIdOrder
end if

call closeDb()
%>