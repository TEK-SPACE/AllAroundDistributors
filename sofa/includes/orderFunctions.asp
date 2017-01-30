<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' United States 
' http://www.comersus.com  
' Details: update stock, increase sales, send understock notifications, etc
%>
<!--#include file="serialFunctions.asp"--> 
<%
Function paymentApproved(pGatewayName, pIdOrder, pAuthCode)
 
 pSendDigitalGoodsByEmail	= getSettingKey("pSendDigitalGoodsByEmail")
 pDateFormat			= getSettingKey("pDateFormat")
 pCompanyPhone			= getSettingKey("pCompanyPhone")
 pEmailSender			= getSettingKey("pEmailSender")
 pEmailAdmin			= getSettingKey("pEmailAdmin")
 pSmtpServer			= getSettingKey("pSmtpServer")
 pEmailComponent		= getSettingKey("pEmailComponent")
 pDebugEmail			= getSettingKey("pDebugEmail")
 pOrderPrefix			= getSettingKey("pOrderPrefix") 
 pBonusPointsPerPrice		= getSettingKey("pBonusPointsPerPrice")
 pCurrencySign			= getSettingKey("pCurrencySign")
 pEmailNoStock          	= getSettingKey("pEmailNoStock")  
 pWholesaleConversionAmount	= getSettingKey("pWholesaleConversionAmount")   
 
 pDigitalGoodsText 		= Cstr("")
 pCustomerEmailBodyInvoice 	= Cstr("")
 pWishListEmail			= Cstr("")
 
 pStockDate			= Date()
 
 ' extract real idorder (without prefix)
 pIdOrder 			= removePrefix(pIdOrder,pOrderPrefix)

 if trim(pIdOrder)="" then
  call sendmail (pCompany, pEmailSender, pEmailAdmin, "Silent Response Error #"&pOrderPrefix&pIdorder, "idOrder not specified."&VBcrlf&VBcrlf & "All Data: "&pAllData)  
  response.redirect "comersus_supportError.asp?error="&Server.Urlencode("Error in "&pGatewayName&" SilentResponse. Order is not specified: "&pIdOrder)
 end if

 ' get order data

 mySQL="SELECT orderStatus, address, zip, city, state, stateCode, countryCode, obs, shipmentDetails, paymentDetails, discountDetails, taxAmount, total, details, wishListIdCustomer, idStore FROM orders WHERE idOrder=" &pIdOrder
 
 call getFromDatabase(mySQL, rstemp,pGatewayName& " Silent Response")  

 if rstemp.eof then
   call sendmail (pCompany, pEmailSender, pEmailAdmin, "Silent Response Error #"&pOrderPrefix&pIdorder, "Cannot find order data")
   response.redirect "comersus_supportError.asp?error="&Server.Urlencode("Error in "&pGatewayName&" SilentResponse. Cannot find order")
 end if
  
 pAddress		=rstemp("address")
 pZip			=rstemp("zip")
 pCity			=rstemp("City")
 pStateCode		=rstemp("stateCode")
 pState			=rstemp("state")
 pCountryCode		=rstemp("countryCode")
 pObs			=rstemp("obs")
 pShipmentDetails	=rstemp("shipmentDetails")
 pPaymentDetails	=rstemp("paymentDetails")
 pDiscountDetails	=rstemp("discountDetails")
 pTaxAmount		=rstemp("taxAmount")
 pTotal			=rstemp("total")
 pDetails		=rstemp("details")
 pWishListIdCustomer	=rstemp("wishListIdCustomer")
 pIdStore		=rstemp("idStore")
 
 if rstemp("orderStatus")<>1 then
    
    call sendmail (pCompany, pEmailSender, pEmailAdmin, "Silent Response Error #"&pOrderPrefix&pIdorder, "The order was already processed.")      
    response.redirect "comersus_supportError.asp?error="&Server.Urlencode("Error in "&pGatewayName&" SilentResponse. Order is not pending")
    
 end if
 
 ' get customer data 
 
 mySQL="SELECT customers.idCustomer, customers.idCustomerType, email, name, lastName, phone FROM customers, orders WHERE orders.idCustomer=customers.idCustomer AND orders.idOrder=" &pIdOrder

 call getFromDatabase(mySQL, rstemp, pGatewayName&" Silent Response")  
 
 if rstemp.eof then 
  call sendmail (pCompany, pEmailSender, pEmailAdmin, "Silent Response Error #"&pOrderPrefix&pIdorder, "Cannot locate idCustomer for the order.")
 end if

 pIdCustomer		= rstemp("idCustomer") 
 pCustomerName		= rstemp("name")& " "&rstemp("lastName")
 pCustomerEmail		= rstemp("email") 
 pCustomerPhone		= rstemp("phone") 
 pIdCustomerType	= rstemp("idCustomerType") 
 pCustomerEmailBody	= Cstr("")
 pCustomerEmailBody	= getMsg(638,"Payment for the order # ") & pOrderPrefix&pIdOrder &" " & getMsg(639," was accepted")

 ' if the order was placed by a third party for a valid wish list retrieve wish list customers data
 if pWishListIdCustomer<>0 then
 
  mySQL="SELECT email FROM customers WHERE idCustomer="&pWishListIdCustomer

  call getFromDatabase(mySQL, rstemp, pGatewayName&" Silent Response")  
 
  if rstemp.eof then 
    call sendmail (pCompany, pEmailSender, pEmailAdmin, "Silent Response Error #"&pOrderPrefix&pIdorder, "Cannot locate email for wish list.")
  end if
 
 pWishListEmail=rstemp("email")
end if
  
  ' iterates through cartRows
  
  mySQL="SELECT idProduct, quantity, personalizationDesc FROM cartRows, dbSessionCart WHERE dbSessionCart.idDbSessionCart=cartRows.idDbSessionCart and dbSessionCart.idOrder=" &pIdOrder    
  
  call getFromDatabase(mySQL, rsTemp2, pGatewayName&" Silent Response")       
         
  do while not rstemp2.eof  
  
   pPersonalizationDesc=rstemp2("personalizationDesc")    
  
   ' remove from wish list
   if pWishListIdCustomer<>0 then
     mySQL="DELETE FROM wishList WHERE idProduct=" &rstemp2("idProduct")& " AND idCustomer="&pWishListIdCustomer  
     call updateDatabase(mySQL, rsTemp, pGatewayName&" Silent Response")      
   end if
  
   pCustomerEmailText=""
  
   mySQL="SELECT sales, description, emailText, rental FROM products WHERE idProduct=" &rstemp2("idProduct")
  
   call getFromDatabase(mySQL, rsTemp, pGatewayName&" Silent Response")             
    
   ' field used for serials, download links, etc
   pEmailText  		 =trim(rstemp("emailText")) 
   pRental		 =rstemp("rental")  
   pDescription     	 =rstemp("description")
   pStock		 =getStock(rstemp2("idProduct"))
   
  if pEmailText<>"" and pSendDigitalGoodsByEmail="-1" then
   pDigitalGoodsText=pDigitalGoodsText&pDescription&" -> "&pEmailText&Vbcrlf
  end if
  
  pTextEmailNoStock=""
  
  ' add no stock text to payment confirmation email
  if pStock<1 and pEmailText="" and pEmailNoStock = "-1" then             
    pTextEmailNoStock = getMsg(640,"** Some items in your order have no stock. Please give us enough time to prepare your shipment.")
  end if 
  
   ' try to get serial codes (if any)
  
   pDigitalGoodsText=pDigitalGoodsText&getSerials(rstemp2("idProduct"),rstemp2("quantity"))
      
   ' insert stock movement
   mySQL="INSERT INTO stockMovements (dateMovement, idProduct, quantity, obs) VALUES ('" &pStockDate& "'," &rstemp2("idProduct")& ",-" &rstemp2("quantity")& ",'#" &pOrderPrefix&pIdOrder&"')"
   
   call updateDatabase(mySQL, rstemp, pGatewayName&" Silent Response")               
   
   ' decrease stock     
   call updateStock(rstemp2("idProduct"),-rstemp2("quantity"))
   
   ' get idSupplier for the item
    
    mySQL="SELECT suppliers.idSupplier, supplierName, supplierEmail, receiveUnderStockAlert FROM suppliers, products WHERE products.idSupplier=suppliers.idSupplier AND idProduct=" &rstemp2("idProduct")
    
    call getFromDatabase(mySQL, rstemp, pGatewayName&" Silent Response")                
    
    if not rstemp.eof then
     if rstemp("receiveUnderStockAlert")=-1 and rstemp("supplierEmail")<>"" and getStock(rstemp2("idProduct"))<=0 then
      ' send email to supplier with under stock alert
      pSupplierEmail = rstemp("supplierEmail")
      pSupplierName  = rstemp("supplierName")
      call sendmail (pCompany, pEmailSender, pSupplierEmail, getMsg(641,"Under Stock Email Alert"), getMsg(642,"Hi ")&pSupplierName&", " & getMsg(643,"we need to purchase")&pDescription& " "& getMsg(644,"Please contact us ASAP to re-order.")&VBCrlf&getMsg(645,"Our phone number is: ") & " " &pCompanyPhone&VBCrlf&pCompany&VBCrlf&pEmailAdmin&VBcrlf) 
     end if    
    end if
 
  ' update sales   
  
   mySQL="UPDATE products SET sales=sales+" &rstemp2("quantity")&" WHERE idProduct=" &rstemp2("idProduct")
  
   call updateDatabase(mySQL, rsTemp3, pGatewayName&" Silent Response")           
   
   ' rental products
   if pRental=-1 then
  
     ' parse from/until
     pArrayDates	=split(pPersonalizationDesc,"|")
     pFrom		=pArrayDates(0)
     pUntil		=pArrayDates(1)
     pIsAvailable	=-1
     pReason		=""
     pIdProduct		=rstemp2("idProduct")
     
     pQuantity	=0
     
    ' verify if rental period is still valid at payment time
    call checkRentalAvailability(pIdProduct, pFrom, pUntil, pIsAvailable, pReason, pQuantity)
   
   if pIsAvailable=-1 then
    ' insert into rentals
    	'Get sku of rental item
    Mysql = "SELECT sku FROM products WHERE idProduct = " & rstemp2("idProduct")
    call getFromDatabase(mySQL, rsTemp, pGatewayName&" Silent Response")             
		pSku = rstemp("sku")
		
   	'Get rental products with same sku
    Mysql = "SELECT idProduct FROM products WHERE sku = '" & pSku & "' AND rental=-1"
    call getFromDatabase(mySQL, rsTemp, pGatewayName&" Silent Response")             
		
		'Insert rental reservation for all items with same sku
		Do while not rstemp.eof		
	    mySQL="INSERT INTO rentals (idProduct, fromDate, untilDate, idOrder) VALUES (" &rsTemp("idProduct")& ",'" &pFrom& "','" &pUntil& "'," &pIdOrder& ")"      
	    call updateDatabase(mySQL, rsTemp3, "updateTransactionResults.asp")          
    	rstemp.movenext
    Loop
   else
   ' send notification
    call sendmail (pCompany, pEmailSender, pEmailAdmin, "Silent Response Rental Error #"&pOrderPrefix&pIdorder, "Reservation conflict in Silent Response for idOrder: "&pIdOrder&", idProduct: "&pIdProduct&" Reason: "&pReason&VBcrlf)
   end if
   
   end if
  
   rstemp2.movenext
 loop
 
 ' change order status
 mySQL="UPDATE orders SET orderStatus=4, transactionResults='" &getMsg(646,"Authorization Code:") & " " &pAuthCode& "' WHERE idOrder=" &pIdOrder

 call updateDatabase(mySQL, rstemp, pGatewayName&" Silent Response")     

 ' if bonus points were not used to pay the order, add bonus points
 if instr(pAuthCode,getMsg(158,"Bonus Points"))=0 then
 
  if pBonusPointsPerPrice>0 then
   mySQL="UPDATE customers SET bonusPoints=bonusPoints + (" &pBonusPointsPerPrice& "*" &formatNumberForDb(pTotal)& ") WHERE idCustomer=" &pIdCustomer
   call updateDatabase(mySQL, rsTemp, "saveOrder")      
  end if
  
 end if
 
  
 ' add no stock message to payment confirmation email
 if pEmailNoStock<>"0" then
    pCustomerEmailBody = pCustomerEmailBody & Vbcrlf &pEmailNoStock
 end if
 
 ' check wholesale conversion
 dim pTextEmailWholeSale
 pTextEmailWholeSale=""
 if pIdCustomerType=1 and Cdbl(pWholesaleConversionAmount)>0 and Cdbl(getCustomerTotals(pIdCustomer))>Cdbl(pWholesaleConversionAmount) then
  mySQL="UPDATE customers SET idCustomerType=2 WHERE idCustomer=" &pIdCustomer
  call updateDatabase(mySQL, rsTemp3, "orderFunctions.asp")  
  pTextEmailWholeSale=getMsg(754,"Congratulations")
 end if
 
 ' create the invoice design
 pCustomerEmailBodyInvoice=                                 getMsg(647,"Order: ")&" "&pOrderPrefix&pIdorder
 pCustomerEmailBodyInvoice=pCustomerEmailBodyInvoice&Vbcrlf&getMsg(648,"Name: ")&" "&pCustomerName
 pCustomerEmailBodyInvoice=pCustomerEmailBodyInvoice&Vbcrlf&getMsg(649,"Email: ")&" "&pCustomerEmail
 pCustomerEmailBodyInvoice=pCustomerEmailBodyInvoice&Vbcrlf&getMsg(650,"Phone: ")&" "&pCustomerPhone
 pCustomerEmailBodyInvoice=pCustomerEmailBodyInvoice&Vbcrlf&getMsg(651,"Address: ")&" "&pAddress&" - "&pZip&" - " &pState&pStateCode&" - "&pCountryCode 
 pCustomerEmailBodyInvoice=pCustomerEmailBodyInvoice&Vbcrlf&getMsg(652,"Comments: ")&" "&pObs
 pCustomerEmailBodyInvoice=pCustomerEmailBodyInvoice&Vbcrlf&getMsg(653,"Payment: ")&" "&pPaymentDetails
 pCustomerEmailBodyInvoice=pCustomerEmailBodyInvoice&Vbcrlf&getMsg(654,"Shipment: ")&" "&pShipmentDetails
 pCustomerEmailBodyInvoice=pCustomerEmailBodyInvoice&Vbcrlf&getMsg(655,"Discounts: ")&" "&pDiscountDetails
 pCustomerEmailBodyInvoice=pCustomerEmailBodyInvoice&Vbcrlf&getMsg(656,"Taxes: ")&" "&pCurrencySign&money(pTaxAmount)
 pCustomerEmailBodyInvoice=pCustomerEmailBodyInvoice&Vbcrlf&getMsg(657,"Total: ")&" "&pCurrencySign&money(pTotal)
 
 ' add invoice to the main email
 pCustomerEmailBody=pCustomerEmailBody&Vbcrlf&Vbcrlf&pCustomerEmailBodyInvoice
 
 if pTextEmailNoStock<>"" then
  pCustomerEmailBody = pCustomerEmailBody &Vbcrlf&Vbcrlf&pTextEmailNoStock  
 end if
 
 ' add digital email text like downloads, serials, etc only if was not a purchase for a wish list
 if len(pDigitalGoodsText)>0 and trim(pWishListEmail)="" then
   pCustomerEmailBody = pCustomerEmailBody &Vbcrlf&Vbcrlf&getMsg(658,"** Digital goods in this order: ")&Vbcrlf&pDigitalGoodsText   
 end if
 
 ' send invoice, payment confirmation and Digital Goods (serials, download links, etc)
 call sendmail (pCompany, pEmailSender, pCustomerEmail, getMsg(659,"Payment approved, order #") &pOrderPrefix&pIdorder, pCustomerEmailBody&VBCrlf&VBCrlf&pCompany&VBCrlf&pEmailAdmin&VBcrlf&VBcrlf)   
 
 ' send notification to the wish list owner and digital goods if any
 if pWishListEmail<>"" then
   pWishListEmailBody=pCustomerName&" "&pCustomerLastName &" "&getMsg(660,"has placed an order for an item of your Wish List.")&Vbcrlf&pDetails
   
   if len(pDigitalGoodsText)>0 then
    pWishListEmailBody=pWishListEmailBody&Vbcrlf&Vbcrlf&getMsg(658,"** Digital goods in this order: ")&Vbcrlf&pDigitalGoodsText   
   end if
   
   call sendmail (pCompany, pEmailSender, pCustomerEmail, getMsg(647,"Order # ")& " " &pOrderPrefix&pIdorder, pWishListEmailBody&VBCrlf&VBCrlf&pCompany&VBCrlf&pEmailAdmin&VBcrlf&VBcrlf)   
 end if     
  
end function 

function getCustomerTotals(pIdCustomer)

 dim mySQL, rstemp
 
 mySQL="SELECT SUM(total) AS sumTotal FROM orders, customers WHERE customers.idCustomer=" &pIdCustomer& " AND orders.idCustomer=customers.idCustomer AND orders.orderStatus=4"
 call getFromDatabase(mySQL, rstemp, pGatewayName&" Silent Response") 
 
 getCustomerTotals=rstemp("sumTotal")
 
end function

 %>