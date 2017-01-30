<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' USA - 2006
' http://www.comersus.com 
' Details: payment method functions
%>

<%

function createArrayPayments(pCartQuantity, pCartTotalWeight, pSubTotal)

' retrieve available payment methods, load a string array with | as row separator

dim mySQL, rstemp, pPaymentSurchargeAmount

pIdCustomerType 	= getSessionVariable("idCustomerType",1)
pToken			= getSessionVariable("token","0")
pBonusPointsUsed	= getSessionVariable("bonusPointsUsed",0)

' check bonus points available
pBonusPointsAvailable		=0
pBonusPointsAvailable		=getBonusPoints(pIdCustomer)

createArrayPayments	=""

' regular payments
if pToken="0" then
 
 mySQL="SELECT idPayment, paymentDesc, priceToAdd, percentageToAdd, redirectionUrl, emailText, quantityFrom, quantityUntil, priceFrom, priceUntil, weightFrom, weightUntil FROM payments WHERE (idCustomerType=" &pIdCustomerType& " OR idCustomerType IS NULL OR idCustomerType=0) AND idStore=" &pIdStore& " ORDER by paymentDesc"

 call getFromDatabase(mySQL, rstemp, "orderForm")

 if  rstemp.eof then 
   response.redirect "comersus_supportError.asp?error="&Server.Urlencode("There are no payments in database")
 end if

 do until rstemp.eof                        
      
      pDbquantityFrom	= Cdbl(rstemp("quantityFrom"))
      pDbquantityUntil	= Cdbl(rstemp("quantityUntil"))
            
      pDbpriceFrom	= Cdbl(rstemp("priceFrom")) 
      pDbpriceUntil	= Cdbl(rstemp("priceUntil"))
      
      pDbweightFrom	= Cdbl(rstemp("weightFrom"))
      pDbweightUntil	= Cdbl(rstemp("weightUntil"))            
      
      ' insert if all rules are ok            
      
      if pCartQuantity>=pDbquantityFrom and pCartQuantity<=pDbquantityUntil and pCartTotalWeight>=pDbweightFrom and pCartTotalWeight<=pDbweightUntil and pSubTotal>=pDbpriceFrom and pSubTotal<=pDbpriceUntil then	              	   	   
	   pPaymentSurchargeAmount	=Cdbl(rstemp("priceToAdd")) + Cdbl(rstemp("percentageToAdd"))*pSubTotal/100	   
	   createArrayPayments		=createArrayPayments&rstemp("paymentDesc")&"&"&pPaymentSurchargeAmount&"&"&rstemp("idPayment")&"|"
      end if
	
    rstemp.movenext
loop

end if

' get Express Checkout ID
if pToken<>"0" then
   mySQL="SELECT idPayment, paymentDesc FROM payments WHERE redirectionUrl='comersus_gatewayPayPalExpress3.asp' AND idStore=" &pIdStore
   
   call getFromDatabase(mySQL, rstemp, "paymentFunctions")
   
   if not rstemp.eof then
    pPaymentSurchargeAmount=0   
    createArrayPayments=rstemp("paymentDesc")&"&"&pPaymentSurchargeAmount&"&"&rstemp("idPayment")&"|"
   end if
   
end if 
 
' bonus points

if pBonusPointsAvailable>=pSubTotal then
   mySQL="SELECT idPayment, paymentDesc FROM payments WHERE redirectionUrl='comersus_gatewayBonusPoints.asp' AND idStore=" &pIdStore
   
   call getFromDatabase(mySQL, rstemp, "paymentFunctions")
   
   if not rstemp.eof then
    pPaymentSurchargeAmount=0   
    createArrayPayments=createArrayPayments&rstemp("paymentDesc")&"&"&pPaymentSurchargeAmount&"&"&rstemp("idPayment")&"|"
   end if
   
end if 

end function

function isOffLinePayment(pIdPayment)

dim rstemp, mysql

mySQL="SELECT redirectionUrl FROM payments WHERE idPayment=" &pIdPayment
call getFromDatabase(mySQL, rstemp, "isOffLinePayment")

isOffLinePayment=0

if not rstemp.eof then
 if instr(rstemp("redirectionUrl"),"offLinePaymentForm.asp")<>0 then
  isOffLinePayment=-1 
 end if
end if

end function

function isBonusPayment(pIdPayment)
 dim rstemp, mysql

 mySQL="SELECT redirectionUrl FROM payments WHERE idPayment=" &pIdPayment
 call getFromDatabase(mySQL, rstemp, "isOffLinePayment")

 isBonusPayment=0

 if not rstemp.eof then
  if instr(rstemp("redirectionUrl"),"comersus_gatewayBonusPoints.asp")<>0 then
   isBonusPayment=-1 
  end if
 end if

end function

%>

