<%
' Comersus 6.0x Sophisticated Cart 
' Developed by Rodrigo S. Alhadeff for Comersus Open Technologies
' United States 
' Open Source License can be found at License.txt
' http://www.comersus.com  
' Details: Fraud Prevention routines 
%>

<%

function rejectOrder()

'on error resume next

dim connTemp, rsTmpFraud, mySql

rejectOrder=""

' get settings

pStoreFrontDemoMode 		= getSettingKey("pStoreFrontDemoMode")
pCompany	 		= getSettingKey("pCompany")
pFraudPreventionExpensiveItem	= getSettingKey("pFraudPreventionExpensiveItem")
pFraudPreventionExpensiveOrder	= getSettingKey("pFraudPreventionExpensiveOrder")
pFraudPreventionMode    	= lcase(getSettingKey("pFraudPreventionMode"))

pEmailSender			= getSettingKey("pEmailSender")
pEmailAdmin			= getSettingKey("pEmailAdmin")
pSmtpServer			= getSettingKey("pSmtpServer")
pEmailComponent			= getSettingKey("pEmailComponent")
pDebugEmail			= getSettingKey("pDebugEmail")

' get from session
pIdCustomer     		= getSessionVariable("idCustomer","0")
pIdDbSession	 		= checkSessionData()
pIdDbSessionCart 		= checkDbSessionCartOpen()

pUserIp	 			= getUserInput(request.ServerVariables("REMOTE_HOST"),20)
pFiltersReturned		= ""

' verify locked keywords.
mySQL="SELECT name, lastName, address, email, customerCompany FROM customers WHERE idCustomer=" &pIdCustomer

call getFromDatabase(mySQL, rsTmpFraud, "comersus_optFraudPrevention") 

if rsTmpFraud.eof then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(711,"Invalid customer inside fraud prevention"))
end if

pCustomerName=rsTmpFraud("name")&" "&rsTmpFraud("lastName")

' change to lower case 4 comparison
pSearchString = lcase(rsTmpFraud("name"))&"|"&lcase(rsTmpFraud("lastName"))&"|"&lcase(rsTmpFraud("email"))&"|"&lcase(rsTmpFraud("address"))&"|"&lcase(rsTmpFraud("customerCompany"))

mySQL="SELECT keyWordText FROM blockKeywords "
call getFromDatabase(mySQL, rsTmpFraud, "comersus_optFraudPrevention") 
pBlockKeyFound = 0

do while not rsTmpFraud.eof and pBlockKeyFound = 0 
	if instr(pSearchString, lcase(rsTmpFraud("keyWordText")))>0 then
		pBlockKeyFound = 1
	end if
	rsTmpFraud.movenext
loop

if pBlockKeyFound = 1 then
	pFiltersReturned = pFiltersReturned & getMsg(712,"Blocked Keywords found in the profile")
end if

' verify expensive item product.
if findProduct(pIdDbSessionCart, pFraudPreventionExpensiveItem) = -1 then 
	pFiltersReturned = pFiltersReturned & getMsg(713,"Expensive Item found in this order")
end if

' verify cart amount.
if cSng(calculateCartTotal(pIdDbSessionCart)) > cSng(pFraudPreventionExpensiveOrder) then 
	pFiltersReturned = pFiltersReturned & getMsg(714,"Expensive order total")
end if

' verify if the user ip is local ip
if  LEFT(pUserIp,4) = "192." OR _
	LEFT(pUserIp,4) = "127." OR _
	LEFT(pUserIp,4) = "172." OR _
	LEFT(pUserIp,3) = "10." Then
	pFiltersReturned = pFiltersReturned & getMsg(715,"Suspicious IP: ")&pUserIp
end if

if pFiltersReturned<>"" and pFraudPreventionMode="paranoid" then  
 call sendmail (pCompany, pEmailSender, pEmailAdmin,"Bloqued checkout", "Customer: "&pCustomerName&vbcrlf&"Report: "&pFiltersReturned) 
 rejectOrder=pFiltersReturned
end if

if pFiltersReturned<>"" and pFraudPreventionMode="caution" then 
 call sendmail (pCompany, pEmailSender, pEmailAdmin,"Suspicious checkout", "Customer: "&pCustomerName&vbcrlf&"Report: "&pFiltersReturned) 
end if

end function

%>