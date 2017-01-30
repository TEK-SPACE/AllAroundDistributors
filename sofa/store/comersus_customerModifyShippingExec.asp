<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: update personal info for a logged customer
%>

<!--#include file="comersus_customerLoggedVerify.asp"-->   
<!--#include file="../includes/settings.asp"-->    
<!--#include file="../includes/screenMessages.asp"-->    
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->  
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/encryption.asp"-->    
<!--#include file="../includes/stringFunctions.asp"-->

<%
on error resume next

dim connTemp, rsTemp, mySql

' get settings

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pDisableState		= getSettingKey("pDisableState")
pBlockPOBoxes 		= getSettingKey("pBlockPOBoxes")

' get from session
pIdCustomer     	= getSessionVariable("idCustomer","0")

' form data
pShippingName			= formatForDb(getUserInput(request.form("shippingName"),30))
pShippingLastName		= formatForDb(getUserInput(request.form("shippingLastName"),30))
pShippingAddress		= formatForDb(getUserInput(request.form("shippingaddress"),50))
pShippingCity			= formatForDb(getUserInput(request.form("shippingcity"),30))
pShippingZip			= getUserInput(request.form("shippingzip"),20)
pShippingState			= formatForDb(getUserInput(request.form("shippingstate"),20))
pShippingStateCode		= getUserInput(request.form("shippingstateCode"),4)
pShippingCountryCode		= getUserInput(request.form("shippingcountryCode"),4)

if pStoreFrontDemoMode="-1" then
  response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(188,"Option disabled for this demo."))
end if

if pshippingState<>"" and pshippingStateCode<>"" then
   response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(189,"Please enter state code or another state. You cannot select both."))
end if

' PO Box
pAddressTest=uCase(pShippingAddress)
if pBlockPOBoxes="-1" and (instr(pAddressTest,"PO")<>0 or instr(pAddressTest,"P.O.")<>0) then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(750,"PO not valid"))
end if

' update customer record
		
mySql="UPDATE customers SET shippingName='" &pShippingName& "', shippingLastName='" &pshippingLastName&"', shippingCity='" &pshippingCity& "',shippingzip='" &pshippingZip& "',shippingcountryCode='" &pshippingCountryCode& "',shippingstate='" &pshippingState& "', shippingstateCode='" &pshippingStateCode& "',shippingAddress='" &pshippingAddress& "' WHERE idCustomer="&pIdCustomer
 	
call updateDatabase(mySQL, rstemp, "customerModifyShippingExec")

response.redirect "comersus_checkout2.asp"

call closeDb()
%>