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
<!--#include file="comersus_optDes.asp"-->    
<!--#include file="../includes/stringFunctions.asp"-->

<%
on error resume next

dim connTemp, rsTemp, mySql

' get settings

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pEncryptionPassword 	= getSettingKey("pEncryptionPassword")
pEncryptionMethod	= getSettingKey("pEncryptionMethod")
pDisableState		= getSettingKey("pDisableState")
pBlockPOBoxes 		= getSettingKey("pBlockPOBoxes")

' get from session
pIdCustomer     	= getSessionVariable("idCustomer","0")

' form data
pCustomerName		= formatForDb(getUserInput(request.form("customerName"),30))
pLastName		= formatForDb(getUserInput(request.form("LastName"),30))
pCustomerCompany	= formatForDb(getUserInput(request.form("customerCompany"),30))
pEmail			= getUserInput(request.form("email"),50)
pPassword		= getUserInput(EnCrypt(request.form("password"), pEncryptionPassword),50)
pAddress		= formatForDb(getUserInput(request.form("address"),50))
pCity			= formatForDb(getUserInput(request.form("city"),30))
pZip			= getUserInput(request.form("zip"),20)
pState			= formatForDb(getUserInput(request.form("state"),20))
pStateCode		= getUserInput(request.form("stateCode"),4)
pPhone			= getUserInput(request.form("phone"),20)
pCountryCode		= getUserInput(request.form("countryCode"),4)
pUser1			= formatForDb(getUserInput(request.form("user1"),30))
pUser2			= formatForDb(getUserInput(request.form("user2"),30))
pUser3			= formatForDb(getUserInput(request.form("user3"),30))
pRedirect		= getUserInput(request.form("redirect"),20)

if pStoreFrontDemoMode="-1" then
  response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(188,"Option disabled for this demo."))
end if

if pState<>"" and pStateCode<>"" then
   response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(189,"Please enter state code or another state. You cannot select both."))
end if

' email verification
if instr(pEmail,"@")=0 or instr(pEmail,".")=0 then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(252,"Invalid email entered"))
end if 

pAddressTest=uCase(pAddress)
if pBlockPOBoxes="-1" and (instr(pAddressTest,"PO")<>0 or instr(pAddressTest,"P.O.")<>0) then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(750,"PO not valid"))
end if

' get current email
mySQL="SELECT email FROM customers WHERE idCustomer=" &pIdCustomer
call getFromDatabase(mySql, rsTemp,"customerModifyExec")

if not rstemp.eof then 
 pCurrentEmail = rsTemp("email")
end if

if pEmail <> pCurrentEmail then
    ' check if new email is in db
    mySQL="SELECT idCustomer FROM customers WHERE email='" &pEmail&"'"
    
    call getFromDatabase(mySql, rsTemp,"customerModifyExec")
    
    if not rstemp.eof then
     response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(254,"The email was already used in this store. Please retrieve your password instead of registering again."))
    end if
end if

' update welcome back session variable

session("customerName")	  = pCustomerName & " " & pLastName

' update customer record
		
mySql="UPDATE customers SET name='" &pCustomerName& "',lastName='" &pLastName&"', customerCompany='" &pCustomerCompany& "', email='" &pEmail& "', password='" &pPassword&"',city='" &pCity& "',zip='" &pZip& "',countryCode='" &pCountryCode& "',state='" &pState& "', stateCode='" &pStateCode& "',phone='" &pPhone& "',address='" &pAddress& "', user1='" &pUser1& "', user2='" &pUser2& "', user3='" &pUser3& "' WHERE idCustomer="&pIdCustomer
 	
call updateDatabase(mySQL, rstemp, "customerModifyExec")

if pRedirect="" then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(190,"Your personal information was updated."))
else
 response.redirect "comersus_checkout2.asp"
end if 

call closeDb()
%>