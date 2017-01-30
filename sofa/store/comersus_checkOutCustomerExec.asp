<%
' Comersus 6.0x Sophisticated Cart 
' Developed by Rodrigo S. Alhadeff for Comersus Open Technologies
' United States 
' Open Source License can be found at License.txt
' http://www.comersus.com  
' Details: save customer profile and continue checkout
%>

<!--#include file="../includes/settings.asp" --> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/encryption.asp" --> 
<!--#include file="../includes/screenMessages.asp" --> 
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/sendMail.asp"-->

<% 

on error resume next

dim  mySQL, connTemp, rsTemp

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pEncryptionPassword 	= getSettingKey("pEncryptionPassword")
pEncryptionMethod 	= getSettingKey("pEncryptionMethod")
pRandomPassword		= getSettingKey("pRandomPassword")
pByPassShipping		= getSettingKey("pByPassShipping")
pBlockPOBoxes 		= getSettingKey("pBlockPOBoxes")

pEmailSender		= getSettingKey("pEmailSender")
pEmailAdmin		= getSettingKey("pEmailAdmin")
pSmtpServer		= getSettingKey("pSmtpServer")
pEmailComponent		= getSettingKey("pEmailComponent")
pDebugEmail		= getSettingKey("pDebugEmail")

' form parameters

pName			= formatForDb(getUserInput(request.form("name"),50))
pLastName		= formatForDb(getUserInput(request.form("lastName"),50))
pCustomerCompany	= formatForDb(getUserInput(request.form("customerCompany"),50))
pEmail			= formatForDb(getUserInput(request.form("email"),80))
pPassword		= formatForDb(getUserInput(EnCrypt(request.form("password"), pEncryptionPassword),50))
pPassword2		= formatForDb(getUserInput(EnCrypt(request.form("password2"), pEncryptionPassword),50))
pPhone			= formatForDb(getUserInput(request.form("phone"),20))

pAddress		= formatForDb(getUserInput(request.form("address"),100))
pCity			= formatForDb(getUserInput(request.form("city"),30))
pZip			= formatForDb(getUserInput(request.form("zip"),12))
pState			= formatForDb(getUserInput(request.form("state"),30))
pStateCode		= formatForDb(getUserInput(request.form("stateCode"),30))
pCountryCode		= formatForDb(getUserInput(request.form("countryCode"),4))

pShippingAddress	= formatForDb(getUserInput(request.form("shippingAddress"),100))
pShippingCity		= formatForDb(getUserInput(request.form("shippingCity"),30))
pShippingZip		= formatForDb(getUserInput(request.form("shippingZip"),12))
pShippingState		= formatForDb(getUserInput(request.form("shippingState"),30))
pShippingStateCode	= formatForDb(getUserInput(request.form("shippingStateCode"),30))
pShippingCountryCode	= formatForDb(getUserInput(request.form("shippingCountryCode"),4))

' retrieve custom fields
pUser1			= formatForDb(getUserInput(request.form("user1"),30))
pUser2			= formatForDb(getUserInput(request.form("user2"),30))
pUser3			= formatForDb(getUserInput(request.form("user3"),30))

' validation

if pName="" or pLastName="" or pEmail="" or pPhone="" or pAddress="" or pCity="" or pZip="" or pCountryCode="" then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(251,"Please enter all the fields for your registration."))   
end if

' PO box verification
pAddressTest=uCase(pAddress)
if pBlockPOBoxes="-1" and (instr(pAddressTest,"PO")<>0 or instr(pAddressTest,"P.O.")<>0) then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(750,"PO not valid"))
end if

pAddressTest=uCase(pShippingAddress)
if pBlockPOBoxes="-1" and (instr(pAddressTest,"PO")<>0 or instr(pAddressTest,"P.O.")<>0) then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(750,"PO not valid"))
end if

' pass
if pRandomPassword="0" then
 
 if pPassword="" then
  response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(251,"Please enter all the fields for your registration."))   
 end if

 if pPassword<>pPassword2 then
  response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(623,"Password and password verification are not the same."))
 end if
 
else

 ' autogenerate password
 pPassword=EnCrypt(randomNumber(999999),pEncryptionPassword)
 
end if

' email verification
if instr(pEmail,"@")=0 or instr(pEmail,".")=0 then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(252,"Invalid email entered"))
end if 

' verify only state or stateCode
if pState<>"" and pStateCode<>"" then
   response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(624,"You cannot enter a state and a non listed state."))
end if

' check if the entire shipping address was entered
if pShippingAddress<>"" then
 if pShippingZip="" or pShippingCity="" or pShippingCountryCode="" then response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(625,"Please enter the complete shipping address"))
end if

' check if entered email already exists in database
mySQL="SELECT idCustomer FROM customers WHERE email='" &pEmail&"'"

call getFromDatabase(mySql, rsTemp,"customerRegistrationExec")

if not rstemp.eof then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(626,"Sorry, your email was already used in this store. Retrieve your password or register with a another email")) 
end if

' insert to customers table with retail customer type by default (admin can change to BtoB later with BackOffice)

mySQL="INSERT INTO customers (name, lastName, customerCompany, email, city, countryCode, phone, address, zip, password, state, stateCode, idCustomerType, active, user1, user2, user3, bonusPoints, idStore) VALUES ('" &pName& "','" &pLastName& "','" &pCustomerCompany& "','" &pEmail& "','" &pCity& "', '" &pCountryCode& "','" &pPhone& "','" &pAddress& "','" &pZip& "','" &pPassword& "','" &pState& "','" &pStateCode& "',1,-1,'" &pUser1& "','" &pUser2& "','" &pUser3& "',0,"&pIdStore&")"

call updateDatabase(mySQL, rsTemp, "customerRegistrationExec")

' obtain the idCustomer of the new record

mySQL="SELECT idCustomer FROM customers WHERE email='" &pEmail&"'"

call getFromDatabase(mySql, rsTemp, "customerRegistrationExec")

if rstemp.eof then response.redirect "comersus_supportError.asp?error="&Server.Urlencode("Cannot locate customer ID inside checkoutCustomerExec for "&pEmail)

pIdCustomer=rstemp("idCustomer")  
' login the registered user
session("idCustomer")	 = pIdCustomer

' BtoC (retail) customer by default
session("idCustomerType")  = 1

' update shipping address

if pShippingAddress<>"" then
 mySQL="UPDATE customers SET shippingAddress='" &pShippingAddress& "', shippingZip='" &pShippingZip& "', shippingStateCode='" &pShippingStateCode& "', shippingState='" &pShippingState& "', shippingCity='" &pShippingCity&"', shippingCountryCode='" &pShippingCountryCode& "' WHERE idCustomer=" &pIdCustomer
 call updateDatabase(mySQL, rsTemp, "customerRegistrationExec")
end if

if pRandomPassword<>"0" then
	'Send autogenerated password by mail
	 pPassword		= deCrypt(pPassword, pEncryptionPassword)
	 call sendmail (pCompany, pEmailSender, pEmail,"Your password for store " & pCompany, "Your registration in "&pCompany&" is complete."&vbCrlf&vbCrlf&"You can now login with:"&vbCrlf&"E-Mail= "&pEmail&vbCrlf&"Password= "&pPassword)
end if
call closeDb()	
 

response.redirect "comersus_checkOut2.asp"

%>
