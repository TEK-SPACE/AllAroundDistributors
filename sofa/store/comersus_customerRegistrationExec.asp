<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: register a new customer in the store 
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
pRandomPassword		= getSettingKey("pRandomPassword")
pEncryptionPassword 	= getSettingKey("pEncryptionPassword")
pEncryptionMethod 	= getSettingKey("pEncryptionMethod")
pTeleSignCustomerId 	= getSettingKey("pTeleSignCustomerId")
pBlockPOBoxes 		= getSettingKey("pBlockPOBoxes")

pEmailSender		= getSettingKey("pEmailSender")
pEmailAdmin		= getSettingKey("pEmailAdmin")
pSmtpServer		= getSettingKey("pSmtpServer")
pEmailComponent		= getSettingKey("pEmailComponent")
pDebugEmail		= getSettingKey("pDebugEmail")

' form parameters

pCustomerName	= formatForDb(getUserInput(request.form("customerName"),50))
pLastName	= formatForDb(getUserInput(request.form("lastName"),50))
pCustomerCompany= formatForDb(getUserInput(request.form("customerCompany"),50))
pEmail		= formatForDb(getUserInput(request.form("email"),80))
pPassword	= formatForDb(getUserInput(EnCrypt(request.form("password"), pEncryptionPassword),50))
pPhone		= formatForDb(getUserInput(request.form("phone"),20))

pAddress	= formatForDb(getUserInput(request.form("address"),100))
pCity		= formatForDb(getUserInput(request.form("city"),30))
pZip		= formatForDb(getUserInput(request.form("zip"),12))
pState		= formatForDb(getUserInput(request.form("state"),30))
pStateCode	= formatForDb(getUserInput(request.form("stateCode"),30))
pCountryCode	= formatForDb(getUserInput(request.form("countryCode"),4))

' retrieve custom fields
pUser1		= formatForDb(getUserInput(request.form("user1"),30))
pUser2		= formatForDb(getUserInput(request.form("user2"),30))
pUser3		= formatForDb(getUserInput(request.form("user3"),30))

' validation

if pCustomerName="" or pLastName="" or pEmail="" or pPassword="" or pPhone="" or pAddress="" or pCity="" or pZip="" or pCountryCode="" then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(251,"Please enter all the fields for your registration."))
end if

' email verification
if instr(pEmail,"@")=0 or instr(pEmail,".")=0 then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(252,"Invalid email entered"))
end if 

' state and non listed state entered
if pState<>"" and pStateCode<>"" then
   response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(253,"You cannot enter a listed and a non-listed state"))
end if

' PO boxes verification
pAddressTest=uCase(pAddress)
if pBlockPOBoxes="-1" and (instr(pAddressTest,"PO")<>0 or instr(pAddressTest,"P.O.")<>0) then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(750,"PO not valid"))
end if

if pRandomPassword="0" then
 ' password verification
 pUnencryptedPassword=getUserInput(request.form("password"),50)
 dim regEx
 set regEx = New RegExp
 regEx.IgnoreCase = True
 regEx.Global   = True
 regEx.Pattern = "[\W]"
 
 if regEx.Test(pUnencryptedPassword) then
   response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(761,"Only az 09"))
 end if
end if

' check if entered email already exists in database
mySQL="SELECT idCustomer FROM customers WHERE email='" &pEmail&"'"

call getFromDatabase(mySql, rsTemp,"customerRegistrationExec")

if not rstemp.eof then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(254,"The email was already used in this store. Please retrieve your password instead of registering again."))
end if

' for Telesign, save customer with active=false
if pTeleSignCustomerId = "0" then
    pActive = -1
else
    pActive = 0
end if

' insert to customers table with retail type by default (admin can change to BtoB later with BackOffice)
mySQL="INSERT INTO customers (name, lastName, customerCompany, email, city, countryCode, phone, address, zip, password, state, stateCode, idCustomerType, active, user1, user2, user3, bonusPoints, idStore) VALUES ('" &pcustomerName& "','" &pLastName& "','" &pCustomerCompany& "','" &pemail& "','" &pcity& "', '" &pCountryCode& "','" &pPhone& "','" &pAddress& "','" &pZip& "','" &pPassword& "','" &pState& "','" &pStateCode& "',1," & pActive & ",'" &pUser1& "','" &pUser2& "','" &pUser3& "',0,"&pIdStore&")"

call updateDatabase(mySQL, rsTemp, "customerRegistrationExec")

' obtain the idCustomer of the new record

mySQL="SELECT idCustomer FROM customers WHERE email='" &pEmail&"'"

call getFromDatabase(mySql, rsTemp, "customerRegistrationExec")


' retail customer by default
session("idCustomerType")  = 1

pMsg = getMsg(255,"Thanks. Your profile has been saved.")

if pRandomPassword="-1" then
    ' send password by email
    pPassword	= deCrypt(pPassword, pEncryptionPassword)
    call sendmail (pCompany, pEmailSender, pEmail, pCompany, getMsg(758,"Your reg for:") &pCompany& getMsg(759,"is complete")&vbCrlf&vbCrlf&getMsg(760,"You can login with:")&vbCrlf&"E-Mail= "&pEmail&vbCrlf&"Password= "&pPassword)
end if


if pTeleSignCustomerId <> "0" then
    response.redirect "comersus_optTeleSignForm.asp?idCustomer="&rstemp("idCustomer")& "&message="&Server.Urlencode(pMsg)&"&email="&Server.UrlEncode(pEmail)
else
    ' login the registered user
    session("idCustomer")= rstemp("idCustomer")  
    response.redirect "comersus_message.asp?message="&Server.Urlencode(pMsg)
end if

call closeDb()
%>
