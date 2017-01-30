<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: verify email+password and redirect to user utilities menu or selected URL
%>

<!--#include file="../includes/screenMessages.asp" -->  
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/sessionFunctions.asp"-->  
<!--#include file="../includes/currencyFormat.asp" -->
<!--#include file="../includes/encryption.asp" -->  

<%
on error resume next

dim mySQL, conntemp, rstemp, pEmail, pPassword

' get from the form
pEmail 			= getLoginField(request("email"),50)
pPassword 		= getLoginField(request("password"),20)

pRedirect 		= getUserInputL(request("redirect"),20)
pIdProduct 		= getUserInputL(request("idProduct"),20)

if pEmail="Email" and pPassword="Pass" then
 pEmail=""
 pPassword=""
end if

' get settings

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pEncryptionPassword 	= getSettingKey("pEncryptionPassword")
pEncryptionMethod 	= getSettingKey("pEncryptionMethod")
pForgotPassword		= getSettingKey("pForgotPassword")

' check if everything was entered
if pEmail="" or pPassword="" then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(144,"Enter all required information"))
end if

' verify password for that email

mySQL="SELECT idCustomer, idCustomerType, name, lastName, active FROM customers WHERE email='" &pEmail& "' AND password='" &EnCrypt(pPassword, pEncryptionPassword)& "'"

call getFromDatabase(mySQL, rstemp, "customerAuthenticateExec") 

if rstemp.eof then
 
 if pForgotPassword="-1" then
  response.redirect "comersus_optForgotPasswordForm.asp?email="&Server.urlEncode(pEmail)
 else
  response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(145,"Incorrect login information"))
 end if
 
end if

if rstemp("active")=0 then
  response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(146,"You cannot login. Your profile is disabled."))
end if

session("idCustomer") 	  = rstemp("idCustomer")
session("idCustomerType") = rstemp("idCustomerType")
pName			  = rstemp("name")
pLastName		  = rstemp("lastName")
session("customerName")	  = pName & " " & pLastName

call closeDb()

if pRedirect="" then
  response.redirect "comersus_customerUtilitiesMenu.asp?name=" &Server.UrlEncode(pName)& "&lastName=" &Server.UrlEncode(pLastName)  
end if

if pRedirect="auction" then
  response.redirect "comersus_optAuctionListAll.asp"
end if

if pRedirect="wish" then
  response.redirect "comersus_customerWishListAdd.asp?idProduct="&pIdProduct
end if

if pRedirect="checkout" then
  response.redirect "comersus_checkout2.asp"
end if

%>
