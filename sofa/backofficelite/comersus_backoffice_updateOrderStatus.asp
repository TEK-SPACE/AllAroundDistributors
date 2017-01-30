<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"-->  
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"-->   
<!--#include file="../includes/stringFunctions.asp"-->   
<!--#include file="../includes/orderFunctions.asp"-->   
<!--#include file="../includes/itemFunctions.asp"-->   
<!--#include file="../includes/screenMessages.asp"-->   
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->  
<!--#include file="../includes/settings.asp"-->    
<!--#include file="../includes/sendMail.asp"-->    

<% 
on error resume next 

dim mySQL, conntemp, rstemp, rstemp2

' get settings 
pDefaultLanguage 	= getSettingKey("pDefaultLanguage")
pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanyLogo	 	= getSettingKey("pCompanyLogo")

pEmailSender		= getSettingKey("pEmailSender")
pEmailAdmin		= getSettingKey("pEmailAdmin")
pSmtpServer		= getSettingKey("pSmtpServer")
pEmailComponent		= getSettingKey("pEmailComponent")
pDebugEmail		= getSettingKey("pDebugEmail")
pOrderPrefix	 	= getSettingKey("pOrderPrefix")

pIdOrder 		= getUserInput(request.form("idOrder"),12)
pOrderStatus 		= getUserInput(request.form("orderStatus"),2)
pOldOrderStatus		= getUserInput(request.form("oldOrderStatus"),2)

if trim(pIdOrder)="" or trim(pOrderStatus)="" then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Please enter the order and the new status")
end if

' check previous status 
if pOldOrderStatus<>1 and pOrderStatus=4 then
 response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("You can only change to paid a current pending order")
end if

' just for orders changed to PAID
if pOrderStatus=4 then
    call paymentApproved("BackOffice Lite Change Status", pOrderPrefix&pIdOrder, "Authorized") 
else
     mySQL = "UPDATE orders SET orderstatus=" & pOrderstatus & " WHERE idOrder=" & pIdOrder
     call updateDatabase(mySQL, rsTemp, "comersus_backoffice_updateOrderStatus.asp")
end if
 
call closeDb()
 
if pOrderStatus=5 then
  response.redirect "comersus_backoffice_chargeback.asp?idOrder="&pIdOrder
end if
 
response.redirect "comersus_backoffice_message.asp?message="&Server.UrlEncode("Order status updated.")

%>

