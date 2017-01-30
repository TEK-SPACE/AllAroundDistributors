<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: cancel order
%>
<!--#include file="comersus_customerLoggedVerify.asp"-->
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/screenMessages.asp"-->
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/sendMail.asp"--> 
<!--#include file="../includes/cartFunctions.asp"-->

<% 

on error resume next

dim mySQL, connTemp, rsTemp, rsTemp2

pCompany		= getSettingKey("pCompany")
pOrderPrefix		= getSettingKey("pOrderPrefix")
pEmailSender		= getSettingKey("pEmailSender")
pEmailAdmin		= getSettingKey("pEmailAdmin")
pSmtpServer		= getSettingKey("pSmtpServer")
pEmailComponent		= getSettingKey("pEmailComponent")
pDebugEmail		= getSettingKey("pDebugEmail")

pIdOrder 		= getUserInput(request("idOrder"),12)
pIdCustomer		= getSessionVariable("idCustomer", 0)


' extract real idorder (without prefix)
pIdOrder 	= removePrefix(pIdOrder,pOrderPrefix)

if pIdOrder="" then 
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(586,"Invalid order #")) 
end if

' check if the order belongs to that customer
mySQL="SELECT idCustomer, orderStatus FROM orders WHERE idOrder=" &pIdOrder

call getFromDatabase(mySQL, rstemp, "cancelOrder")

pIdCustomer2	=rstemp("idCustomer")
pOrderStatus	=rstemp("orderStatus")

if pIdCustomer<>pIdCustomer2 then 
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(588,"The order is not inside your history"))
end if

' check if the order is pending

if pOrderStatus<>1 then 
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(589,"The order is not pending")) 
end if

' cancel
mySQL="UPDATE orders SET orderStatus=3 WHERE idOrder=" &pIdOrder
call updateDatabase(mySQL, rstemp, "cancelOrder")

' send email notification
call sendmail (pCompany, pEmailSender, pEmailAdmin, "Cancellation request, order: #"&pOrderPrefix&pIdorder, "The customer has requested an order cancellation"&vbcrlf)

response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(590,"Thanks, your cancellation request has been saved.")) 

call closeDB()

%>
