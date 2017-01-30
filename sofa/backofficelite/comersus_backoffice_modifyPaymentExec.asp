<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"-->   
<!--#include file="comersus_backoffice_functions.asp"-->   
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/settings.asp" --> 
<!--#include file="../includes/stringFunctions.asp" --> 

<% 
on error resume next 

dim mySQL, conntemp, rstemp

pPaymentDesc		= getUserInput(request.form("paymentDesc"),100)
pPriceToAdd		= getUserInput(request.form("priceToAdd"),20)
pPercentageToAdd	= getUserInput(request.form("percentageToAdd"),20)
pRedirectionUrl		= getUserInput(request.form("redirectionUrl"),250)
pEmailText		= getUserInput(request.form("emailText"),250)

pQuantityFrom		= getUserInput(request.form("quantityFrom"),20)
pQuantityUntil		= getUserInput(request.form("quantityUntil"),20)

pWeightFrom			= 0
pWeightUntil		= 999999

pPriceFrom		= getUserInput(request.form("priceFrom"),20)
pPriceUntil		= getUserInput(request.form("priceUntil"),20)

pIdCustomerType		= getUserInput(request.form("customerType"),8)

pIdPayment		= getUserInput(request.form("idPayment"),8)


pPaymentDesc		= formatForDb(pPaymentDesc)
pRedirectionUrl		= formatForDb(pRedirectionUrl)

if trim(pIdPayment)="" then
 response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Invalid payment id.")
end if

if pPaymentDesc="" then
 response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Please enter a valid payment description")
end if

' select insert sentence 
mySQL="UPDATE payments SET paymentDesc = '" &pPaymentDesc& "', priceToAdd=" &pPriceToAdd& ", percentageToAdd = " &pPercentageToAdd& ", quantityFrom =" &pQuantityFrom& ", quantityUntil =" &pQuantityUntil& ", priceFrom =" &pPriceFrom& ", priceUntil =" &pPriceUntil& ", redirectionUrl='" &pRedirectionUrl& "', idCustomerType=" &pIdCustomerType& ", emailText= '"&pEmailText&"' WHERE idPayment=" & pIdPayment

call updateDatabase(mySQL, rstemp, "comersus_backoffice_addPaymentExec.asp") 

call closeDb()

response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Payment modified.")
%>
