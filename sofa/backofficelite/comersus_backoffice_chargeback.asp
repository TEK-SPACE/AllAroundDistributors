<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Developed by Rodrigo S. Alhadeff, contributions of Ricardo Fuga
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"-->    
<!--#include file="../includes/databaseFunctions.asp"-->  
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->   
<!--#include file="../includes/stringFunctions.asp"--> 
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/settings.asp"--> 

<% 
on error resume next 

dim mySQL, conntemp, rstemp

' get settings 
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pOrderPrefix		= getSettingKey("pOrderPrefix")

pIdOrder 		= getUserInput(request("idOrder"),12)

if pIdOrder="" then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Please enter a valid order")
end if

mySQL="SELECT orders.idOrder, name, lastName, orderstatus, phone, email, browserIp, viewed, orders.address AS address, orders.state AS state, orders.stateCode as stateCode, orders.zip AS zip, orders.city AS city, orders.countryCode AS countryCode, taxAmount, orderDate, ShipmentDetails, paymentDetails, discountDetails, obs, total, details, orders.shippingAddress, orders.shippingCity, orders.shippingState, orders.shippingStateCode, orders.shippingZip, orders.shippingCountryCode FROM orders, customers WHERE orders.idCustomer=customers.idCustomer AND orders.idOrder=" &pIdOrder

call getFromDatabase(mySQl, rstemp, "showOrder")

if  rstemp.eof then 
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Cannot get order details")
end if 
  
   ' contact
   pName		= rstemp("name")
   pLastName		= rstemp("lastName")
   pEmail		= rstemp("email")
   pPhone		= rstemp("phone")   
   
   ' order
   pOrderDate		= rstemp("orderDate")
   pDetails		= rstemp("details")
   pTotal		= rstemp("total")
   pPaymentDetails	= rstemp("paymentDetails")
   pShipmentDetails	= rstemp("ShipmentDetails")
   pDiscountDetails	= rstemp("discountDetails")
   pTaxAmount		= rstemp("taxAmount")
   pAddress		= rstemp("address")
   pZip			= rstemp("zip")
   pState		= rstemp("state")
   pStateCode		= rstemp("stateCode")
   pCity		= rstemp("city")
   pCountryCode		= rstemp("countryCode")
   pShippingaddress	= rstemp("shippingAddress")
   pShippingzip		= rstemp("shippingzip")
   pShippingstate	= rstemp("shippingstate")
   pShippingstateCode	= rstemp("shippingstateCode")
   pShippingcity	= rstemp("shippingcity")
   pShippingcountryCode	= rstemp("shippingCountryCode")
   pOrderStatus		= rstemp("orderstatus")
   pViewed		= rstemp("viewed")   
   pBrowserIp		= rstemp("browserIp")   

pPaymentDetails=replace(pPaymentDetails,"$0.00","")

%> 

<!--#include file="header.asp"-->
<!--#include file="./includes/settings.asp"-->   

<br><b>Chargeback details</b><br>
<br>
<table width="595" border="0">
  
  <tr> 
    <td colspan="2" height="20">Date of purchase</td>
    <td><%=pOrderDate%></td>
  </tr>  
  <tr> 
    <td colspan="2" height="20">Name</td>
    <td><%=pName&" "&pLastName%></td>
  </tr>
  <tr> 
    <td colspan="2" height="20">IP</td>
    <td><%=pBrowserIp%></td>
  </tr>
  <tr> 
    <td colspan="2" height="20">Email</td>
    <td><a href="mailto:<%=pEmail%>"><%= pEmail%></a></td>
  </tr>
  <tr> 
    <td colspan="2" height="20">Phone</td>
    <td><%=pPhone%></td>
  </tr>
  <tr> 
    <td colspan="2" height="20">Address</td>
    <td><%=pAddress & " (" &pZip& ") " &pState&pStateCode& " - " &pCity& " " &pCountryCode%> 
    </td>
  </tr>  
  <tr> 
    <td colspan="2" height="20">Shipping address</td>
    <td> 
    <%if pShippingAddress<>"" then%>
     <%=pShippingaddress &" (" &pShippingzip&") " &pShippingstate&pShippingStateCode& " - " &pShippingcity& " " & pShippingcountryCode%> 
    <%else%>
     Same as billing
    <%end if%>
    </td>
  </tr>  
   <tr> 
    <td height="21" colspan="2">Tax</td>
    <td><%= pCurrencySign & money(pTaxAmount)%></td>
  </tr>
  <tr> 
    <td height="21" colspan="2">Total amount</td>
      <td><b> <%=pCurrencySign & money(pTotal)%></b>
    </td>
  </tr>
  <tr> 
    <td height="21" colspan="3">&nbsp;</td>      
  </tr> 
<form method="get" action="http://www.chargebackprotection.org/frontend/loginForm.asp" target="_blank">  
  <tr> 
    <td height="16" colspan="3">
      <input type=submit value="Report this chargeback abuse"> Not a member of ChargebackProtection? <a href="http://www.chargebackprotection.org/frontend/registerForm.asp" target="_blank">Register for free</a>
    </td>
  </tr>
</form>  
</table>
<%call closeDb()%> 
<!--#include file="footer.asp"-->