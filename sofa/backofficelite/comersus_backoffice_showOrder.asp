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

mySQL="SELECT orders.idOrder, name, lastName, orderstatus, phone, email, viewed, orders.address AS address, orders.state AS state, orders.stateCode as stateCode, orders.zip AS zip, orders.city AS city, orders.countryCode AS countryCode, taxAmount, orderDate, ShipmentDetails, paymentDetails, discountDetails, obs, total, details, orders.shippingAddress, orders.shippingCity, orders.shippingState, orders.shippingStateCode, orders.shippingZip, orders.shippingCountryCode FROM orders, customers WHERE orders.idCustomer=customers.idCustomer AND orders.idOrder=" &pIdOrder

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

' change viewed mark
if pViewed=0 and pBackOfficeDemoMode=0 then

 mySQL="UPDATE orders SET viewed=-1 WHERE idOrder=" &pIdOrder
 
 call updateDatabase(mySql, rsTemp, "showOrder")

end if

pPaymentDetails=replace(pPaymentDetails,"$0.00","")

%> 

<!--#include file="header.asp"-->
<!--#include file="./includes/settings.asp"-->   

<br><b>Order details</b><br>
<table width="595" border="0">
  <tr bgcolor="#CCCCCC"> 
    <td height="14" colspan="3"><img src="images/smallIcoBackOffice2.gif" width="11" height="11"> 
      <font color="#FFFFFF"><b><%= pOrderPrefix&pIdOrder%></b></font>, Date <%=pOrderDate%> </td>
  </tr>
  <tr> 
    <td colspan="2" height="20">Name</td>
    <td><%=pName&" "&pLastName%></td>
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
    <td colspan="2" height="20">Details</td>
    <td><%=pDetails%></td>
  </tr>
  <tr> 
    <td colspan="2" height="20">Shipment</td>
    <td><%=pShipmentDetails%></td>
  </tr>
  <tr> 
    <td colspan="2" height="20">Payment</td>
    <td><%=pPaymentDetails%></td>
  </tr>
  <tr> 
    <td colspan="2" height="20">Discount</td>
    <td><%=pDiscountDetails%></td>
  </tr>
  <tr>
    <td colspan="2" valign="top" height="20">Status</td>
    <td> 
    <form method="post" action="comersus_backoffice_updateOrderStatus.asp">
    
    <input type="hidden" name="idOrder" value="<%=pIdOrder%>">
    <input type="hidden" name="oldOrderStatus" value="<%=pOrderStatus%>">
    
      <select name="orderStatus">
        <option value="1" 
        <%if pOrderStatus=1 then
        response.write "selected"
        end if%>
        >Pending</option>
                
        <option value="4"
        <%if pOrderStatus=4 then
        response.write "selected"
        end if%>>Paid</option>
        
        <option value="2"
        <%if pOrderStatus=2 then
        response.write "selected"
        end if%>>Delivered</option>
        
        <option value="3"
        <%if pOrderStatus=3 then
        response.write "selected"
        end if%>>Cancelled</option>     
        
        <option value="5"
        <%if pOrderStatus=5 then
        response.write "selected"
        end if%>>Chargeback</option>     
           
      </select>
        <input alt="Change status" border=0 height=15 name=update src="images/yellow_arrow.gif" type=image width=18>        
        <br>
        <%if pOrderStatus=4 then%>
         <br><a href="comersus_backoffice_rollbackorder.asp?idOrder=<%=pIdOrder%>">Roll back</a> <i>Will cancel the order and update sales and stock.</i>
        <%end if%>
        <br><a href="comersus_backoffice_deleteOrderExec.asp?idOrder=<%= pIdOrder%>">Delete</a> <i>Warning: You cannot undo this action!</i>
      </form>      
    </td>
  </tr>
<%if pUseComersusOLPayment=-1 then%>
  <tr> 
    <td colspan="2">Credit Card</td>
    <td>         
     <%if pUseSslOLPayment=-1 then%>
     <a href='<%="https://"&pStoreLocation&"/backofficeLite/"%>comersus_backoffice_showCreditCardData.asp?idOrder=<%=pIdOrder%>'>
     View credit card</a> 
    <%else%>    
     <a href='comersus_backoffice_showCreditCardData.asp?idOrder=<%=pIdOrder%>'>
      View credit card</a> <i>SSL is disabled in this installation</i>
    <%end if%>              
      </td>
  </tr>
<%end if%>
  <tr> 
    <td height="21" colspan="2">Tax</td>
    <td><%= pCurrencySign & money(pTaxAmount)%></td>
  </tr>
  <tr> 
    <td height="16" colspan="3"><hr> 
      <div align="right"><b>Total <%=pCurrencySign & money(pTotal)%></b></div>
    </td>
  </tr>
</table>
<%call closeDb()%> 
<!--#include file="footer.asp"-->