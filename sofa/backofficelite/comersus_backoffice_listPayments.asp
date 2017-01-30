<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"-->   
<!--#include file="../includes/databaseFunctions.asp"-->   
<!--#include file="../includes/settings.asp" -->  

<%
on error resume next

dim mySQL, conntemp, rstemp

mySQL="SELECT * FROM payments ORDER BY paymentDesc"

call getFromDatabase(mySQL, rstemp, "comersus_backoffice_listPayments.asp") 


%> 

<!--#include file="header.asp"-->
<!--#include file="./includes/settings.asp"--> 

<br><b>Payments</b><br><br>

<%if  rstemp.eof then 
  response.write "You don't have payments defined. Add one now."
else
%>
<table>
<tr bgcolor="#CCCCCC">
 <td>Description</td>
 <td>Redirection</td>
 <td>Customer</td>
 <td>Actions</td>
</tr>
<%

 do while not rstemp.eof 

	pPaymentDesc		= rstemp("PaymentDesc")
	pIdPayment		= rstemp("idPayment")	
	pRedirectionUrl		= rstemp("redirectionUrl")
	pIdCustomerType		= rstemp("idCustomerType")   
   %>
 <tr>
 <td><%=pPaymentDesc%></td>
 <td><%=pRedirectionUrl%></td>
 <td><%
 if pIdCustomerType=1 then
  response.write "Retail"
 end if
 if pIdCustomerType=2 then
  response.write "Wholesale"
 end if
 if isNull(pIdCustomerType) then
  response.write "All"
 end if
 %></td>
 <td><a href="comersus_backoffice_modifyPaymentForm.asp?idPayment=<%=pIdPayment%>">Modify</a>&nbsp; - &nbsp;<a href="comersus_backoffice_deletePaymentExec.asp?idPayment=<%=pIdPayment%>">Delete</a></td>
</tr>      
  <%   
   rstemp.MoveNext
 loop
%>  
</table>
<%end if%>
<br><br>
>> <a href="comersus_backoffice_addPaymentForm.asp">Click to add a payment method</a>
<br><br>Want to accept Google Checkout and PayPal Express? <a href="http://www.comersus.com/power-pack.html" target="_blank">Use the Power Pack...</a>
<br>
<!--#include file="footer.asp"--> 
<%call closeDb()%>
