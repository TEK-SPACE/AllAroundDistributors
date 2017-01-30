<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"-->  
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/stringFunctions.asp"-->    
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->   
<!--#include file="../includes/currencyFormat.asp"--> 

<% 
on error resume next 

dim mySQL, conntemp, rstemp

' get settings 
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pOrderPrefix		= getSettingKey("pOrderPrefix")

pIdCustomer	        = getUserInput(request("idCustomer"),20)

' sum order total
pOrderTotals=0

mySQL="SELECT SUM(total) AS orderTotal FROM orders"
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_listorders.asp") 

if not rstemp.eof then 
 pOrderTotals=rstemp("orderTotal")
end if

rstemp.close

' sum order total for paid and delivered orders
pOrderTotalsPaidDel=0

mySQL="SELECT SUM(total) AS orderTotal FROM orders WHERE orderStatus=4 or orderStatus=2"
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_listorders.asp") 

if not rstemp.eof then
 if not isnull(rstemp("orderTotal")) then
	pOrderTotalsPaidDel=rstemp("orderTotal")			
 end if
end if
	
	
if pIdCustomer = "" then
    ' get orders
    mySQL="SELECT orders.idOrder, name, lastName, phone, email, orderDate, total, orderStatus FROM orders, customers WHERE orders.idCustomer=customers.idCustomer ORDER BY idOrder DESC"
    call getFromDatabase(mySQL, rstemp, "comersus_backoffice_listorders.asp") 
else        
    ' get orders for one customer
    mySQL="SELECT orders.idorder, name, lastName, phone, email, orderDate, total, orderStatus FROM orders, customers WHERE customers.idCustomer=" &pIdCustomer&" AND orders.idCustomer=customers.idCustomer ORDER BY idOrder DESC"
    call getFromDatabase(mySQL, rstemp, "comersus_backoffice_listorders.asp") 
end if

if  rstemp.eof then 
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("No orders found")
end if

pCounter=1

%>

<!--#include file="header.asp"-->
<!--#include file="./includes/settings.asp"-->

<br><b>Orders</b><br><br>

<br>All time sales: <%=pCurrencySign & money(pOrderTotals)%>
<br>All time delivered and paid:  <%=pCurrencySign & money(pOrderTotalsPaidDel)%>

<%if cdbl(pOrderTotalsPaidDel) > 1000 then%>
<%
 pPercentage=money(299*100/Cdbl(pOrderTotalsPaidDel))
%>
 <br><br><i>Support Comersus development: with only <%=pPercentage%>% of your sales you can get the <a href='http://www.comersus.com/power-pack.html' target='_blank'>Power Pack</a>. 
 <br>Includes complete control panel, real time shipping, anti-fraud, Youtube integration and more. 
 <br><br><a href='http://www.comersus.com/power-pack.html' target='_blank'><img src="images/getPowerPackNow.gif" border=0 alt="Get Power Pack"></a></i><br><br>
<%end if%>

<table bgcolor="#CCCCCC" cellpadding="1" cellspacing="1" border="0" width="650">

  <tr> 
    <td>Order</td>
    <td>Date</td>
    <td>Customer</td>
    <td>Amount</td>
    <td>Status</td>
    <td>View</td>
    
  </tr>
  <%
do  while not rstemp.eof And pCounter < 25

   ' contact
   pName		= rstemp("name") & " " & rstemp("lastName")

   ' order
   pIdorder		= rstemp("idOrder")
   pOrderDate		= rstemp("orderDate")
   pOrderStatus		= rstemp("orderStatus")   
   pTotal		= rstemp("total")   
%> 
  <tr> 
    <td bgcolor="#FFFFFF"><%=pOrderPrefix&pIdOrder%></td>
    <td bgcolor="#FFFFFF"><%=pOrderDate%></td>
    <td bgcolor="#FFFFFF"><%=pName%></td>
    <td bgcolor="#FFFFFF"><%=pCurrencySign & money(ptotal)%></td>
    <td bgcolor="#FFFFFF"> <%
    select case pOrderStatus
	case 1
   		response.write "Pending"
	case 2
   		response.write "Delivered"
	case 3
   		response.write "Cancelled"
   	case 4
   		response.write "Paid"   		
	case 5
   		response.write "Chargeback"
   	case 6
   		response.write "Refunded"       
    end select
    %></td>
    <td bgcolor="#FFFFFF"> 
      <a href="comersus_backoffice_showOrder.asp?idOrder=<%=pIdOrder%>">View</a>
    </td>
    
  </tr>
  <%
 pCounter=pCounter+1 
 rstemp.movenext 
loop
%> 
</table><br>

<%if pIdCustomer = "" then%>
 <i>Only latest 25 orders are displayed in this listing</i>
<%end if%>

<%call closeDb()%>
<!--#include file="footer.asp"-->