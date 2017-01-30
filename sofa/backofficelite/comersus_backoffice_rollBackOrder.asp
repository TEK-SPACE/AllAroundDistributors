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
<!--#include file="../includes/itemFunctions.asp"-->    
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/settings.asp"--> 

<% 
on error resume next 

dim f, mySQL, conntemp, rstemp, rstemp2

pIdOrder 	= getUserInput(request.querystring("idOrder"),12)

if trim(pIdOrder)="" then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Enter a valid order")
end if

' check if the order status is paid
mySQL="SELECT orderStatus FROM orders WHERE idOrder=" &pIdOrder

call getFromDatabase(mySQL, rstemp, "comersus_backoffice_rollBackOrder.asp") 

if rstemp("orderStatus")<>4 AND rstemp("orderStatus")<>2 then 
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Order is not paid")
end if

' iterates through items
mySQL="SELECT idProduct, quantity FROM cartRows, dbSessionCart WHERE dbSessionCart.idDbSessionCart=cartRows.idDbSessionCart AND dbSessionCart.idOrder=" &pIdOrder

call getFromDatabase(mySQL, rstemp, "comersus_backoffice_rollBackOrder.asp") 

do while not rstemp.eof

   pIdProduct	=rstemp("idProduct") 
   pQuantity	=rstemp("quantity") 
   
   ' rollback sales 
   mySQL="UPDATE products SET sales=sales-" &pQuantity& " WHERE idProduct=" &pIdProduct           
   call updateDatabase(mySQL, rstemp2, "comersus_backoffice_rollBackOrder.asp") 
   
   ' rollback stock
   
   call updateStock(pIdProduct, pQuantity)         

   rstemp.movenext
loop

' cancel order
mySQL="UPDATE orders set orderStatus=3 WHERE idOrder=" &pIdOrder

call updateDatabase(mySQL, rstemp, "comersus_backoffice_rollBackOrder.asp") 

call closeDb()

response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Order rolled back")
%>
