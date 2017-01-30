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
<!--#include file="./includes/settings.asp"-->

<% 
on error resume next 

if pBackOfficeDemoMode=-1 then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Function disabled in demo mode.")
end if

dim mySQL, conntemp, rstemp, rstemp2

if request("cancel")="Cancel" then
 response.redirect "comersus_backoffice_menuProducts.asp"
end if

if request("sure")<>"Sure" then
 
 %><!--#include file="header.asp"-->
	<b><font size=2>Confirmation</b></font>
	<form method="post" name="deleteItem" action="comersus_backoffice_deleteItemExec.asp">
	<table width="400" border="0">
	  <tr> 
	   <td width="270">
	   	<input type="hidden" name="idProduct" value="<%=request.Querystring("idProduct")%>">	   
	   	Are you sure?
	   </td>	   
	  </tr> 
	  <tr> 
	   <td><br>	    
	    <input type="submit" name="sure" value="Sure">&nbsp;&nbsp;        
	    <input type="submit" name="cancel" value="Cancel">&nbsp;&nbsp;        
	   </td>
	  </tr>
	 </table>
	</form>
 <!--#include file="footer.asp"--><%
else

' form parameters
pIdProduct	= getUserInput(request("idProduct"),20)

if trim(pIdProduct)="" then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Please enter the product number.")
end if

' check cartows
mySQL="SELECT idCartRow FROM cartRows WHERE idProduct=" &pIdProduct
call updateDatabase(mySQL, rstemp, "comersus_backoffice_deleteitemexec.asp") 

if not rstemp.eof then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("This item was sold. You have to delete orders with this item first.")
end if

' delete product from stockmovements
mySQL="DELETE FROM stockMovements WHERE idProduct=" &pIdProduct
call updateDatabase(mySQL, rstemp, "comersus_backoffice_deleteitemexec.asp") 

' delete product from reviews
mySQL="DELETE FROM reviews WHERE idProduct=" &pIdProduct
call updateDatabase(mySQL, rstemp, "comersus_backoffice_deleteitemexec.asp") 

' delete product from relatedProducts
mySQL="DELETE FROM relatedProducts WHERE idProduct=" &pIdProduct
call updateDatabase(mySQL, rstemp, "comersus_backoffice_deleteitemexec.asp") 

mySQL="DELETE FROM relatedproducts WHERE idRelatedProduct=" &pIdProduct
call updateDatabase(mySQL, rstemp, "comersus_backoffice_deleteitemexec.asp") 

' delete from TaxPerProduct
mySQL="DELETE FROM taxPerProduct WHERE idProduct=" &pIdProduct
call updateDatabase(mySQL, rstemp, "comersus_backoffice_deleteitemexec.asp") 

' delete from categories_products
mySQL="DELETE FROM categories_products WHERE idProduct=" &pIdProduct
call updateDatabase(mySQL, rstemp, "comersus_backoffice_deleteitemexec.asp") 

' delete from auctions offers
mySQL="SELECT idAuction FROM auctions WHERE idProduct=" &pIdProduct
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_deleteitemexec.asp") 

do while not rstemp.eof
 
 mySQL="DELETE FROM auctionOffers WHERE idAuction=" &rstemp("idAuction")
 call updateDatabase(mySQL, rstemp2, "comersus_backoffice_deleteitemexec.asp") 
 rstemp.movenext
 
loop

' delete from auctions
mySQL="DELETE FROM auctions WHERE idProduct=" &pIdProduct
call updateDatabase(mySQL, rstemp, "comersus_backoffice_deleteitemexec.asp") 

' delete from special prices
mySQL="DELETE FROM customer_specialPrices WHERE idProduct=" &pIdProduct
call updateDatabase(mySQL, rstemp, "comersus_backoffice_deleteitemexec.asp") 

' delete product from options Groups
mySQL="DELETE FROM optionsGroups_products WHERE idProduct=" &pIdProduct
call updateDatabase(mySQL, rstemp, "comersus_backoffice_deleteitemexec.asp") 

' serials
mySQL="DELETE FROM serials WHERE idProduct=" &pIdProduct

call updateDatabase(mySQL, rstemp, "comersus_backoffice_deleteitemexec.asp")

' delete product from backorder table
mySQL="DELETE FROM backOrder WHERE idProduct=" &pIdProduct
call updateDatabase(mySQL, rstemp, "comersus_backoffice_deleteitemexec.asp") 

' delete product from products table
mySQL="DELETE FROM products WHERE idProduct=" &pIdProduct
call updateDatabase(mySQL, rstemp, "comersus_backoffice_deleteitemexec.asp") 

' stock
mySQL="DELETE FROM stock WHERE idProductMain=" &pIdProduct
call updateDatabase(mySQL, rstemp, "comersus_backoffice_deleteitemexec.asp") 

response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Product deleted.")
  
end if
%>
