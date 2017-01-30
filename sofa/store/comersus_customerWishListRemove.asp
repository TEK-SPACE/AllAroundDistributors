<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: remove an item from wish list
%>
<!--#include file="comersus_customerLoggedVerify.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/databaseFunctions.asp"-->
<!--#include file="../includes/screenMessages.asp" --> 
<!--#include file="../includes/currencyFormat.asp" --> 
<!--#include file="../includes/stringFunctions.asp" --> 
<!--#include file="../includes/sessionFunctions.asp" --> 

<%
on error resume next

dim connTemp, rsTemp, mySql

pIdCustomer	= getSessionVariable("idCustomer",0)
pIdProduct	= getUserInput(request.querystring("idProduct"),10)

' check if that item exists
mySql="SELECT idProduct FROM wishList WHERE idcustomer=" &pIdCustomer& " AND idProduct=" &pIdProduct

call getFromDatabase(mySql, rsTemp, "customerWishListRemove")

if rstemp.eof then
  response.redirect "comersus_message.asp?message="& Server.Urlencode(getMsg(205,"Invalid item to delete"))
end if

mySql="DELETE FROM wishList WHERE idCustomer=" &pIdCustomer& " AND idproduct=" &pIdProduct

call updateDatabase(mySQL, rsTemp, "customerWishListRemove")

response.redirect "comersus_customerWishListView.asp"

call closeDb()
%>
