<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: add one item to customer wish list
%>
<!--#include file="comersus_customerLoggedVerify.asp"-->   
<!--#include file="../includes/databaseFunctions.asp"-->
<!--#include file="../includes/screenMessages.asp" --> 
<!--#include file="../includes/stringFunctions.asp" --> 
<!--#include file="../includes/sessionFunctions.asp" --> 
<!--#include file="../includes/settings.asp"-->

<%
on error resume next

dim connTemp, rsTemp, mySql

pIdCustomer  	= getSessionVariable("idCustomer",0)
pIdProduct	= getUserInput(request.querystring("idProduct"),10)

' check if that item exists in the list already
mySql="SELECT idProduct FROM wishList WHERE idCustomer=" &pIdCustomer& " AND idProduct=" &pIdProduct

call getFromDatabase(mySql, rsTemp, "customerWishListAdd")

if not rstemp.eof then
  response.redirect "comersus_message.asp?message="& Server.Urlencode(getMsg(191,"The item was previously added to your wish list."))
end if

' add to whish list

mySql="INSERT INTO wishList (idCustomer, idProduct) VALUES (" &pIdCustomer& "," &pIdProduct& ")"

call updateDatabase(mySQL, rsTemp, "customerWishListAdd")

call closeDb()

response.redirect "comersus_customerWishListView.asp"
%>
