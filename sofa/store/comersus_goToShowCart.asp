<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com 
' Details: calculate header total and redirect to shopping cart
%>
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/cartFunctions.asp"--> 
<!--#include file="../includes/discountFunctions.asp"--> 
<!--#include file="../includes/screenMessages.asp"--> 
<%
on error resume next

dim mySQL, connTemp, rsTemp

pIdDbSession	 = checkSessionData()
pIdDbSessionCart = checkDbSessionCartOpen()

' update Session Values for Header Cart

session("cartItems")	= calculateCartQuantity(pIdDbSessionCart)

pDiscountDesc	=""
pDiscountTotal	=0
pDiscountCode	= getSessionVariable("discountCode","")
call getDiscount(pDiscountCode, pDiscountDesc, pDiscountTotal) 

session("cartSubTotal") = calculateCartTotal(pIdDbSessionCart)-pDiscountTotal
 
call closeDb()

response.redirect "comersus_showCart.asp"

%>