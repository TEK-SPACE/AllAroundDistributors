<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
%>
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/screenMessages.asp" --> 
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/cartFunctions.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/currencyFormat.asp"-->
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/discountFunctions.asp"--> 


<%
on error resume next

dim mySQL, connTemp, rsTemp

pIdDbSession	 	= checkSessionData()
pIdDbSessionCart 	= checkDbSessionCartOpen()

' get user inputs
session("discountCode") =""

call closeDb()

response.redirect "comersus_showCart.asp"

%>
