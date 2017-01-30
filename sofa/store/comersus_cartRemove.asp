<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: remove and item from cart
%>
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/screenMessages.asp"-->
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/cartFunctions.asp"-->
<!--#include file="../includes/customerTracking.asp"-->  
<!--#include file="../includes/getSettingKey.asp"--> 

<% 

on error resume next

dim mySQL, connTemp, rsTemp, rsTemp2

pIdCartRow 		= getUserInput(request("idCartRow"),10)
pIdDbSessionCart	= getUserInput(request("idDbSessionCart"),10)

' cartRow was not specified
if pIdCartRow="" then
  response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(39,"Enter an item")) 
end if

call customerTracking("comersus_cartRemove.asp", request.querystring)

' get the idProduct
mySQL="SELECT idProduct FROM cartRows WHERE idCartRow=" &pIdCartRow

call getFromDatabase(mySQL, rstemp, "cartRemove")

if not rstemp.eof then 
 pIdProduct=rstemp("idProduct")
else
 pIdProduct=0
end if

' remove from cartRowOptions
mySQL="DELETE FROM cartRowsOptions WHERE idCartRow=" &pIdCartRow

call updateDatabase(mySQL, rstemp, "cartRemove")

' remove from cartRow
mySQL="DELETE FROM cartRows WHERE idCartRow=" &pIdCartRow

call updateDatabase(mySQL, rstemp, "cartRemove")

call closeDB()

response.redirect "comersus_goToShowCart.asp"
%>
