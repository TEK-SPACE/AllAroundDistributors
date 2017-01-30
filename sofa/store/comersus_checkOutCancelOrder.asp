<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: delete current cart items
%>
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/cartFunctions.asp"-->  
<!--#include file="../includes/sessionFunctions.asp"--> 

<%
on error resume next

dim mySql, connTemp, rsTemp, rsTemp2

pIdDbSessionCart2     	= getSessionVariable("idDbSessionCart",0)

if sessionLost() then
 ' nothing to delete since session was lost
 response.redirect "comersus_index.asp"
end if

' iterate through cart rows

mySQL="SELECT idCartRow FROM cartRows WHERE idDbSessionCart="&pIdDbSessionCart2

call getFromDatabase(mySQL, rstemp, "cancelOrder") 

do while not rstemp.eof

 ' clear cartRowsOptions

 mySQL="DELETE FROM cartRowsOptions WHERE idCartRow="&rstemp("idCartRow")

 call updateDatabase(mySQL, rstemp2, "cancelOrder") 
  
 rstemp.movenext
loop

' delete cartRows

mySQL="DELETE FROM cartRows WHERE idDbSessionCart="&pIdDbSessionCart2

call updateDatabase(mySQL, rstemp2, "cancelOrder") 

call closeDb()
  
' clear header cart vars
session("cartSubTotal") = 0
session("cartItems")	= 0

' redirect to store home
response.redirect "comersus_index.asp"
%>