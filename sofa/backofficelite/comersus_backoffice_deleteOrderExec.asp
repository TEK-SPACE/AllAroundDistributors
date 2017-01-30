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

<% 
on error resume next

dim mySQL, conntemp, rstemp, rstemp2

if pBackOfficeDemoMode=-1 then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Function disabled in demo mode")
end if

' form parameter 
pIdOrder	= getUserInput(request.QueryString("idOrder"),12)

' get dbSessionCart related
mySQL="SELECT idDbSessionCart FROM dbSessionCart WHERE idOrder="&pIdOrder
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_deleteOrderExec.asp") 

if not rstemp.eof then
 pIdDbSessionCart=rstemp("idDbSessionCart")
 
 ' iterate through cartRows
 mySQL="SELECT idCartRow FROM cartRows WHERE idDbSessionCart="& pIdDbSessionCart
 call getFromDatabase(mySQL, rstemp, "comersus_backoffice_deleteOrderExec.asp") 

 do while not rstemp.eof
  
   ' delete all cartRowsOptions
   mySQL="DELETE FROM cartRowsOptions WHERE idCartRow="& rstemp("idCartRow")
   call updateDatabase(mySQL, rstemp2, "comersus_backoffice_deleteOrderExec.asp") 
  
   rstemp.movenext
 loop
 
 ' delete all cartRows
  mySQL="DELETE FROM cartRows WHERE idDbSessionCart="& pIdDbSessionCart
  call updateDatabase(mySQL, rstemp2, "comersus_backoffice_deleteOrderExec.asp") 

  ' delete dbSessionCart
  mySQL="DELETE FROM dbSessionCart WHERE idDbSessionCart="&pIdDbSessionCart
  call updateDatabase(mySQL, rstemp, "comersus_backoffice_deleteOrderExec.asp") 
 
end if '  has dbSessionCart related

' credit cards
mySQL="DELETE FROM creditCards WHERE idOrder="&pIdOrder
call updateDatabase(mySQL, rstemp, "comersus_backoffice_deleteOrderExec.asp") 

' orders
mySQL="DELETE FROM orders WHERE idOrder="&pIdOrder
call updateDatabase(mySQL, rstemp, "comersus_backoffice_deleteOrderExec.asp") 

call closeDb()

response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Order deleted")
%>
