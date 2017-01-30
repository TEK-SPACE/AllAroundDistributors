<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: update bday
%>
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/screenMessages.asp"-->
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/cartFunctions.asp"-->
<!--#include file="../includes/dateFunctions.asp"--> 

<% 

on error resume next

dim mySQL, connTemp, rsTemp

pDateSwitch		= getSettingKey("pDateSwitch")
pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pIdCustomer     	= getSessionVariable("idCustomer",0)
pBirthday 		= getUserInput(request("birthday"),20)

' date not entered
if pBirthday="" then
  response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(169,"Enter")) 
end if

if pStoreFrontDemoMode="-1" then
  response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(188,"Option disabled for this demo."))
end if

' update
mySQL="UPDATE customers SET birthday='"&formatDate(pBirthday)&"' WHERE idCustomer=" &pIdCustomer

call updateDatabase(mySQL, rstemp, "birth Exec")

response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(170,"Udated")) 

call closeDB()
%>
