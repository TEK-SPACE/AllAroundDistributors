<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: logout customer
%>
<!--#include file="../includes/settings.asp"--> 
<%
' clear session idCustomer
session("idCustomer")		=Cint(0) 
session("idCustomerType")	=Cint(1)
session("customerName")		=""

response.redirect "comersus_index.asp"
%>
