<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: verifies if admin is logged, case not send to login page
%>
<%
pRedirect	=getUserInputL(request("redirect"),50)
pIdProduct	=getUserInput(request("idProduct"),20)

if Session("idCustomer")=0 then
 response.redirect "comersus_customerAuthenticateForm.asp?redirect="&pRedirect&"&idProduct="&pIdProduct
end if
%>