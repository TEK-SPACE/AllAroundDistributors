<%
' Comersus Sophisticated Cart 
' Developed by Comersus Open Technologies LC
' Software License can be found at License.txt
' http://www.comersus.com 
' Details: redirects to default index page inside the store
%>
<!--#include file="includes/miscFunctions.asp"--> 
<%
call saveCookie()

pIdAffiliate=request.querystring("idAffiliate")

if isNumeric(pIdAffiliate) then 
 response.redirect "store/default.asp?idAffiliate="&pIdAffiliate
else
 response.redirect "store/default.asp"
end if

%>




