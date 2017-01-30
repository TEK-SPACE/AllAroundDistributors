<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"-->   
<!--#include file="../includes/sessionFunctions.asp"-->   
<!--#include file="../includes/getSettingKey.asp"-->  
<!--#include file="../includes/currencyFormat.asp"-->    
<!--#include file="../includes/settings.asp"-->

<% 
on error resume next 

dim mySQL, conntemp, rstemp, rstemp2

' get settings 
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")

dim arrCommissions(100,3)

' get commissions for paid or delivered orders

mySQL="SELECT affiliates.idAffiliate, affiliateName, commission FROM affiliates WHERE affiliates.idAffiliate>1 ORDER by affiliates.idAffiliate"

call getFromDatabase(mySQL, rstemp, "comersus_backoffice_listAffiliatesCommissions.asp") 

if rstemp.eof then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("There are no affiliates in the database")
end If

%>
<!--#include file="header.asp"-->
<!--#include file="includes/settings.asp"-->
<br><b>Affiliates commission</b><br>

<%
pNoInfoAvailable=-1

do while not rstemp.eof 
 
 mySQL="SELECT SUM(total) as pSumTotal FROM orders WHERE idAffiliate="&rstemp("idAffiliate") & " AND (orderStatus=4 OR orderStatus=2)"

 call getFromDatabase(mySQL, rstemp2, "comersus_backoffice_listAffiliatesCommissions.asp") 
 
 if rstemp2("pSumTotal")>0 then
  pNoInfoAvailable=0
 %><br>-<%=rstemp("affiliateName")%> total <%=pCurrencySign%> <%=money(rstemp2("pSumTotal"))%> commission <%=pCurrencySign%> <%=money(rstemp2("pSumTotal")*rstemp("commission")/100)%><%
 end if
 rstemp.moveNext
loop

if pNoInfoAvailable=-1 then
%><br>There is no information to make this listing<%
end if
%>
<%call closeDb()%> 
<!--#include file="footer.asp"-->