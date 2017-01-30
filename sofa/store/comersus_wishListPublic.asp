<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: show wish list to a third-person
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/stringFunctions.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"-->
<!--#include file="../includes/miscFunctions.asp"-->
<!--#include file="../includes/screenMessages.asp" --> 
<!--#include file="../includes/currencyFormat.asp" --> 
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/adSenseFunctions.asp"-->

<%
on error resume next

dim connTemp, mySql, rsTemp

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanyLogo	 	= getSettingKey("pCompanyLogo")
pStoreLocation 		= getSettingKey("pStoreLocation")
pMoneyDontRound	 	= getSettingKey("pMoneyDontRound")
pAuctions		= getSettingKey("pAuctions")
pListBestSellers 	= getSettingKey("pListBestSellers")
pNewsLetter 		= getSettingKey("pNewsLetter")
pPriceList 		= getSettingKey("pPriceList")
pStoreNews 		= getSettingKey("pStoreNews")
pOneStepCheckout	= getSettingKey("pOneStepCheckout")

call saveCookie()

' retrieve from querystring
pIdCustomer    	 	= getUserInput(request.querystring("idCustomer"),4)
pEmail		 	= getUserInput(request.querystring("email"),30)

' authenticate

mySql="SELECT name, lastName, idCustomerType FROM customers WHERE idCustomer="&pIdCustomer& " AND email='" &pEmail& "'"

call getFromDatabase(mySql, rsTemp, "customerWishListView")

if rstemp.eof then 
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(206,"Cannot retrieve"))            
end if

pName		=rstemp("name")
pLastName	=rstemp("lastName")
pIdCustomerType	=rstemp("idCustomerType")

' save into session for checkout
session("wishListIdCustomer")=pIdCustomer

' retrieve wish list

mySql="SELECT products.idProduct, description, sku, smallImageUrl FROM wishList, products WHERE products.idProduct=wishList.idProduct AND idCustomer="&pIdCustomer

call getFromDatabase(mySql, rsTemp, "customerWishListView")

if rstemp.eof then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(207,"There are no items in the wish list"))            
end if

%>
<!--#include file="header.asp"-->

<br><br>
<b><%=getMsg(208,"W L for")%>&nbsp;<%=pName&" "&pLastName%></b><br><br>
<i><%=getMsg(209,"Select from")%></i><br>
<br>
    <table width="100%" border="0">
      <tr bgcolor="#D0CC98">       
        <td bgcolor="#ffcc00"><%=getMsg(210,"Item")%></td>
        <td bgcolor="#ffcc00"><%=getMsg(211,"Desc")%></td>
        <td bgcolor="#ffcc00"><%=getMsg(212,"Price")%></td>        
      </tr>
<%

pWishListTotal=Cdbl(0)

do while not rstemp.eof 
 
 pPrice     = getPrice(rstemp("idProduct"), pIdCustomerType, pIdCustomer)
 %>
 <tr> 
  <td>
  <%if rstemp("smallImageUrl")<>"" then%>
   <a href="comersus_viewItem.asp?idProduct=<%=rstemp("idProduct")%>"><img align=left border=0 src="catalog/<%=rstemp("smallImageUrl")%>" vspace=3></a> 
  <%else%> 
   <img src='catalog/imageNA_sm.gif'> 
  <%end if%>
  </td>
  <td><a href="comersus_viewItem.asp?idProduct=<%=rstemp("idProduct")%>"><%=rstemp("description")%></a></td>
  <td><%=pCurrencySign & money(pPrice)%></td>  
 </tr> 
 <%
 pWishListTotal=pWishListTotal+pPrice
 rstemp.movenext
loop
%>
</table>
<br><%=getMsg(213,"Total")%> &nbsp;<%=pCurrencySign & money(pWishListTotal)%>
<!--#include file="footer.asp"-->
<%call closeDb()%>