<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: show wish list
%>
<!--#include file="comersus_customerLoggedVerify.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"-->
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/screenMessages.asp" --> 
<!--#include file="../includes/currencyFormat.asp" --> 
<!--#include file="../includes/cartFunctions.asp"-->
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
pCompanySlogan	 	= getSettingKey("pCompanySlogan")
pAboutUsLink 		= getSettingKey("pAboutUsLink")
pStoreLocation 		= getSettingKey("pStoreLocation")
pMoneyDontRound	 	= getSettingKey("pMoneyDontRound")
pHeaderKeywords 	= getSettingKey("pHeaderKeywords")

pAuctions		= getSettingKey("pAuctions")
pAllowNewCustomer	= getSettingKey("pAllowNewCustomer")
pNewsLetter 		= getSettingKey("pNewsLetter")
pStoreNews 		= getSettingKey("pStoreNews")
pSuppliersList		= getSettingKey("pSuppliersList")

pIdCustomer     	= getSessionVariable("idCustomer",0)
pIdCustomerType 	= getSessionVariable("idCustomerType",0)
pIdDbSessionCart	= getSessionVariable("idDbSessionCart",0)

' get customers email

mySql="SELECT email FROM customers WHERE idCustomer=" &pIdCustomer

call getFromDatabase(mySql, rsTemp, "customerWishListView")

if not rstemp.eof then 
 pEmail=rstemp("email")
end if

mySql="SELECT products.idProduct, sku, description, smallImageUrl FROM wishList, products WHERE products.idProduct=wishList.idproduct AND idCustomer="&pIdCustomer

call getFromDatabase(mySql, rsTemp, "customerWishListView")

if rstemp.eof then
 response.redirect "comersus_message.asp?message="& Server.Urlencode(getMsg(195,"No items."))
end if

%>
<!--#include file="header.asp"-->
<br><br><b><%=getMsg(192,"tittle")%></b><br><br>
<i><%=getMsg(193,"Info about WL1")%></i><br>
<i><%=getMsg(194,"Info about WL2")%></i><br>
<br>
    <table border="0">
      <tr bgcolor="#CCCCCC">       
        <td><%=getMsg(196,"Item")%></td>
        <td><%=getMsg(197,"Desc")%></td>
        <td><%=getMsg(198,"Price")%></td>
        <td><%=getMsg(199,"Actions")%></td>
      </tr>
<%

wishListTotal=Cdbl(0)

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
  <td><a href="comersus_viewItem.asp?idproduct=<%=rstemp("idProduct")%>"><%=rstemp("description")%></a></td>
  <td><%=pCurrencySign & money(pPrice)%></td>
  <td><a href="comersus_customerWishListRemove.asp?idProduct=<%=rstemp("idProduct")%>"><img src="images/buttonRemoveSmall.gif" border=0 alt="<%=getMsg(48,"Remove")%>"></a></td>
 </tr> 
 <%
 wishListTotal=wishListTotal+pPrice
 rstemp.movenext
loop
%>
 </table>
 <br><%=getMsg(201,"Total")%>&nbsp;<%=pCurrencySign & money(wishListTotal)%>
 <br><br><%=getMsg(202,"Link")%> &nbsp;
 <form>
 <textarea name="link" cols="45" maxlength=500 rows=6>
 <a href="http://<%=pStoreLocation%>/store/comersus_wishListPublic.asp?email=<%=pEmail%>&idCustomer=<%=pIdCustomer%>">http://<%=pStoreLocation%>/store/comersus_wishListPublic.asp?email=<%=pEmail%>&idCustomer=<%=pIdCustomer%></a>
 </textarea>
 </form>
 <br><br>
<!--#include file="footer.asp"-->
<%call closeDb()%>