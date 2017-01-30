<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: advanced search exec process
%>
<!--#include file="../includes/screenMessages.asp"-->
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/stringFunctions.asp"--> 
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/customerTracking.asp"--> 

<% 

on error resume next

dim connTemp, rsTemp2, f, pCounter

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanyLogo	 	= getSettingKey("pCompanyLogo")
pAllowNewCustomer	= getSettingKey("pAllowNewCustomer")
pCompareProducts 	= getSettingKey("pCompareProducts")

const pNumPerPage 	= 6

dim curPage
pCurPage		= getUserInput(request("curPage"),10)

if request.queryString("curPage") = "" then
   pCurPage = 1 
end if

pStrSearch 	= formatForDb(getUserInput(request("keyWord"),100))
pStrSearch	= replace(pStrSearch,"'","")
pStrSearch	= replace(pStrSearch,"""","")
pClearance 	= getUserInput(request("clearance"),2)
pIdCategory 	= getUserInput(request("idCategory"),4)
pIdSupplier	= getUserInput(request("idSupplier"),4)
pWithStock	= getUserInput(request("withStock"),2)

pIdCustomer     = getSessionVariable("idCustomer",0)
pIdCustomerType = getSessionVariable("idCustomerType",1)


call customerTracking("comersus_advancedSearchExec.asp", request.form)

' get items
mySql = "SELECT DISTINCT products.idProduct, sku, description, listPrice, smallImageUrl, isBundleMain, rental, visits, stock.stock FROM products, categories_products, stock WHERE products.idProduct=categories_products.idProduct AND products.idStock=stock.idStock AND listHidden=0 AND active=-1 AND idStore=" &pIdStore  

if pIdSupplier<>"0" then
   mySql = mySql & " AND idSupplier=" &pIdSupplier
end if

if pIdCategory<>"0" then
   mySql = mySql & " AND idCategory=" &pIdCategory
end if

if pWithStock="-1" then
   mySql = mySql & " AND stock.stock>0" 
end if

if pClearance="-1" then
   mySql = mySql & " AND hotDeal=-1" 
end if

if pStrSearch<>"" then
   mySql = mySql & " AND description LIKE '%" &pStrSearch& "%'"
end if
   
call getFromDatabasePerPage(mySql, rsTemp2,"advancedSearchExec")

if rsTemp2.eof then 
   response.redirect "comersus_message.asp?message="&Server.Urlencode("No items under your search")            
end if

rsTemp2.moveFirst
rsTemp2.pageSize = pNumPerPage

' get the max number of pages
dim TotalPages
totalPages = rsTemp2.PageCount

' set the absolute page
rsTemp2.absolutePage = pCurPage
   
%>

<!--#include file="header.asp"--> 
<br><b><%=getMsg(105,"Ad Search")%></b><br><br>
<%=getMsg(106,"Creiteria")%>: 

<%if pStrSearch<>"" then%>
<%=getMsg(107,"Kword")%>: <%=pStrSearch%>
<%end if%>

<%if pClearance="-1" then%>
 <%=getMsg(108,"Clearance")%>
<%end if%>

<%if pWithStock="-1" then%>
 <%=getMsg(109,"W Stock")%>
<%end if%>

<%if pIdSupplier<>"0" then%>
 <%=getMsg(110,"Supplier")%> <%=getSupplierName(pIdSupplier)%>
<%end if%>

<br><br>


<%
' set count equal to zero
pCounter = 0
do while not rsTemp2.eof and pCounter < rsTemp2.pageSize
   
   pIdProduct		= rsTemp2("idProduct")
   pDescription		= rsTemp2("description")
   pSku			= rsTemp2("sku")   
   pListPrice		= rsTemp2("listPrice")   
   pSmallImageUrl 	= rsTemp2("smallImageUrl")
   pStock		= rsTemp2("stock")
   pIsBundleMain	= rsTemp2("isBundleMain")
   pRental		= rsTemp2("rental")        
   pVisits		= rsTemp2("visits")        
   
   pPrice 		= getPrice(pIdProduct, pIdCustomerType, pIdCustomer)

%>
<b>(<%=pSku%>) <%=pDescription%></b>

<a href='comersus_viewItem.asp?idProduct=<%=pIdProduct%>'><%=getMsg(13,"View")%></a> 

<%if pCompareProducts="-1" then%>
 <a href='comersus_optCompareProductsAdd.asp?idProduct=<%=pIdProduct%>'><%=getMsg(14,"Compare")%></a>
<%end if%>

<br><%=getMsg(15,"Price")%>&nbsp; <%=pCurrencySign & money(pPrice)%> 

<%if (pListPrice-pPrice)>0 then%> 
 <%=getMsg(16,"Saving")%>&nbsp;<%=pCurrencySign & money(pListPrice-pPrice)%> 
<%end if%> 

- <%=getMsg(60,"Visits")%>: <%=pVisits%> 

<%if pShowStockView="-1" then%>
 - <%=getMsg(61,"Stock")%>: <%=pStock%>
<%end if%>
  
<div align=center>
<%if pSmallImageUrl<>"" then%>
  <a href='comersus_viewItem.asp?idProduct=<%=pIdProduct%>'><img src="catalog/<%=pSmallImageUrl%>" border="0"></a>
<%else%>
  <img src="catalog/imageNa_sm.gif" >
<%end if%>
</div>

<%
   pCounter = pCounter + 1
   rsTemp2.MoveNext
loop
%> 

<br>
<form method="post" action="" name="nscapeview">
<%="Page" & pCurPage & " of "& TotalPages%>
<%
if pCurPage > 1 then
   response.write("<INPUT TYPE=BUTTON VALUE='<' ONCLICK=""document.location.href='comersus_advancedSearchExec.asp?keyword=" &pStrSearch& "&curPage="& pCurPage - 1 & "&clearance=" &pClearance& "&idCategory=" &pIdCategory& "&IdSupplier=" &pIdSupplier& "&withStock=" &pWithStock& "';"">")
end if

if cInt(pCurPage) <> cInt(totalPages) then
   response.write("<INPUT TYPE=BUTTON VALUE='>' ONCLICK=""document.location.href='comersus_advancedSearchExec.asp?keyword=" &pstrSearch& "&curPage="& pCurPage + 1 & "&clearance=" &pClearance& "&idCategory=" &pIdCategory& "&IdSupplier=" &pIdSupplier& "&withStock=" &pWithStock&  "';"">")
end if
%>
</form>
<!--#include file="footer.asp"--> 
<%call closeDb()%>

