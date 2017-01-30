<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: list items according to filtering
%>

<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/screenMessages.asp"--> 
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/stringFunctions.asp"--> 
<!--#include file="../includes/cartFunctions.asp"--> 
<!--#include file="../includes/customerTracking.asp"-->  
<!--#include file="../includes/adSenseFunctions.asp"-->

<% 

on error resume next 

dim connTemp, rsTemp, rsTemp2
dim pTotalPages, pCounter, pNumPerPage

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanySlogan 		= getSettingKey("pCompanySlogan")
pShowStockView		= getSettingKey("pShowStockView")
pCompareProducts 	= getSettingKey("pCompareProducts")
pAllowNewCustomer 	= getSettingKey("pAllowNewCustomer")
pAboutUsLink 		= getSettingKey("pAboutUsLink")
pShowNews 		= getSettingKey("pShowNews")
pStoreLocation		= getSettingKey("pStoreLocation")
pHeaderKeywords 	= getSettingKey("pHeaderKeywords")
pShowSearchBox 		= getSettingKey("pShowSearchBox")
pAdSenseClient 		= getSettingKey("pAdSenseClient")

pAffiliatesStoreFront   = getSettingKey("pAffiliatesStoreFront")
pSuppliersList		= getSettingKey("pSuppliersList")
pAuctions		= getSettingKey("pAuctions")
pAffiliatesStoreFront   = getSettingKey("pAffiliatesStoreFront")
pRssFeedServer		= getSettingKey("pRssFeedServer")
pNewsLetter 		= getSettingKey("pNewsLetter")
pRealTimeShipping	= lcase(getSettingKey("pRealTimeShipping"))

pProductCustomField1 	= getSettingKey("pProductCustomField1")
pProductCustomField2 	= getSettingKey("pProductCustomField2")
pProductCustomField3 	= getSettingKey("pProductCustomField3")


pIdCustomer     	= getSessionVariable("idCustomer",0)
pIdCustomerType 	= getSessionVariable("idCustomerType",1)
pIdDbSessionCart	= getSessionVariable("idDbSessionCart",0)

pNumPerPage 		= 12

pCurrentPage		= getUserInput(request.querystring("currentPage"),4)
pHotDeal		= getUserInput(request.querystring("hotDeal"),2)
pLastChance		= getUserInput(request.querystring("lastChance"),2)
pIdSupplier		= getUserInput(request.querystring("idSupplier"),8)
pIdCategory		= getUserInput(request.querystring("idCategory"),8)
pOrderBy		= getUserInput(request.querystring("orderBy"),10)
pStrSearch 		= formatForDb(getUserInput(request("strSearch"),30))

pUser1			= getUserInput(request.querystring("user1"),50)
pUser2			= getUserInput(request.querystring("user2"),50)
pUser3			= getUserInput(request.querystring("user3"),50)

if pCurrentPage = "" then
  pCurrentPage = 1 
end if

call customerTracking("comersus_listOneCategory.asp", request.querystring)

' get items

mySql="SELECT DISTINCT products.idProduct, sku, description, visits, price, btoBprice, listPrice, smallImageUrl, sales, dateAdded, isBundleMain, rental, map, freeShipping, stock.stock, emailText FROM products, stock, categories_products WHERE listHidden=0 AND active=-1 AND idStore=" &pIdStore& " AND products.idStock=stock.idStock AND products.idProduct=categories_products.idProduct"  

' hot deal
if pHotDeal<>"" then
 mySQL=mySQL&" AND hotDeal=-1"
end if

' last chance
if pLastChance<>"" then
 mySQL=mySQL&" AND stock>0 AND stock<10"
end if

' supplier
if pIdSupplier<>"" then
 mySQL=mySQL&" AND idSupplier=" &pIdSupplier
end if

' category
if pIdCategory<>"" then
 mySQL=mySQL&" AND idCategory=" &pIdCategory
end if

' description search
if pStrSearch<>"" then 
 mySQL=mySQL&" AND (description LIKE '%" &pStrSearch& "%' OR sku LIKE '%" &pStrSearch& "%')"
end if

if pUser1<>"" then
 mySQL=mySQL&" AND user1='"&pUser1&"'"
end if

if pUser2<>"" then
 mySQL=mySQL&" AND user2='"&pUser2&"'"
end if

if pUser3<>"" then
 mySQL=mySQL&" AND user3='"&pUser3&"'"
end if

' order parameters

if pOrderBy="descr" then
 mySQL=mySQL&" ORDER BY description"
end if 

if pOrderBy="sku" then
 mySQL=mySQL&" ORDER BY sku"
end if 

if pOrderBy="visits" then
 mySQL=mySQL&" ORDER BY visits DESC"
end if 

if pOrderBy="sales" then
 mySQL=mySQL&" ORDER BY sales DESC"
end if 

if pOrderBy="recently" then
 mySQL=mySQL&" ORDER BY products.idProduct DESC"
end if 

if pOrderBy="price" then
 if pIdCustomerType=2 then
    mySQL=mySQL&" ORDER BY bToBprice"
  else
    mySQL=mySQL&" ORDER BY price"
 end if
end if 

if pLastChance<>"" then
 mySQL=mySQL&" ORDER BY stock"
end if

call getFromDatabasePerPage(mySql, rsTemp2, "listItems")

if rsTemp2.eof then 
  response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(62,"No items"))         
end if  

rsTemp2.moveFirst

rsTemp2.PageSize 	= pNumPerPage
pTotalPages 		= rsTemp2.PageCount
rsTemp2.absolutePage 	= pCurrentPage

' for best sellers listing, only show 10 results
if pOrderBy="sales" then
 pNumPerPage		= 12
end if


%>
<!--#include file="header.asp"-->

<br><b><%=getMsg(63,"Items listing")%></b> (<%=getMsg(64,"Search crit")%>: 

<%if pHotDeal<>"" then%>
<%=getMsg(65,"Clearance")%>
<%end if%>

<%if pLastChance<>"" then%>
<%=getMsg(147,"LC")%>
<%end if%>

<%if pOrderBy="sales" then%>
<%=getMsg(66,"Top sel")%>
<%end if%>

<%if pIdSupplier<>"" then%>
<%=getMsg(67,"Supplier list")%>
<%end if%>

<%if pIdCategory<>"" then%>
<%=getMsg(284,"Category")%>&nbsp; <a href="comersus_listOneCategory.asp?idCategory=<%=pIdCategory%>"><%=getCategoryDescription(pIdCategory)%></a>
<%end if%>

<%if pStrSearch<>"" then%>
<%=getMsg(68,"Kword")%> '<%=pStrSearch%>'
<%end if%>

<%if pUser1<>"" then%>
 <%=pProductCustomField1%> <%=pUser1%>
<%end if%>

<%if pUser2<>"" then%>
 <%=pProductCustomField2%> <%=pUser2%>
<%end if%>

<%if pUser3<>"" then%>
 <%=pProductCustomField3%> <%=pUser3%>
<%end if%>

&nbsp;<%=getMsg(69,"ordered by")%>: 
<%
select case pOrderBy
 case "descr"
  response.write getMsg(70,"desc")
 case "sku" 
  response.write getMsg(71,"sku")
 case "visits"
  response.write getMsg(72,"popul")
 case "sales"
  response.write getMsg(73,"sales")
 case "recently"
  response.write getMsg(74,"rece added")
 case ""
  response.write getMsg(75,"no order")
end select
%>
)
<br>

<table width=100%>

<%

pCounter= 0
pColumnCounter=1

do while not rsTemp2.eof and pCounter< pNumPerPage
   
   pIdProduct		= rsTemp2("idProduct")
   pSku			= rsTemp2("sku")
   pDescription		= rsTemp2("description")   
   pListPrice		= rsTemp2("listPrice")      
   pSmallImageUrl	= rsTemp2("smallImageUrl")   
   pStock		= getStock(pIdProduct)
   pIsBundleMain	= rsTemp2("isBundleMain")
   pRental		= rsTemp2("rental")
   pVisits		= rsTemp2("visits")
   pDateAdded		= rsTemp2("dateAdded")
   pMapPrice		= rstemp2("map")
   pFreeShipping	= rstemp2("freeShipping")
   pEmailText		= rstemp2("emailText")
   
   pPrice 		= getPrice(pIdProduct, pIdCustomerType, pIdCustomer)
                 
   %>      
 
<%if pColumnCounter=1 then%> 
 <tr>
<%end if%>

 <td>
 
  <br> 
  <%if pSmallImageUrl<>"" then%>
    <a href='comersus_viewItem.asp?idProduct=<%=pIdProduct%>'><img src="catalog/<%=pSmallImageUrl%>" border="0"></a>
  <%else%>
    <img src="catalog/imageNa_sm.gif" >
  <%end if%>
 <br><b><a href='comersus_viewItem.asp?idProduct=<%=pIdProduct%>'><%=pDescription%></a></b>
 <br>
 <%=getMsg(60,"visits")%>: <%=pVisits%> 
   <%if pShowStockView="-1" then%>
   <i> (<%=getMsg(3,"Stock")%>: <%=pStock%>) </i>
  <%end if%>
 <%if pMapPrice = "-1" then%>
 	<br><div class=priceSmall><%=getMsg(2,"price")%>: [<%=getMsg(745,"add to cart to find out")%>]</div> 	
 <%else%>
 	<br><div class=priceSmall><%=getMsg(2,"price")%>: <%=pCurrencySign & money(pPrice)%></div> 
 <%end if%>
  
  
  <%if pCompareProducts="-1" then%>
   <a href='comersus_optCompareProductsAdd.asp?idProduct=<%=pIdProduct%>'><img src="images/buttonCompareSmall.gif" border=0 alt="<%=getMsg(14,"Compare")%>"></a>
  <%end if%>    
  
  <%if isRental(pIdProduct)=0 then %>
   <a href="comersus_addItem.asp?idProduct=<%=pIdProduct%>&quantity=1"><img src="images/buttonAdd.gif" border=0 alt="<%=getMsg(664,"Add to cart")%>"></a>
  <%end if%>
  
  <%if pFreeShipping=-1 and (pRealTimeShipping="none" or pRealTimeShipping="ups") then%>
   <img src="images/freeShippingTruck.gif" alt="<%=getMsg(32,"Free Shipping")%>">          
  <%end if%>
 
  <%if instr(pEmailText,".zip")<>0 then%>
   <img src="images/download.gif" alt="<%=getMsg(753,"download")%>">
  <%end if%>
 
  </td>
    
 <%if pColumnCounter=4 then%>    
  </tr>
 <%end if%>
   
   <%  
   pCounter= pCounter+ 1
   
   pColumnCounter	= pColumnCounter+1
   
   if pColumnCounter=5 then 
    pColumnCounter=1
   end if
   
   rsTemp2.moveNext
loop
%>
 
<%if pColumnCounter=2 then%>
<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>
<%end if%>

<%if pColumnCounter=3 then%>
 <td>&nbsp;</td><td>&nbsp;</td></tr>
<%end if%>
 
<%if pColumnCounter=4 then%>
 <td>&nbsp;</td></tr>
<%end if%>

</table>

<form method="post" action="" name="nscapeview">

<%if pOrderBy<>"sales" then%>

 <%=getMsg(10,"page")%>&nbsp;<%=pCurrentPage%>&nbsp;<%=getMsg(11,"of")%>&nbsp;<%=pTotalPages%>

 <%if pCurrentPage > 1 then%>
  <INPUT TYPE=BUTTON VALUE='<' ONCLICK="document.location.href='comersus_listItems.asp?currentPage=<%=pCurrentPage-1%>&hotDeal=<%=pHotDeal%>&orderBy=<%=pOrderBy%>&strSearch=<%=pStrSearch%>&idSupplier=<%=pIdSupplier%>&lastChance=<%=pLastChance%>&user1=<%=pUser1%>&user2=<%=pUser2%>&user3=<%=pUser3%>&idCategory=<%=pIdCategory%>'">
 <%end if%>
 <%if CInt(pCurrentPage) <> CInt(pTotalPages) then%>
  <INPUT TYPE=BUTTON VALUE='>' ONCLICK="document.location.href='comersus_listItems.asp?currentPage=<%=pCurrentPage+1%>&hotDeal=<%=pHotDeal%>&orderBy=<%=pOrderBy%>&strSearch=<%=pStrSearch%>&idSupplier=<%=pIdSupplier%>&lastChance=<%=pLastChance%>&user1=<%=pUser1%>&user2=<%=pUser2%>&user3=<%=pUser3%>&idCategory=<%=pIdCategory%>'">
 <%end if%>
<%end if%> 

<%' show order-by only if is not last chance listing and top sellers%>
<%if pLastChance="" and pOrderBy<>"sales" then%>
 &nbsp;<%=getMsg(82,"Order by")%> <a href="comersus_listItems.asp?hotDeal=<%=pHotDeal%>&orderBy=descr&strSearch=<%=pStrSearch%>&idSupplier=<%=pIdSupplier%>&user1=<%=pUser1%>&user2=<%=pUser2%>&user3=<%=pUser3%>&idCategory=<%=pIdCategory%>"><%=getMsg(83,"Desc")%></a> - <a href="comersus_listItems.asp?hotDeal=<%=pHotDeal%>&orderBy=sku&strSearch=<%=pStrSearch%>&idSupplier=<%=pIdSupplier%>&user1=<%=pUser1%>&user2=<%=pUser2%>&user3=<%=pUser3%>&idCategory=<%=pIdCategory%>"><%=getMsg(84,"Sku")%></a> - <a href="comersus_listItems.asp?hotDeal=<%=pHotDeal%>&orderBy=visits&strSearch=<%=pStrSearch%>&idSupplier=<%=pIdSupplier%>&user1=<%=pUser1%>&user2=<%=pUser2%>&user3=<%=pUser3%>&idCategory=<%=pIdCategory%>"><%=getMsg(85,"Pop")%></a> - <a href="comersus_listItems.asp?hotDeal=<%=pHotDeal%>&orderBy=recently&strSearch=<%=pStrSearch%>&idSupplier=<%=pIdSupplier%>&user1=<%=pUser1%>&user2=<%=pUser2%>&user3=<%=pUser3%>&idCategory=<%=pIdCategory%>"><%=getMsg(86,"")%></a> - <a href="comersus_listItems.asp?hotDeal=<%=pHotDeal%>&orderBy=price&strSearch=<%=pStrSearch%>&idSupplier=<%=pIdSupplier%>&user1=<%=pUser1%>&user2=<%=pUser2%>&user3=<%=pUser3%>&idCategory=<%=pIdCategory%>"><%=getMsg(2,"price")%></a>
<%end if%>
</form>
<!--#include file="footer.asp"--> 
<%call closeDb()%>