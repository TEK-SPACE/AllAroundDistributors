<% 
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: list one category
%>

<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/screenMessages.asp"-->
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/cartFunctions.asp"-->
<!--#include file="../includes/sessionFunctions.asp"-->  
<!--#include file="../includes/miscFunctions.asp"-->  
<!--#include file="../includes/itemFunctions.asp"-->  
<!--#include file="../includes/customerTracking.asp"-->  
<!--#include file="../includes/adSenseFunctions.asp"-->

<% 
dim mySQL, connTemp, rsTemp, pIdCategory, pCategoryDesc, pIdAffiliate, pStoreFrontDemoMode, pCurrencySign, pDecimalSign, pCompany, pAuctions, pListBestSellers, pNewsLetter, pPriceList, pStoreNews, pAffiliatesStoreFront, pCategoriesAlphOrder, pAllowNewCustomer, pHeaderKeywords, pCustomerName, pHeaderCartItems, pHeaderCartSubtotal, pMoneyDontRound, pImageCategory, pCategoryString, pSomethingInside

on error resume next 

call saveCookie()

' set affiliate
pIdAffiliate=getUserInput(request.querystring("idAffiliate"),4)

if isNumeric(pIdAffiliate)then
   session("idAffiliate")= pIdAffiliate
end if

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanySlogan 		= getSettingKey("pCompanySlogan")
pAboutUsLink 		= getSettingKey("pAboutUsLink")
pStoreLocation		= getSettingKey("pStoreLocation")
pHeaderKeywords 	= getSettingKey("pHeaderKeywords")
pShowSearchBox 		= getSettingKey("pShowSearchBox")
pAdSenseClient 		= getSettingKey("pAdSenseClient")

pAuctions		= getSettingKey("pAuctions")
pSuppliersList		= getSettingKey("pSuppliersList")
pNewsLetter 		= getSettingKey("pNewsLetter")
pAffiliatesStoreFront   = getSettingKey("pAffiliatesStoreFront")
pAllowNewCustomer	= getSettingKey("pAllowNewCustomer")
pRssFeedServer		= getSettingKey("pRssFeedServer")
pCategoriesAlphOrder	= getSettingKey("pCategoriesAlphOrder")
pShowNews		= getSettingKey("pShowNews")

' session
pIdCustomer     	= getSessionVariable("idCustomer",0)
pIdCustomerType 	= getSessionVariable("idCustomerType",1)
pIdDbSessionCart	= getSessionVariable("idDbSessionCart",0)

pIdCategoryRoot		= getCategoryStart(pIdStore)

' get category
pIdCategory		= getUserInput(request.querystring("idCategory"),8)  

pSomethingInside	= 0

call customerTracking("comersus_listOneCategory.asp", request.querystring)

' get category details
  
mySQL="SELECT idCategory, categoryDesc, details, imageCategory, keywords FROM categories WHERE idCategory=" & pIdCategory
call getFromDatabase (mySql, rsTemp, "listCategoriesAndProducts") 

if rstemp.eof then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(6,"Invalid ctg"))
end if

pCategoryDesc	=rstemp("categoryDesc")
pDetails	=rstemp("details")
pImageCategory	=rstemp("imageCategory")
pSearchKeywords = rsTemp("keywords")

pCategoryString=pCategoryDesc

call getCategoryPath(pIdCategory, pCategoryString)

pArrayCategories=split(pCategoryString,"|")

' check redirection to listItems
if isCategoryLeaf(pIdCategory) then
 
 mySQL="SELECT COUNT(*) AS productsCount FROM products, categories_products WHERE products.idProduct=categories_products.idProduct AND categories_products.idCategory="& pIdCategory&" AND listHidden=0 AND active=-1 AND idStore=" &pIdStore
 call getFromDatabase (mySql, rsTemp, "listCategoriesAndProducts") 

 if not rsTemp.eof then
  response.redirect "comersus_listItems.asp?idCategory="&pIdCategory
 end if
    
end if

%>

<!--#include file="header.asp"--> 
<br><b><%=getMsg(8,"You are at")%> 

<%for f=uBound(pArrayCategories) to 0 step -1%>
 > <a href="comersus_listOneCategory.asp?idCategory=<%=getCategoryId(pArrayCategories(f))%>"><%=pArrayCategories(f)%></a>
<%next%>

</b><br>

<%if pImageCategory<>"" then%>
   <br><br><img src="catalog/<%=pImageCategory%>" alt="<%=pCategoryDesc%>"><br>
<%end if%>
  
<br><%=pDetails%>
<br><br><b><%=getMsg(744,"subcat")%></b><br><br>

<table width=100%>  

    <%
    
    mySQL="SELECT idCategory, categoryDesc, details, imageCategory FROM categories WHERE active=-1 AND idParentCategory=" & pIdCategory& " ORDER BY displayOrder"
    call getFromDatabase (mySql, rsTemp, "listCategoriesAndProducts") 
    
    pColNumber=1
    
    do while not rsTemp.eof
    	pSomethingInside=-1
    	%>
    
    <%if pColNumber=1 then%>
     <tr>
    <%end if%>
    
    <td>    
     <%if rstemp("imageCategory")<>"" then%>
      <img src="catalog/<%=rstemp("imageCategory")%>" alt="<%=rstemp("categoryDesc")%>"><br>
     <%end if%>
     <a href="comersus_listOneCategory.asp?idCategory=<%=rstemp("idCategory")%>"><%=rstemp("categoryDesc")%></a><br><br>
    </td>
     
   <%if pColNumber=4 then%>    
    </tr>
   <%end if%>
      
    	<% 	   
   
   pColNumber	= pColNumber+1
   
   if pColNumber=5 then 
    pColNumber=1
   end if
       	
    rsTemp.movenext
    loop  
%>

<%if pColNumber=2 then%>
<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>
<%end if%>

<%if pColNumber=3 then%>
 <td>&nbsp;</td><td>&nbsp;</td></tr>
<%end if%>
 
<%if pColNumber=4 then%>
 <td>&nbsp;</td></tr>
<%end if%>
    
 <tr><td colspan=4>&nbsp;</td></tr>
 
</table>

<%if pSomethingInside=0 then%>
 <%=getMsg(9,"No items")%>
<%end if%>

<!--#include file="footer.asp"--> 

<%call closeDb()%>
