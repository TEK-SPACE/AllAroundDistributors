<% 
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: list active root categories
%>

<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/screenMessages.asp"-->
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/sessionFunctions.asp"-->  
<!--#include file="../includes/miscFunctions.asp"-->  
<!--#include file="../includes/itemFunctions.asp"-->  
<!--#include file="../includes/customerTracking.asp"-->  
<!--#include file="../includes/cartFunctions.asp"--> 
<!--#include file="../includes/adSenseFunctions.asp"-->

<% 
dim mySQL, connTemp, rsTemp, pIdCategory, pCategoryDesc, pIdAffiliate, pDefaultLanguage, pStoreFrontDemoMode, pCurrencySign, pDecimalSign, pCompany, pAuctions, pListBestSellers, pNewsLetter, pPriceList, pStoreNews, pAffiliatesStoreFront, pCategoriesAlphOrder, pAllowNewCustomer, pHeaderKeywords, pCustomerName, pHeaderCartItems, pHeaderCartSubtotal, pMoneyDontRound, pImageCategory

call saveCookie()

on error resume next 

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
   
   
' all root categories from category start
  
mySQL="SELECT idCategory, categoryDesc, imageCategory FROM categories WHERE active=-1 AND idParentCategory=" & pIdCategoryRoot & " AND idCategory<>" &pIdCategoryRoot
call getFromDatabase (mySql, rsTemp, "listCategoriesAndProducts") 

if rstemp.eof then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(7,"No ctg")) 
end if

call customerTracking("comersus_listCategories.asp", request.querystring)

%>

<!--#include file="header.asp"--> 
<br><b><%=getMsg(12,"Categories")%></b><br><br>
<table width=100% cellspacing="2" cellpadding="2">
<%
pColumnCounter=1

do while not rstemp.eof
 
 pIdCategory	=rstemp("idCategory")
 pCategoryDesc	=rstemp("categoryDesc")
 pImageCategory	=rstemp("imageCategory")
%>

 <%if pColumnCounter=1 then%> 
  <tr>
 <%end if%>
  
  <td>  
  <%if pImageCategory<>"" then%>
   <img src="catalog/<%=pImageCategory%>" alt="<%=pCategoryDesc%>"><br>
  <%end if%>
  <a href="comersus_listOneCategory.asp?idCategory=<%=pIdCategory%>"><%=pCategoryDesc%></a><br><br>  
  </td>
    
 <%if pColumnCounter=3 then%>    
  </tr>
 <%end if%>

<%
 
  pColumnCounter	= pColumnCounter+1
  
  if pColumnCounter=4 then 
   pColumnCounter=1
  end if
   
rstemp.movenext
loop

' adjust table
if pColumnCounter=2 then
 response.write "<td>&nbsp;</td><td>&nbsp;</td></tr>"
end if

if pColumnCounter=3 then
 response.write "<td>&nbsp;</td></tr>"
end if

%>
</table>

<!--#include file="footer.asp"--> 

<%call closeDb()%>