<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: customer menu
%>
<!--#include file="comersus_customerLoggedVerify.asp"-->
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/screenMessages.asp"-->
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->  
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/stringFunctions.asp"-->  
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/cartFunctions.asp"-->
<!--#include file="../includes/adSenseFunctions.asp"-->
<%
dim connTemp, rsTemp, mySql

on error resume next

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

pAuctions		= getSettingKey("pAuctions")
pSuppliersList		= getSettingKey("pSuppliersList")
pNewsLetter 		= getSettingKey("pNewsLetter")
pAffiliatesStoreFront   = getSettingKey("pAffiliatesStoreFront")
pAllowNewCustomer	= getSettingKey("pAllowNewCustomer")
pRssFeedServer		= getSettingKey("pRssFeedServer")
pDateSwitch		= getSettingKey("pDateSwitch")
pShowNews		= getSettingKey("pShowNews")
pWholesaleConversionAmount	= getSettingKey("pWholesaleConversionAmount")
pAdSenseClient 		= getSettingKey("pAdSenseClient")

pRecommendations	= getSettingKey("pRecommendations")
pBonusPointsPerPrice	= getSettingKey("pBonusPointsPerPrice")

pIdCustomer	 	= getSessionVariable("idCustomer",0)
pIdCustomerType	 	= getSessionVariable("idCustomerType",1)
pIdDbSessionCart	= getSessionVariable("idDbSessionCart",0)

pWishListCount		= 0

mySQL="SELECT name, lastName, bonusPoints, birthDay FROM customers WHERE idCustomer=" &pIdCustomer

call getFromDatabase(mySQL, rstemp, "utilitiesMenu")

if rstemp.eof then
 response.redirect "comersus_supportError.asp?error=Cannot locate customer record"
end if

pCustomerName	=rstemp("name") & " " &rstemp("lastName")
pBounsPoints	=rstemp("bonusPoints")
pBirthDay	=rstemp("birthDay")

' count orders
mySQL="SELECT COUNT(*) AS ordersCount FROM orders WHERE idCustomer=" &pIdCustomer

call getFromDatabase(mySQL, rstemp, "utilitiesMenu")

pOrdersCount=rstemp("ordersCount")

' count wish list

mySQL="SELECT COUNT(*) AS wlCount FROM wishList WHERE idCustomer=" &pIdCustomer

call getFromDatabase(mySQL, rstemp, "utilitiesMenu")

pWishListCount=rstemp("wlCount")

%>
<!--#include file="header.asp"--> 
<br><br><b><%=getMsg(148,"Acc")%></b><br>
<br>
<%if pIdCustomerType=2 then%>
  <img src="images/buttonWholesaleSmall.gif" alt="<%=getMsg(151,"Whole")%>">  
<%else%>
  <img src="images/buttonRetailSmall.gif" alt="<%=getMsg(152,"Ret")%>">
  <%if pWholesaleConversionAmount>0 then%>
   <i>* <%=getMsg(755,"Spend")%>&nbsp; <%=pCurrencySign&money(pWholesaleConversionAmount)%>&nbsp; <%=getMsg(756,"2 reach")%></i>
  <%end if%>
<%end if%>

<br><%=getMsg(149,"Name")%>: <%=pCustomerName%>

<br><%=getMsg(153,"Orders")%>: <a href="comersus_customerShowOrders.asp"><%=pOrdersCount%></a>
<br><%=getMsg(203,"Your WL has")%> <a href="comersus_customerWishListView.asp"><%=pWishListCount%></a> <%=getMsg(204,"Items")%>
<br><%=getMsg(154,"Birth")%>: 

<%if pBirthDay="" or isNull(pBirthDay) then%>
 <%=getMsg(155,"N def")%> &nbsp; <a href="comersus_birthdayEditForm.asp"><%=getMsg(156,"Enter")%></a>
<%else%>
 <%=formatDate(pBirthDay)%> <a href="comersus_birthdayEditForm.asp"><%=getMsg(157,"Change")%></a>
<%end if%>

<%if pBonusPointsPerPrice<>"0" then%>
 <br><%=getMsg(158,"Bon points")%>: <%=money(pBounsPoints)%>
<%end if%>

<br>
<br>- <a href="comersus_customerModifyForm.asp"><%=getMsg(159,"Modify Info")%></a>
<br>- <a href="comersus_customerWishListView.asp"><%=getMsg(160,"W List")%></a>


<%if pAuctions="-1" then%>
  <br>- <a href="comersus_optAuctionListAll.asp"><%=getMsg(162,"Auct")%></a>          
<%end if%>

<%if pRecommendations="-1" then%>
  <br>- <a href="comersus_optGetRecommendations.asp"><%=getMsg(163,"Cust recom")%></a>          
<%end if%>
<br>- <a href="comersus_rmaForm.asp"><%=getMsg(734,"Request RMQ")%></a></a>
<br>- <a href="comersus_customerContactAdminForm.asp"><%=getMsg(164,"Contact")%></a>
<br>- <a href="comersus_customerLogout.asp"><%=getMsg(165,"Logout")%></a>  
<br><br>

<!--#include file="footer.asp"--> 
<%call closeDb()%>


