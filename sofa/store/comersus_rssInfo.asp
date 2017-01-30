<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: information about RSS 
%>
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/screenMessages.asp"-->
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/cartFunctions.asp"--> 
<!--#include file="../includes/adSenseFunctions.asp"-->


<%
on error resume next 

dim connTemp, rsTemp

' get settings 
pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanySlogan 		= getSettingKey("pCompanySlogan")
pStoreLocation	 	= getSettingKey("pStoreLocation")
pAboutUsLink 		= getSettingKey("pAboutUsLink")
pHeaderKeywords 	= getSettingKey("pHeaderKeywords")
pShowSearchBox 		= getSettingKey("pShowSearchBox")

pAuctions		= getSettingKey("pAuctions")
pAllowNewCustomer	= getSettingKey("pAllowNewCustomer")
pNewsLetter 		= getSettingKey("pNewsLetter")
pStoreNews 		= getSettingKey("pStoreNews")
pSuppliersList		= getSettingKey("pSuppliersList")
pRssFeedServer		= getSettingKey("pRssFeedServer")

' session
pIdCustomer	 	= getSessionVariable("idCustomer",0)
pIdCustomerType  	= getSessionVariable("idCustomerType",1)
pIdDbSessionCart	= getSessionVariable("idDbSessionCart",0)

%> 
<!--#include file="header.asp"--> 
<br><b><%=getMsg(324,"title")%></b><br><br>
<img src="images/rss.gif" alt="RSS Feed - Shopping Cart" border=0>
<br><%=getMsg(325,"this store...")%>
<br><br><%=getMsg(328,"connect")%> <a href="http://<%=pStoreLocation%>/xml/news.asp"><%=getMsg(327,"here")%></a> 
<br><%=getMsg(326,"connect")%> <a href="http://<%=pStoreLocation%>/xml/products.asp"><%=getMsg(327,"here")%></a>
<br><br>
<!--#include file="footer.asp"--> 
<%call closeDb()%>