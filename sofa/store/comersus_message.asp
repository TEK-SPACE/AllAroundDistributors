<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: prints a custom message with store header and footer
%>
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/screenMessages.asp"--> 
<!--#include file="../includes/stringFunctions.asp"--> 
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/cartFunctions.asp"--> 
<!--#include file="../includes/adSenseFunctions.asp"-->


<%
on error resume next

dim connTemp, rsTemp, mySql

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanySlogan 		= getSettingKey("pCompanySlogan")
pHeaderKeywords 	= getSettingKey("pHeaderKeywords")
pAdSenseClient 		= getSettingKey("pAdSenseClient")
pAuctions		= getSettingKey("pAuctions")
pSuppliersList		= getSettingKey("pSuppliersList")
pNewsLetter 		= getSettingKey("pNewsLetter")
pAffiliatesStoreFront   = getSettingKey("pAffiliatesStoreFront")
pAllowNewCustomer	= getSettingKey("pAllowNewCustomer")
pRssFeedServer		= getSettingKey("pRssFeedServer")
pShowNews 		= getSettingKey("pShowNews")
pShowSearchBox		= getSettingKey("pShowSearchBox")

' session
pIdCustomer     	= getSessionVariable("idCustomer",0)
pIdCustomerType 	= getSessionVariable("idCustomerType",1)
pIdDbSessionCart	= getSessionVariable("idDbSessionCart",0)

pMessage		= getScreenMessage(request.querystring("message"),3000)

%>
<!--#include file="header.asp"--> 
<br><b><%=getMsg(581,"attention")%></b><br><br>
<%=pMessage%>&nbsp;<%=session("message")%><br><br>
<%session("message")=""%>
<!--#include file="footer.asp"--> 
<%call closeDb()%> 
