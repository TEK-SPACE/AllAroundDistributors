<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: contact us
%>

<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/miscFunctions.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/cartFunctions.asp"--> 
<!--#include file="../includes/screenMessages.asp"--> 
<!--#include file="../includes/adSenseFunctions.asp"-->

<%
on error resume next

dim connTemp, rsTemp

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanySlogan		= getSettingKey("pCompanySlogan")
pAboutUsLink 		= getSettingKey("pAboutUsLink")
pCompanyPhone	 	= getSettingKey("pCompanyPhone")
pCompanyAddress	 	= getSettingKey("pCompanyAddress")
pCompanyStateCode 	= getSettingKey("pCompanyStateCode")
pCompanyCity		= getSettingKey("pCompanyCity")
pCompanyZip	 	= getSettingKey("pCompanyZip")
pCompanyCountryCode 	= getSettingKey("pCompanyCountryCode")
pEmailAdmin		= getSettingKey("pEmailAdmin")
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
pShowNews		= getSettingKey("pShowNews")

pIdCustomer	 	= getSessionVariable("idCustomer",0)
pIdCustomerType  	= getSessionVariable("idCustomerType",1)
pIdDbSessionCart	= getSessionVariable("idDbSessionCart",0)

%>
<!--#include file="header.asp"--> 
<br><b><%=getMsg(557,"contact us")%></b>
<br><br><%=pCompany%>
<br><i><%=pCompanySlogan%></i>
<br><br><%=pCompanyAddress%> 
<br><%=getStateName(pCompanyStateCode)%> - <%=pCompanyCity%> (<%=pCompanyZip%>)
<br><%=getCountryName(pCompanyCountryCode)%>
<br><%=getMsg(558,"phone")%>&nbsp; <%=pCompanyPhone%>
<br><%=getMsg(559,"email")%>&nbsp; <a href="mailto:<%=pEmailAdmin%>"><%=pEmailAdmin%></a><br>
<%if pCompanyCountryCode = "US" or pCompanyCountryCode = "CA" then%>
    <br><a target="_blank" href="http://maps.yahoo.com/maps_result?addr=<%=pCompanyAddress%>&csz=<%=pCompanyZip%>&country=<%=pCompanyCountryCode%>&new=1&name=&qty="><%=getMsg(722,"map & direct")%></a>
<%end if%>
<!--#include file="footer.asp"--> 
<%call closeDb()%>