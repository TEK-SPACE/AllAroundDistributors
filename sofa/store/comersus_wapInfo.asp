<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: information about WAP 
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
pAboutUsLink 		= getSettingKey("pAboutUsLink")
pHeaderKeywords 	= getSettingKey("pHeaderKeywords")
pShowSearchBox 		= getSettingKey("pShowSearchBox")
pAdSenseClient 		= getSettingKey("pAdSenseClient")
pAuctions		= getSettingKey("pAuctions")
pAllowNewCustomer	= getSettingKey("pAllowNewCustomer")
pNewsLetter 		= getSettingKey("pNewsLetter")
pStoreNews 		= getSettingKey("pStoreNews")
pSuppliersList		= getSettingKey("pSuppliersList")
pRssFeedServer		= getSettingKey("pRssFeedServer")
pStoreLocation		= getSettingKey("pStoreLocation")

' session
pIdCustomer     	= getSessionVariable("idCustomer",0)
pIdCustomerType 	= getSessionVariable("idCustomerType",1)
pIdDbSessionCart	= getSessionVariable("idDbSessionCart",0)


%> 
<!--#include file="header.asp"--> 
<br><b><%=getMsg(317,"Wap Enabled")%></b><br><br>

<table border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="images/wap.jpg" width="90" height="81" alt="Wap Pic"></td>
    <td><%=getMsg(318,"you can also")%>
    <br><br><%=getMsg(319,"Point")%><i><a href="http://<%=pStoreLocation%>/wap">&nbsp;http://<%=pStoreLocation%>/wap</a></i>
    </td>
  </tr>
</table>
<br><br>
<!--#include file="footer.asp"--> 
<%call closeDb()%>