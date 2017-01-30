<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: shows order confirmation
%>
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/screenMessages.asp"--> 
<!--#include file="../includes/stringFunctions.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/adSenseFunctions.asp"-->

<%
on error resume next

dim mySQL, conntemp, rsTemp

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanyLogo	 	= getSettingKey("pCompanyLogo")

pAuctions		= getSettingKey("pAuctions")
pListBestSellers 	= getSettingKey("pListBestSellers")
pNewsLetter 		= getSettingKey("pNewsLetter")
pPriceList 		= getSettingKey("pPriceList")
pStoreNews 		= getSettingKey("pStoreNews")
pOneStepCheckout	= getSettingKey("pOneStepCheckout")
pExitSurvey		= getSettingKey("pExitSurvey")

pOrderPrefix	 	= getSettingKey("pOrderPrefix")

pIdOrder		= getUserInput(request.querystring("idOrder"),10)

if pIdOrder="" then
   response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(490,"Invalid order"))
end if
%>

<!--#include file="header.asp"--> 
<br><b><%=getMsg(491,"confirmation")%></b><br>
<br><%=getMsg(492,"thanks")%>
<br><%=getMsg(493,"payment 4 order #")%><%=pIdOrder%>&nbsp; <%=getMsg(494,"received")%>
<br><br><%=getMsg(495,"track")%> <a href="comersus_customerShowOrderDetails.asp?idOrder=<%=pIdOrder%>"><%=getMsg(496,"here")%></a>

<%if pStoreFrontDemoMode="-1" then%>
 <br><br><%=getMsg(497,"thanks for testing")%>
 <br><%=getMsg(498,"visit")%>&nbsp; <a href="http://www.comersus.com/backOfficeTest/backOfficePlus"><%=getMsg(499,"admin demo")%></a>
<%end if%>

<!--#include file="footer.asp"--> 
<%call closeDb()%> 
