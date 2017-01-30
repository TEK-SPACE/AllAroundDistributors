<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: show store news
%>

<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/screenMessages.asp"--> 
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"-->  
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/customerTracking.asp"-->
<!--#include file="../includes/cartFunctions.asp"-->
<!--#include file="../includes/adSenseFunctions.asp"-->

<% 

on error resume next

dim mySQL, connTemp, rsTempNews

' get settings 
pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanySlogan 		= getSettingKey("pCompanySlogan")
pAboutUsLink 		= getSettingKey("pAboutUsLink")
pStoreLocation		= getSettingKey("pStoreLocation")
pShowSearchBox 		= getSettingKey("pShowSearchBox")
pAdSenseClient 		= getSettingKey("pAdSenseClient")
pAffiliatesStoreFront   = getSettingKey("pAffiliatesStoreFront")
pAuctions		= getSettingKey("pAuctions")
pNewsLetter 		= getSettingKey("pNewsLetter")
pSuppliersList		= getSettingKey("pSuppliersList")
pAllowNewCustomer	= getSettingKey("pAllowNewCustomer")
pRssFeedServer		= getSettingKey("pRssFeedServer")
pShowNews		= getSettingKey("pShowNews")

pIdDbSessionCart	= getSessionVariable("idDbSessionCart",0)

call customerTracking("comersus_newsPageExec.asp", request.form)

' gets news
mySQL="SELECT newsDate, newsText FROM storeNews WHERE idStore=" &pIdStore& " ORDER BY idStoreNews DESC"

call getFromDatabase(mySql, rsTempNews, "newsPageExec")

if  rsTempNews.eof then 
  response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(302,"No news available")) 
end if

%>
<!--#include file="header.asp"--> 
 
<br><b><%=getMsg(303,"news at")%>&nbsp; <%=pCompany%></b>  
<%do while not rsTempNews.eof%>
	<br><br><hr width="450" align="left">
	<br><img src="images/news.jpg" alt="<%=rsTempNews("newsDate")%>">&nbsp;<%=rsTempNews("newsDate")%>
	<br><%=rsTempNews("newsText")%>
<%rsTempNews.movenext
loop   
%>
<!--#include file="footer.asp"--> 
<%call closeDb()%>
