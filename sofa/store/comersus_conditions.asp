<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: show general conditions
%>

<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/screenMessages.asp"--> 
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
pCompanySlogan	 	= getSettingKey("pCompanySlogan")
pAboutUsLink 		= getSettingKey("pAboutUsLink")
pHeaderKeywords 	= getSettingKey("pHeaderKeywords")
pShowSearchBox 		= getSettingKey("pShowSearchBox")

pAuctions		= getSettingKey("pAuctions")
pListBestSellers 	= getSettingKey("pListBestSellers")
pNewsLetter 		= getSettingKey("pNewsLetter")
pPriceList 		= getSettingKey("pPriceList")
pStoreNews 		= getSettingKey("pStoreNews")

' session
pIdDbSessionCart	= getSessionVariable("idDbSessionCart",0)

' get conditions

mySql="SELECT conditions FROM conditions WHERE idStore="&pIdStore

call getFromDatabase (mySql, rsTemp, "generalConditionsWindow")
		
if rsTemp.eof then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(474,"Conditions cannot be found in this store.")) 
end if
	
%>
<!--#include file="header.asp"--> 
<br><b><%=pCompany%>&nbsp;<%=getMsg(475,"terms & conditions")%></b>
<br><br><%=rsTemp("conditions")%>
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
if (window.print) {
document.write('<form>'
+ '<input type=button name=print value="<%=getMsg(476,"close")%>" '
+ 'onClick="javascript:window.close()"> </form>');
}
// End -->
</SCRIPT>    
<!--#include file="footer.asp"--> 
<%call closeDb()%>