<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: show information about Cookies
%>

<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/currencyFormat.asp"-->
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
pCompanyLogo	 	= getSettingKey("pCompanyLogo")
pAboutUsLink 		= getSettingKey("pAboutUsLink")
pHeaderKeywords 	= getSettingKey("pHeaderKeywords")

pEmailAdmin	 	= getSettingKey("pEmailAdmin")

pAuctions		= getSettingKey("pAuctions")
pListBestSellers 	= getSettingKey("pListBestSellers")
pNewsLetter 		= getSettingKey("pNewsLetter")
pPriceList 		= getSettingKey("pPriceList")
pStoreNews 		= getSettingKey("pStoreNews")
pOneStepCheckout	= getSettingKey("pOneStepCheckout")
%>

<!--#include file="header.asp"--> 
<br><p><b>Attention</b></p>
<p>Our tests indicate that your browser is not accepting cookies<br>
  This store relies heavily on cookies, small bits of information which are saved 
  by your browser and are used by the store to identify you.</p>
<p><b>How to enable Cookies?</b></p>
<p> <b>Internet Explorer 4.x for PC</b>: From inside the IE window: Click on View 
  in the top navigation bar and then click on the Internet Options link. When 
  the new window opens, click on the Advanced tab. Scroll down until you find 
  the "Security" heading. Scroll down to the "Cookies" heading. Under "Allow cookies 
  that are stored on your computer," please check "Enable,". Under "Allow per 
  session cookies (not stored)," check "Allow per session cookies." <br>
  <br>
  <b>Internet Explorer 5.x for PC</b>: From inside the IE window: Click on Tools 
  in the top navigation bar and then click Internet Options. Click the Security 
  tab. Select the Internet for the zone that you want to set the security level 
  for. Click on the "Custom Level" button located at the bottom of the page. When 
  the "Settings:" window opens, scroll down to "Cookies". Under "Allow cookies 
  that are stored on your computer," please check "Enable,". Under "Allow per 
  session cookies (not stored)," check "Allow per session cookies." <br>
</p>
<p><b>Netscape 4.x for PC</b>: From inside the Netscape window: Go to Edit, then 
  click on Preferences. On the left pane, click on Advanced, then on the right, 
  check "accept all cookies" or "accept only cookies that get sent back to the 
  originating server". Click OK to save. <br>
  <br>
  <b>Netscape 3.x for PC </b>(Note: Navigator 3.0x and 2.0x don't have an option 
  to disable the setting of cookies globally. It is a feature added to version 
  4.x.): Go to the Options menu and choose Network Preferences, Protocols Check 
  the option to "show alert before accepting cookies." You can also check out 
  information from the Netscape Web site at http://help.netscape.com/kb/consumer/19990515-1.html 
</p>
<p><i>Note: consider that certain security software like Norton Internet Security and others could also block Cookies</i></p>
<p>If you need further assistance please contact us: <%=pEmailAdmin%></p>
<!--#include file="footer.asp"--> 
<%call closeDb()%>