<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="../includes/settings.asp"-->    
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/encryption.asp"--> 

<%
on error resume next

dim mySQL, conntemp, rstemp

pRunInstallationWizard 	= getSettingKey("pRunInstallationWizard")
pEncryptionMethod 	= getSettingKey("pEncryptionMethod")
pEncryptionPassword	= getSettingKey("pEncryptionPassword")

if pRunInstallationWizard<>"-1" then 
 response.redirect "comersus_backoffice_message.asp?message="&Server.UrlEncode("Sorry, the Installation Wizard is disabled")
end if

%>
<!--#include file="header.asp"--> 
<!--#include file="./includes/settings.asp"-->
<br><font size="3"><b>Installation Wizard - Google AdSense Setup</b></font><br><br>
<img src="images/adsense.gif" alt="Example of additional revenue with Google AdSense" border=1>
<br>
<br><br>Want to get additional revenue from your store? 
<br><br>You can also display Ads related to your products or services and get paid for each click. You can configure permanent Ads in the header or display just one AD in payment confirmation page. Certain clicks may get up to $5 dollars.

<br><br><a href="comersus_backoffice_install11.asp">No thanks...</a>

<br><br>Yes? Please follow these steps:

<br><br>1. <a href="https://www.google.com/adsense/" target="_blank">Sign Up for Google AdSense</a>
<br><br>2. Enter your Google Ad Sense Id below (My Account Tab, Property Information, Ad Sense for content, example: ca-pub-99999999999999)and click update <form method=post action="comersus_backoffice_updateAdSense.asp" target="_blank"><input type=text name="pAdSenseClient" size=20><input type=submit value="Update"></a>
<br><br>3. Select where do you want the ads: <a href="comersus_backoffice_updateAdSense.asp?pAdSenseType=permanent" target="_blank">Permanent Ads in the header and item details</a> - <a href="comersus_backoffice_updateAdSense.asp?pAdSenseType=confirmation" target="_blank">One Ad in payment confirmation page</a>
<br><br>4. Finished? <a href="comersus_backoffice_install11.asp">Continue with the Wizard</a> 

<br><br><i>Note: 
<br>Please read Google AdSense complete terms and conditions at https://www.google.com/adsense/localized-terms and Comersus notes about adSense in the User´s Guide.</i>
<!--#include file="footer.asp"-->