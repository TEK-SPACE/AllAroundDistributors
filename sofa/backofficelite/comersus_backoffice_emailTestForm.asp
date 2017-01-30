<%
' Comersus Diagnostics
' e-commerce ASP Open Source
' CopyRight Rodrigo S. Alhadeff, Comersus
' September 2004
' http://www.comersus.com 
' Diagnostics for Comersus store
%>

<!--#include file="comersus_backoffice_adminVerify.asp"-->    
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"-->  
<!--#include file="../includes/getSettingKey.asp"--> 
<% 
on error resume next 

dim connTemp, rsTemp

' get settings 
pDefaultLanguage 	= getSettingKey("pDefaultLanguage")
pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanyLogo	 	= getSettingKey("pCompanyLogo")

pEmailSender		= getSettingKey("pEmailSender")
pEmailAdmin		= getSettingKey("pEmailAdmin")
pSmtpServer		= getSettingKey("pSmtpServer")
pEmailComponent		= getSettingKey("pEmailComponent")
pDebugEmail		= getSettingKey("pDebugEmail")

%>

<!--#include file="header.asp"--> 
<b>Email test form</b>
<br><br>
Checking email settings for this store... <br>
Admin email: <%=pEmailAdmin%> <br>
Sender email account: <%=pEmailSender%> <br>
SMTP Server: <%=pSmtpServer%> <br>
Email component: <%=pEmailComponent%> 

<br><br> 
<hr>
<form name="form1" method="post" action="comersus_backoffice_emailTestExec.asp">
  <table width="590" border="0">
    <tr> 
      <td>SMTP Server</td>
      <td>  
        <input type="text" name="pSmtpServer" value="<%=pSmtpServer%>">
        </td>
    </tr>
    <tr> 
      <td>Email From</td>
      <td>  
        <input type="text" name="pEmailFrom" value="<%=pEmailSender%>">
        </td>
    </tr>
    <tr> 
      <td>Destination Email</td>
      <td>  
        <input type="text" name="pEmailTo" value="Your email address here" size=25>
        </td>
    </tr>
    <tr> 
      <td> 
        <input type="submit" name="Test" value="Test">
      </td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form><br>
<%call closeDb()%>
<!--#include file="footer.asp"--> 
