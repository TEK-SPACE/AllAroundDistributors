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
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/miscFunctions.asp"--> 

<% 
on error resume next 

dim mySQL, conntemp, rstemp

pRunInstallationWizard 	= getSettingKey("pRunInstallationWizard")

if pRunInstallationWizard<>"-1" then 
 response.redirect "comersus_backoffice_message.asp?message="&Server.UrlEncode("Sorry, the Installation Wizard is disabled")
end if

' save settings into db

for each field in request.form 
 fieldValue = request.form(field)
   
 if fieldValue<>"" and field<>"Submit" then
  mySQL1="UPDATE settings SET settingValue='" &fieldValue& "' WHERE settingKey='" &field& "' AND idStore="&pIdStore
  call updateDatabase(mySQL1, rstemp, "comersus_backoffice_install.asp") 
 end if
 
next 

pEncryptionPassword="AJALWIYSJAH"&randomNumber(99999)&"ASQQTYAL"&randomNumber(999999)
%>
<!--#include file="header.asp"--> 
<!--#include file="./includes/settings.asp"-->

<br><font size="3"><b>Installation Wizard</b></font><br>
<br>Step 8 - Security<br><br>
<form method="post" name="addCateg" action="comersus_backoffice_install9.asp">
<table width="418" border="0">
     <tr> 
      <td width="189">Administrator password (1)</td>
      <td width="219">                
        <input type="password" name="pAdminPassword" value="" maxlength=12>
        <input type="hidden" name="pEncryptionPassword" value="<%=pEncryptionPassword%>" maxlength=30>  
      </td>
    </tr>                           
    </tr>                
    <tr> 
      <td width="189">&nbsp;</td>
      <td width="219">&nbsp;</td>
    </tr>        
    <tr> 
      <td colspan="2"> 
        <br>
        <input type="submit" name="Submit" value="Continue">
      </td>
    </tr>
</table>    
</form>
<br>Additional information 
<br>(1) Entering User=admin and the password selected here you will be able to login into the BackOffice Lite/+ 
<br>
<br>(2) Read also Security Concerns in Comersus User's Guide. There you can find tips before going live to maintain crackers out of your store. Read also this <a href="http://www.comersus.com/news-security.html" target="_blank">security article</a>
<!--#include file="footer.asp"-->
<%call closeDb()%>