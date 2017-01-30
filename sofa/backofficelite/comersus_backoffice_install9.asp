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
<!--#include file="../includes/encryption.asp"--> 
<!--#include file="../store/comersus_optDes.asp"--> 

<% 
on error resume next 

dim mySQL, conntemp, rstemp

pRunInstallationWizard 	= getSettingKey("pRunInstallationWizard")
pEncryptionMethod 	= getSettingKey("pEncryptionMethod")

if pRunInstallationWizard<>"-1" then 
 response.redirect "comersus_backoffice_message.asp?message="&Server.UrlEncode("Sorry, the Installation Wizard is disabled")
end if

pEncryptionPassword	=request.form("pEncryptionPassword")

if pEncryptionPassword<>"" then

  mySQL1="UPDATE settings SET settingValue='" &pEncryptionPassword& "' WHERE settingKey='pEncryptionPassword' AND idStore="&pIdStore            
  call updateDatabase(mySQL1, rstemp, "comersus_backoffice_install.asp") 

 ' update admin profile with new password
 pAdminPassword		=request.form("pAdminPassword")
 
 if len(pAdminPassword)<6 then
  response.redirect "comersus_backoffice_message.asp?message="&Server.UrlEncode("Please use a password with 6 characters or more. Go back and try again.")
 end if

 if pAdminPassword="" then pAdminPassword="123456"

 pEncryptedAdminPassword =Cstr(EnCrypt(pAdminPassword, pEncryptionPassword))

 ' clear previous admins
 mySQL1="DELETE FROM admins"
 call updateDatabase(mySQL1, rstemp, "comersus_backoffice_install.asp") 

 ' insert into admins
 mySQL1="INSERT INTO admins (adminName, adminLevel, adminPassword, idStore) VALUES ('Admin',0,'" &pEncryptedAdminPassword& "'," &pIdStore& ")"
 call updateDatabase(mySQL1, rstemp, "comersus_backoffice_install.asp")     
  
end if

%>
<!--#include file="header.asp"--> 
<!--#include file="./includes/settings.asp"-->

<br><font size="3"><b>Installation Wizard</b></font><br>
<br>Step 9 - Power Packs<br><br>
<form method="post" name="install" action="comersus_backoffice_install10.asp">
<table width="418" border="0">
     <tr> 
      <td width="189">Power Pack Installed (1)</td>
      <td width="219">                
      <select name="pPowerPacksInstalled">
        <option value='NONE'>No</option>
        <option value='PPP'>Yes</option>
        </select>  
      </td>
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
<br>(1) With Power Pack you can obtain advanced features for your store. <a href="http://www.comersus.com/power-pack.html" target="_blank">More information about Power Pack here...</a>
<!--#include file="footer.asp"-->