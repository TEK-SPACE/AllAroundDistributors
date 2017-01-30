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

<% 
on error resume next

dim mySQL, conntemp, rstemp

pRunInstallationWizard 	= getSettingKey("pRunInstallationWizard")

if pRunInstallationWizard<>"-1" then 
 response.redirect "comersus_backoffice_message.asp?message="&Server.UrlEncode("Sorry, the Installation Wizard is disabled")
end if

pTaxValueSameCountry	   =request.form("taxValueSameCountry")
pTaxValueDifferentCountry  =request.form("taxValueDifferentCountry")
pTaxCountryCode		   =request.form("taxCountryCode")
pTaxStateCode		   =request.form("taxStateCode")

pTaxValueDifferentCountry  = replace(pTaxValueDifferentCountry,",",".")
pTaxValueSameCountry	   = replace(pTaxValueSameCountry,",",".")

if pTaxValueSameCountry<>"" or pTaxValueDifferentCountry<>"" then

	mySQL="DELETE FROM taxPerPlace WHERE idStore=" &pIdStore
	call updateDatabase(mySQL, rstemp, "comersus_backoffice_install6.asp") 
	
	mySQL="INSERT INTO taxPerPlace (idStore, taxPerPlace,"	
	
	if pTaxCountryCode<>"" then
		mySQL= mySQL & "countryCode, countryCodeEq,"
	end if	
	if pTaxStateCode<>"" then
		mySQL= mySQL & "stateCode, stateCodeEq,"
	end if
	mySQL1= mySQL & " idCustomerType) VALUES (" &pIdStore & "," &pTaxValueDifferentCountry& ","
	mySQL= mySQL & " idCustomerType) VALUES  (" &pIdStore & ","  &pTaxValueSameCountry& ","	
	if pTaxCountryCode<>"" then
		mySQL= mySQL & "'" &pTaxCountryCode& "',-1,"
		mySQL1= mySQL1 & "'" &pTaxCountryCode& "',0,"
	end if			
	if pTaxStateCode<>"" then
		mySQL= mySQL & "'" &pTaxStateCode& "',-1,"
		mySQL1= mySQL1 & "'" &pTaxStateCode& "',0,"
	end if	
	mySQL= mySQL & "1)"
	mySQL1= mySQL1 & "1)"
end if	
if pTaxValueSameCountry<>"" then
	call updateDatabase(mySQL, rstemp, "comersus_backoffice_install6.asp") 
end if
if pTaxValueDifferentCountry<>"" and mySQL<>mySQL1 then 
	call updateDatabase(mySQL1, rstemp, "comersus_backoffice_install6.asp") 
end if

%>
<!--#include file="header.asp"--> 
<!--#include file="./includes/settings.asp"-->

<br><font size="3"><b>Installation Wizard</b></font><br>
<br>Step 7 - Email Settings<br><br>
<form method="post" name="addCateg" action="comersus_backoffice_install8.asp">
<table width="418" border="0">
     <tr> 
      <td width="189">SMTP Server (1)</td>
      <td width="219">                
	<input type="text" name="pSmtpServer" value=""> 
      </td>
    </tr>
    <tr> 
      <td width="189">Valid email POP account for your domain </td>
      <td width="219">                
        <input type="text" name="pEmailSender" value=""> 
      </td>
    </tr>    
    <tr> 
      <td width="189">Email account of the store administrator</td>
      <td width="219">                
        <input type="text" name="pEmailAdmin" value=""> 
      </td>
    </tr>
    <tr> 
      <td width="189">Email component (2)</td>
      <td width="219">       
      <select name="pEmailComponent">
        <option value='NONE' selected>None</option>
        <option value='CDONTS'>Cdonts</option>
        <option value='JMAIL'>Jmail</option>
        <option value='PersitsASPMail'>ASPMail</option>
        <option value='CDO'>CdoSys for Windows 2003</option>
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

Additional information 
<br>(1) You can retrieve this information from your hosting service welcome email message or from your email client configuration (Tools/Account/Properties for Outlook Express) An example is: mail.yourdomain.com or smtp.yourdomain.com
<br>
<br>(2) Ask your hosting support about which ASP email components are installed in your server. Cdonts and Jmail are very popular. If the component is not listed here you can configure that setting from BackOffice/Configuration/Settings If you are not sure leave NONE or you are going to get errors in the StoreFront
<!--#include file="footer.asp"-->
<%call closeDb()%>