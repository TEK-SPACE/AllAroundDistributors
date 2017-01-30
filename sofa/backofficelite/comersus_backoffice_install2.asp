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
<!--#include file="../includes/stringFunctions.asp"--> 
<!--#include file="../includes/dateFunctions.asp"--> 

<% 
on error resume next 

dim mySQL, conntemp, rstemp

pRunInstallationWizard 	= getSettingKey("pRunInstallationWizard")

if pRunInstallationWizard<>"-1" then 
 response.redirect "comersus_backoffice_message.asp?message="&Server.UrlEncode("Sorry, the Installation Wizard is disabled")
end if

' save settings into db

pLoadStateCodes = 0
for each field in request.form 
 fieldValue = request.form(field)
   
 if fieldValue<>"" and field<>"Submit" then
  
  pFieldValue=formatForDb(fieldValue)
  
  mySQL1="UPDATE settings SET settingValue='" &pFieldValue& "' WHERE settingKey='" &field& "' AND idStore=" &pIdStore
  call updateDatabase(mySQL1, rstemp, "comersus_backoffice_install.asp") 
 end if

	if field = "pCompanyCountryCode" AND pFieldValue<>"US" Then
		'Delete US state codes
	  mySQL1="DELETE FROM stateCodes"
	  call updateDatabase(mySQL1, rstemp, "comersus_backoffice_install.asp") 
		pLoadStateCodes = -1
	end if 
next 
 
%>
<!--#include file="header.asp"--> 
<!--#include file="./includes/settings.asp"-->

<br><font size="3"><b>Installation Wizard</b></font><br>
<font size=2><br>Step 2 - Regionalization<br><br></font>
<form method="post" name="install2" action="comersus_backoffice_install3.asp">
<table width="418" border="0">
	<% if pLoadStateCodes = -1 Then%>
    <tr> 
      <td width="189">State Codes</td>
      <td width="219">        
      	<i>Load state codes for your country. Format: StateCode StateName</i>
	      <textarea name="pStateCodes" cols="35" rows=4></textarea>
      </td>
    </tr>    
  <%end if%>
    <tr> 
      <td width="189">Currency sign</td>
      <td width="219">        
        <input type="text" name="pCurrencySign" value="$">                
      </td>
    </tr>    
    <tr> 
      <td width="189">Decimal sign</td>
      <td width="219">        
        <input type="text" name="pDecimalSign" value=".">                        
      </td>
    </tr>
    <tr> 
      <td width="189">Date format</td>
      <td width="219">        
        <select name="pDateSwitch">      
<%
if getServerDateFormat()="MM/DD/YYYY" then%>
 <option value='-1'>DD/MM/YYYY</option>
 <option value='0'>MM/DD/YYYY</option>
<%else%>
 <option value='0'>DD/MM/YYYY</option>
 <option value='-1'>MM/DD/YYYY</option>
<%end if%>          
 
        </select>
      </td>
    </tr>   
    <tr> 
      <td colspan="2"> 
        <br>
        <input type="submit" name="Submit" value="Continue">
      </td>
    </tr>
</table>    
</form>
<br>You can use other languages as well. 
<br>Download free language translations at http://www.comersus.com/free.html 
<%call closeDb()%>
<!--#include file="footer.asp"-->