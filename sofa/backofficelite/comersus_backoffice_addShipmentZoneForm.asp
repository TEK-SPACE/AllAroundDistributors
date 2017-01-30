<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"-->  
<!--#include file="../includes/settings.asp"-->   
<!--#include file="../includes/databaseFunctions.asp"-->   
<!--#include file="../includes/getSettingKey.asp"-->  
<%
on error resume next

dim mySQL, conntemp, rstemp

' get settings 
pCurrencySign	 	= getSettingKey("pCurrencySign")
%>
<!--#include file="./includes/settings.asp"-->

<!--#include file="header.asp"-->
<br><b>Add shipment zone</b><br>
<form method="post" name="addSh" action="comersus_backoffice_addShipmentZoneExec.asp">
  <table width="550" border="0">
    <tr> 
      <td width="150">Zone Name</td>
      <td width="400">
        <input type=text name="zoneDesc">        
      </td>
    </tr>
    
    <tr> 
      <td colspan=2>&nbsp;</td>            
    </tr>
    <tr> 
      <td> 
        <input type="submit" name="Submit" value="Add">
      </td>
    </tr>
  </table>
</form>
<!--#include file="footer.asp"-->
<%call closeDb()%>
