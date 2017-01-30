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

<% 
'on error resume next 

dim mySQL, conntemp, rstemp, rstemp2

pRunInstallationWizard 	= getSettingKey("pRunInstallationWizard")
pCompanyCountryCode	= getSettingKey("pCompanyCountryCode")
pCompanyStateCode	= getSettingKey("pCompanyStateCode")

if pRunInstallationWizard<>"-1" then 
 response.redirect "comersus_backoffice_message.asp?message="&Server.UrlEncode("Sorry, the Installation Wizard is disabled")
end if

pByPassShipping	=getUserInput(request.form("pByPassShipping"),4)

if pByPassShipping<>"" then
  mySQL1="UPDATE settings SET settingValue='" &pByPassShipping& "' WHERE settingKey='pByPassShipping' AND idStore=" &pIdStore
  call updateDatabase(mySQL1, rstemp, "comersus_backoffice_install.asp") 
end if

' insert zone into db
mySQL="INSERT INTO shippingZones(zoneName) VALUES ('Local zone')"
call updateDatabase(mySQL, rstemp, "comersus_backoffice_addShipmentZoneExec.asp") 

' get the ZoneId
mySQL="SELECT Max(idShippingZone) as MaxId FROM shippingZones WHERE zoneName='Local zone'"
call getFromDatabase(mySQl, rstemp, "install5b")

pIdShippingZone = cInt(rstemp("MaxId"))

' insert Zone Contents
mySQL="INSERT INTO shippingZonesContents(idShippingZone, countryCode) VALUES ("&pIdShippingZone&", '" &pCompanyCountryCode& "')"
call updateDatabase(mySQL, rstemp, "comersus_backoffice_addShipmentZoneExec.asp") 

pServiceType1	 ="Default shipping carrier service"
pWeightFrom1	 =0
pWeightUntil1	 =99999
pShippingAmount1 =request.form("pShippingAmount1")

pShippingAmount1 = formatNumberForDb(pShippingAmount1)

if pByPassShipping="0" then
  mySQL1="INSERT INTO shipments (shipmentDesc, weightFrom, weightUntil, idShippingZone, priceToAdd, quantityUntil, priceUntil, idCustomerType, idStore) VALUES ('"&pServiceType1&"'," &pWeightFrom1& "," &pWeightUntil1& ",'" &pIdShippingZone& "'," &pShippingAmount1& ",9999,999999, NULL,"&pIdStore&")"
  call updateDatabase(mySQL1, rstemp, "comersus_backoffice_install.asp") 
end if

' insert zone into db
mySQL="INSERT INTO shippingZones(zoneName) VALUES ('Worldwide zone')"
call updateDatabase(mySQL, rstemp, "comersus_backoffice_addShipmentZoneExec.asp") 

' get the ZoneId
mySQL="SELECT Max(idShippingZone) as MaxId FROM shippingZones WHERE zoneName='Worldwide zone'"
call getFromDatabase(mySQl, rstemp, "install5b")

pIdShippingZone = cInt(rstemp("MaxId"))

' insert Zone Contents, all countries but local

mySQL="SELECT * FROM countryCodes WHERE countryCode<>'"&pCompanyCountryCode&"'"
call updateDatabase(mySQL, rstemp, "comersus_backoffice_addShipmentZoneExec.asp") 

do while not rstemp.eof 
 mySQL="INSERT INTO shippingZonesContents(idShippingZone, countryCode) VALUES ("&pIdShippingZone&", '" &rstemp("countryCode")& "')"
 call updateDatabase(mySQL, rstemp2, "comersus_backoffice_addShipmentZoneExec.asp") 
 rstemp.movenext
loop

pServiceType2	 ="Default shipping carrier service"
pWeightFrom2	 =0
pWeightUntil2	 =99999

pShippingAmount2 =request.form("pShippingAmount2")

pShippingAmount2 = formatNumberForDb(pShippingAmount2)

if pByPassShipping="0" then
  mySQL1="INSERT INTO shipments (shipmentDesc, weightFrom, weightUntil, idShippingZone, priceToAdd, quantityUntil, priceUntil, idCustomerType, idStore) VALUES ('"&pServiceType2&"'," &pWeightFrom2& "," &pWeightUntil2& ",'" &pIdShippingZone& "'," &pShippingAmount2& ",9999,999999, NULL,"&pIdStore&")"
  call updateDatabase(mySQL1, rstemp, "comersus_backoffice_install.asp") 
end if


%>
<!--#include file="header.asp"--> 
<!--#include file="./includes/settings.asp"-->

<br><font size="3"><b>Installation Wizard</b></font><br>
<br>Step 6 - Taxes<br><br></font>
<form method="post" name="addCateg" action="comersus_backoffice_install7.asp">
 
<input type="hidden" name="taxCountryCode" VALUE="<%=pCompanyCountryCode%>">

<%if len(pCompanyStateCode)=2 then%>
 <input type="hidden" name="taxStateCode" VALUE="<%=pCompanyStateCode%>">
<%end if%>
 
<table width="618" border="0">
     <tr> 
      <td width="269">Enter the tax percentage for customers located in       
      <%if len(pCompanyStateCode)=2 then%>
       <%=pCompanyStateCode%> - 
      <%end if%>                  
      <%=pCompanyCountryCode%>: </td>
      <td width="139">                
	<input type="text" name="taxValueSameCountry" value="0.00" size=5>%  (example 5.5%)
      </td>
    </tr>
    <tr> 
      <td width="269">Enter the tax percentage for customers located outside <%=pCompanyCountryCode%>:</td>
      <td width="139">                
        <input type="text" name="taxValueDifferentCountry" value="0.00" size=5>%
      </td>
    </tr>           
    <tr> 
      <td width="269">&nbsp;</td>
      <td width="139">&nbsp;</td>
    </tr>        
    <tr> 
      <td colspan="2"> 
        <br>
        <input type="submit" name="Submit" value="Continue">
      </td>
    </tr>
</table>    
</form>
<br>Note: you can create more tax rules at a later time
<!--#include file="footer.asp"-->
<%call closeDb()%>