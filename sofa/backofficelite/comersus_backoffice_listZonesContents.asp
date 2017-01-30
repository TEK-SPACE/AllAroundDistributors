<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"-->  
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/currencyFormat.asp"-->    
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/getSettingKey.asp"-->  
<!--#include file="../includes/stringFunctions.asp" --> 
<!--#include file="../includes/MiscFunctions.asp" --> 

<%
on error resume next

dim mySQL, conntemp, rstemp

pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")

pIdShippingZone = getUserInput(request.querystring("idZone"),20)


' get all shipment zone contents
mySQL="SELECT shippingZonesContents.*, shippingZones.zoneName FROM shippingZones, shippingZonesContents WHERE shippingZones.idShippingZone = shippingZonesContents.idShippingZone AND shippingZones.idShippingZone = "&pIdShippingZone&" ORDER BY idShippingZonesContents"

call getFromDatabase(mySQL, rstemp, "comersus_backoffice_listZonesContents.asp") 

%>

<!--#include file="header.asp"-->
<!--#include file="./includes/settings.asp"--> 

<%
if not rstemp.eof then%>
<br><b><%=rstemp("zoneName")%> Contents</b><br><br>
<%
 do while not rstemp.eof

   pIdZoneContent		= rstemp("idShippingZonesContents")
   if not isnull(rstemp("zip")) Then
   		pZip		= rstemp("zip")
   else
   	  pZip		= ""
   end if
   if not isnull(rstemp("stateCode")) Then
   		pStateCode		= rstemp("stateCode")
   else
   	  pStateCode		= ""
   end if
   if not isnull(rstemp("countryCode")) Then
   		pCountryCode		= rstemp("countryCode")
   else
   	  pCountryCode		= ""
   end if
   %> 
   <input type=hidden name="idZone" value=<%=pIdShippingZone%>>	
   
  <table border=0 width="100%">
    <tr> 
      <td width="100"><b>Zip Code</b></td>
      <td width="400">
		      <%If pZip="" Then 	
		      		response.write "All"
		      	else
		      		response.write pZip
		      	end if%>
      </td>
    </tr>
    <tr> 
      <td width="100"><b>State</b></td>
      <td width="400">
		      <%If pStateCode="" Then 	
		      		response.write "All"
		      	else
		      		response.write getStateName(pStateCode)
		      	end if%>
      </td>
    </tr>
    <tr> 
      <td width="100"><b>Country Code</b></td>
      <td width="400">
		      <%If pCountryCode="" Then 	
		      		response.write "All"
		      	else
		      		response.write getCountryName(pCountryCode)
		      	end if%>
      </td>
    </tr>
    <tr> 
      <td colspan="2">
      	<a href="comersus_backoffice_modifyZonesContents.asp?idZoneContent=<%=pIdZoneContent%>&idZone=<%=pIdShippingZone%>">Modify</a>&nbsp;<a href="comersus_backoffice_deleteZonesContents.asp?idZone=<%=pIdShippingZone%>&idZoneContent=<%=pIdZoneContent%>">Delete</a>
      </td>
    </tr>

  </table>
      <hr>
   <%  
   rstemp.MoveNext
 loop
else
%> 
<p>There are no Zone Contents in this Shipping Zone.<br></p>
<%end if %>
<div align="center"><input type="button" value="Add Zone Content" onclick="location.href='comersus_backoffice_addZonesContents.asp?idZone=<%=pIdShippingZone%>'"></div>
<br>
<a href="comersus_backoffice_listShipmentZones.asp">Back</a>
<!--#include file="footer.asp"-->
