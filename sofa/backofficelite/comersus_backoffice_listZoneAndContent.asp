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
mySQL="SELECT shippingZones.zoneName FROM shippingZones WHERE shippingZones.idShippingZone = "&pIdShippingZone

call getFromDatabase(mySQL, rstemp2, "comersus_backoffice_listZonesContents.asp") 

' get all shipment zone contents
mySQL="SELECT shippingZonesContents.*, shippingZones.zoneName FROM shippingZones, shippingZonesContents WHERE shippingZones.idShippingZone = shippingZonesContents.idShippingZone AND shippingZones.idShippingZone = "&pIdShippingZone&" ORDER BY idShippingZonesContents"

call getFromDatabase(mySQL, rstemp, "comersus_backoffice_listZonesContents.asp") 

%>

<!--#include file="header.asp"-->
<!--#include file="./includes/settings.asp"--> 

<form action='comersus_backoffice_shipmentZoneModifyExec.asp'	 method=post>
<input type="hidden" name="idZone" value="<%=pIdShippingZone%>">
<br>Zone Name: <input type="text" name="zoneName" value="<%=rstemp2("zoneName")%>"> <input type="submit" value="Modify">|<input type="button" value="Delete" onclick="location.href='comersus_backoffice_shipmentZoneDeleteExec.asp?idZone=<%=pIdShippingZone%>'"><br><br>
</form>
<hr>
<br>
<%
if not rstemp.eof then%>
<b>Zone Contents</b>
<BR>

  <table border=0 width="100%">
    <tr> 
    	<td><b>Zip</b></td>
    	<td>|</td>
    	<td><b>State</b></td>
    	<td>|</td>
    	<td><b>Country</b></td>
    	<td>|</td>
    	<td><b>Actions</b></td>
    </tr>
    
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
	<tr>   
      <td>
		      <%If pZip="" Then 	
		      		response.write "All"
		      	else
		      		response.write pZip
		      	end if%>
      </td>
      <td>|</td>
      <td>
		      <%If pStateCode="" Then 	
		      		response.write "All"
		      	else
		      		response.write getStateName(pStateCode)
		      	end if%>
      </td>
      <td>|</td>
      <td>
		      <%If pCountryCode="" Then 	
		      		response.write "All"
		      	else
		      		response.write getCountryName(pCountryCode)
		      	end if%>
      </td>
      <td>|</td>
      <td>
      	<a href="comersus_backoffice_modifyZonesContents.asp?idZoneContent=<%=pIdZoneContent%>&idZone=<%=pIdShippingZone%>">Edit</a>&nbsp;<a href="comersus_backoffice_deleteZonesContents.asp?idZone=<%=pIdShippingZone%>&idZoneContent=<%=pIdZoneContent%>">Delete</a>
      </td>
    </tr>
   <%  
   rstemp.MoveNext
 loop%>
   </table>
<%
else
%> 
<p>There are no Zone Contents in this Shipping Zone.<br></p>
<%end if %>
<br><BR>
<hr>
<form action='comersus_backoffice_addZonesContentsExec.asp'	 method=post>
<b>Add Zone Content</b>
	<input type="hidden" name="IdZone" value=<%=pIdShippingZone%>>
   <table border="0" width="100%">
    <tr> 
      <td width="100">Zip</td>
      <td width="450">
		<input type=text name="zip" value="">	
      </td>
    </tr>
    <tr> 
    <td width="100">State</td>
      <td width="450">  
        <%
    
      ' get stateCodes
      mySQL="SELECT * FROM stateCodes"
      
      call getFromDatabase(mySQL, rstemp2, "comersus_backoffice_addShipmentExec.asp") 
      %>
	      <select name=stateCode size=1>
		      <option value="">All</option>
		      <%do while not rstemp2.eof
		      pStateCode2=rstemp2("stateCode")%>
		        <option value="<%=pStateCode2%>"><%=rstemp2("stateName")%>
		      <%rstemp2.movenext
		      loop%>
	  	      	</option>
	      </select>
      </td>      
    </tr>
     <tr> 
      <td width="100">Country</td>
      <td width="450">                 
        <%
      ' get CountryCodes
      mySQL="SELECT * FROM countryCodes"

      call getFromDatabase(mySQL, rstemp2, "comersus_backoffice_addShipmentExec.asp") 
      
      %>
	      <select name=countryCode>
		      <option value="">All</option>
		      <%do while not rstemp2.eof
		      pCountryCode2=rstemp2("countryCode")%>
		        <option value="<%=pCountryCode2%>"><%=rstemp2("countryName")%>
		      <%rstemp2.movenext
		      loop%>
		      </option>
	      </select>       
      </td>      
    </tr>
    <tr> 
      <td colspan=2><i>If you leave zip unset, the rule will work for all zips.</i></td>            
    </tr>      

  </table>
	<br>
	<input type=submit value="Add location to the zone" name=AddZone>
</form>
<hr>
<b>Current shipment methods assigned to this zone</b><br>
<%
' get all shipment zone contents
mySQL="SELECT * FROM shipments WHERE idShippingZone="&pIdShippingZone
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_listZonesContents.asp") 

do while not rstemp.eof
%>
<br><%=rstemp("shipmentDesc")%>&nbsp; <%=pCurrencySign%><%=rstemp("priceToAdd")%>, <%=rstemp("percentageToAdd")%>% <a href="comersus_backoffice_shipmentDeleteExec.asp?idShipment=<%=rstemp("idShipment")%>">Delete</a>
<%rstemp.movenext
loop%>
<hr>
<b>Add new shipment method for this Zone</b>
<br>
<form action='comersus_backoffice_addShipmentForm.asp'	 method=get>
	<input type="hidden" name="IdZone" value=<%=pIdShippingZone%>>
	<input type=submit value="Add Shipment"name=AddShipment>
</form>

<br>
<a href="comersus_backoffice_listShipmentZones.asp">Back</a>
<!--#include file="footer.asp"-->
