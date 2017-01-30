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

<%
on error resume next

dim mySQL, conntemp, rstemp

pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")


pIdZoneContent	= getUserInput(request("idZoneContent"),100)
pIdZone			= getUserInput(request("idZone"),100)

If getUserInput(request.form("submit"),20) <> "" Then
	'update shipping zone
	pZip = getUserInput(request("zip"),50)
	pStateCode = getUserInput(request("statecode"),10)
	pCountryCode = getUserInput(request("CountryCode"),10)
	
	MySQL	= "UPDATE shippingZonesContents SET zip='"&pZip&"', stateCode='"&pStateCode&"', countryCode='"&pCountryCode&"' WHERE idShippingZonesContents=" & pIdZoneContent
	call updateDatabase(mySQL, rstemp, "comersus_backoffice_modifyZonesContents.asp") 
	response.redirect "comersus_backoffice_listZoneAndContent.asp?IdZone="&pIdZone
end if


' get shipment
mySQL="SELECT * FROM shippingZonesContents WHERE idShippingZonesContents = " & pIdZoneContent

call getFromDatabase(mySQL, rstemp, "comersus_backoffice_modifyZoneContents.asp") 

if  rstemp.eof then 
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Couldn't find the zone in database")
end if  

%>

<!--#include file="header.asp"-->
<!--#include file="includes/settings.asp"--> 

<form action='<%=request.ServerVariables("SCRIPT_NAME")%>'	 method=post>
<br><b>Zone Contents</b><br><br>
<%

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
   <input type=hidden name="idZoneContent" value=<%=pIdZoneContent%>>	
   <input type=hidden name="idZone" value=<%=pIdZone%>>	
   <table border="0" width="100%">
    <tr> 
      <td width="100">Zip</td>
      <td width="450">
		<input type=text name="zip" value="<%=pZip%>">	
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
		        <option value="<%=pStateCode2%>" <%If pStateCode2 = pStateCode Then response.write "selected" end if%>><%=rstemp2("stateName")%>
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
		        <option value="<%=pCountryCode2%>"<%
		        if pCountryCode=pCountryCode2 then
		            response.write "selected"           
		        end if
		        %>><%=rstemp2("countryName")%>
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
	<input type=submit value=Modify name=submit>
</form>
<!--#include file="footer.asp"-->