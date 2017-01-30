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

<%
on error resume next

dim mySQL, conntemp, rstemp

pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")

' get all shipments
mySQL="SELECT * FROM shippingZones ORDER BY zoneName"

call getFromDatabase(mySQL, rstemp, "comersus_backoffice_listShipmentZones.asp") 

if  rstemp.eof then 
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("No zones in database")
end if  

%>

<!--#include file="header.asp"-->
<!--#include file="./includes/settings.asp"--> 

<br><b>Zones</b><br><br>
<%
 do while not rstemp.eof

   pIdZone			= rstemp("idShippingZone")
   pZoneName		= rstemp("zoneName")
  
   %> 
   <input type=hidden name="idZone" value=<%=pIdZone%>>	
   
  <table border=0 width="100%">
    <tr> 
      <td width="80"><b><%=pZoneName%></b></td>
      <td width="430">
		   <a href="comersus_backoffice_listZoneAndContent.asp?idZone=<%=pIdZone%>">Edit & Manage Zones and Contents</a>&nbsp;<a href="comersus_backoffice_shipmentZoneDeleteExec.asp?idZone=<%=pIdZone%>">Delete</a>
      </td>
    </tr>

  </table>
      <hr>
   <%  
   rstemp.MoveNext
 loop

%> 
<!--#include file="footer.asp"-->
