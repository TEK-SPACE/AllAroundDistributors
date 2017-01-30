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
<!--#include file="../includes/stringFunctions.asp" --> 
<%
on error resume next

dim mySQL, conntemp, rstemp

' get settings 
pCurrencySign	 	= getSettingKey("pCurrencySign")

pShipmentZone 		= getUserInput(request.querystring("idZone"),20)

if pShipmentZone = "" then 
 pShipmentZone = 0
end if

%>
<!--#include file="./includes/settings.asp"-->
<%%>

<!--#include file="header.asp"-->

<br><b>Add shipment</b><br>
<form method="post" name="addSh" action="comersus_backoffice_addShipmentExec.asp">
  <table width="550" border="0">
    <tr> 
      <td width="150">Description</td>
      <td width="400">
        <input type=text name="shipmentDesc">        
      </td>
    </tr>
    
    <tr> 
      <td colspan=2>This shipment method will be used only of the following conditions are true:</td>      
    </tr>
  </table>
  <table width="550" border="0">
  
  <tr> 
      <td width="140">Customer type</td>
      <td width="160"><select name="customerType">
		<option selected value="null">Select</option>
		<%
	 	
		mySQL="SELECT idCustomerType, customerTypeDesc FROM customerTypes"
		
		call getFromDatabase(mySQL, rstemp, "comersus_backoffice_addShipmentForm.asp") 

		do while not rstemp.eof %>
			<option value=<%=rstemp("idCustomerType")%>><%=rstemp("customerTypeDesc")%></option>
			<%		
			rstemp.movenext
		loop 		
		%>
	</select>
      </td>
      <td width="130">&nbsp;        
      </td>
    </tr>      
	
    <tr> 
      <td width="140">Quantity</td>
      <td width="160">From
        <input type=text name="quantityFrom" value=0>
      </td>
      <td width="130">Until 
        <input type=text name="quantityUntil" value=9999>
      </td>
    </tr>
    <tr> 
      <td width="140">Weight</td>
      <td width="160">From
        <input type=text name="WeightFrom" value=0>
      </td>
      <td width="130">Until
        <input type=text name="WeightUntil" value=9999>
      </td>
    </tr>
    <tr> 
      <td width="140">Cart total</td>
      <td width="160">From
        <input type=text name="priceFrom" value=0.00>
      </td>
      <td width="130">Until
        <input type=text name="priceUntil" value=99999>
      </td>
    </tr>
  </table>
  <table width="550" border="0">
   <tr> 
      <td colspan=4 >&nbsp;</td>                              
    </tr>
    
    <tr> 
      <td width="150">Fixed price <%=pCurrencySign%>:</td>
      <td width="90">         
        <input type=text name="priceToAdd" value=0.00>
      </td>
      <td width="190">Percentage</td>
      <td width="70">          
        <input type=text name="percentageToAdd" value=0>
      </td>
    </tr>
    <tr> 
      <td width="150">Shipment time</td>
      <td width="90">
        <input type="text" name="shipmentTime" size="15" value="Undetermined">
      </td>
      <td width="190"></td>
      <td width="70">         
      </td>
    </tr>
  </table>
  <table width="550" border="0">
  
    <tr> 
      <td colspan=3>&nbsp;</td>            
    </tr>
    <tr> 
      <td colspan=3>This shipment method will be used if the customer address is in this Zone:</td>            
    </tr>
    <tr>
      <td width="150">Shipment Zone</td>
      <td width="90">
      	<select name="shippingZone">
    	<%
    		MySql = "SELECT * FROM shippingZones ORDER BY zoneName"
			call getFromDatabase(mySQL, rstemp, "comersus_backoffice_addShipmentForm.asp") 
			Do While not rstemp.eof
	    	%>
	    		<option value="<%=rstemp("idShippingZone")%>" <%If cint(pShipmentZone) = cInt(rstemp("idShippingZone")) Then response.write "selected"%>> <%=rstemp("zoneName")%></option>
	    	<%
	    		rstemp.movenext
    		loop
    	%>
    	</select>
      </td>
      <td width="190"></td>
      <td width="70">         
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
