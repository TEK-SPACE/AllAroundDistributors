<%
' Comersus BackOffice 
' e-commerce ASP Open Source
' CopyRight Rodrigo S. Alhadeff, Comersus
' 2004
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"-->    
<!--#include file="../includes/databaseFunctions.asp"-->  
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->   
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/settings.asp"--> 

<!--#include file="header.asp"--> 
<br><b>Import Customers</b>
<br><br>
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
oldvalue = "";
function passText(passedvalue) {
  if (passedvalue != "") {
       	if (oldvalue != "") {	  
    		var totalvalue = oldvalue+"\n"+passedvalue;
  	}else{  
    		var totalvalue = passedvalue;
  	}      
    document.selectform.itemsbox.value = totalvalue;
    oldvalue = document.selectform.itemsbox.value;
  }
}
//  End -->
</script>
<%
on error resume next

dim mySQL, conntemp, rstemp
mySQL="SELECT name, lastName, customerCompany, phone, email, address, zip, stateCode, state, city, countryCode, shippingName, shippingLastName, shippingaddress, shippingcity, shippingStateCode, shippingState, shippingCountryCode, shippingZip FROM Customers"
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_importUtilityForm.asp") 

'check if the form was loaded first

if request("itemsbox")="" then
	%>
	<br><br>1. Select the fields to import 
	<br>2. Create a CSV text file customers.txt in BackOffice folder respecting the fields order and separated with ,
	<br>3. Click on Import button
	<form name="selectform" method="post" action="comersus_backoffice_importCustomersForm.asp">
	<table>
	<tr>
		<td>Available fields</td>
		<td><select name="CustomerFields" size=1>
		<%for each whatever in rstemp.fields%>
			<option value="<%=whatever.name%>"><%=whatever.name%></option>
		<%next%>	
		</select>
		<input type=button value="Add to list" onClick="passText(this.form.CustomerFields.options[this.form.CustomerFields.selectedIndex].value);">
		</td>  
	</tr>
	<tr>
		<td>Fields to import<br><br></td>
		<td><textarea cols="15" rows="8" name="itemsbox" ></textarea></td>
	</tr>
	
	<tr>
		<td>&nbsp;&nbsp;&nbsp;</td>
		<td>		
			<input type=submit value="Import">
		</td>
	</tr>
	
	
	</table>   
	</form>   
	
<%else%>
 <br><br>Importing /backofficelite/customers.txt with record format: 
 <%
 pFields	= replace(replace(request("itemsbox"),chr(13),","),chr(10),"")
 do while instr(pFields,",")<>0	
	pIndexField=mid(pFields,1,instr(pFields,",")-1)
	pFields=mid(pFields,instr(pFields,",")+1)							
	for each whatever in rstemp.fields						
		if whatever.name=rtrim(ltrim(pIndexField)) then
			'if (whatever.type=202 or whatever.type=203) then
			'	response.write "'"&pIndexField&"',"
			'else
				response.write pIndexField&","				
			'end if
		else
		
		end if
	next
 loop
	pIndexField=trim(pFields)				
	pFields	= replace(replace(request("itemsbox"),chr(13),","),chr(10),"")
	for each whatever in rstemp.fields						
		if whatever.name=rtrim(ltrim(pIndexField)) then
			if (whatever.type=202 or whatever.type=205) then
				response.write "'"&pIndexField&"'"
			else
				response.write pIndexField				
			end if
		else
		
		end if
	next	

%>
  <br><br>If you are ready to begin <a href="comersus_backoffice_importCustomersExec.asp?fields=<%=pFields%>&includeCategory=<%=request("includeCategory")%>">click here...</a><br>  
<%end if%>
<!--#include file="footer.asp"--> 
<%call closeDb()%>