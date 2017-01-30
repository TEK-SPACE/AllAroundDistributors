<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"-->   
<!--#include file="../includes/databaseFunctions.asp"-->   
<!--#include file="../includes/settings.asp" --> 
<!--#include file="../includes/getSettingKey.asp"-->  

<%
on error resume next

dim mySQL, conntemp, rstemp

' get settings 
pCurrencySign	 	= getSettingKey("pCurrencySign")

%>

<!--#include file="header.asp"-->
<!--#include file="./includes/settings.asp"-->

<br><b>Add payment</b><br>
<form method="post" name="AddPay" action="comersus_backoffice_addPaymentExec.asp">
<table width="550" border="0">
  <tr> 
    <td width="150">Description</td>
    <td width="400">         
        <input type=text name=paymentDesc>        
    </td>
  </tr>
</table>
<table width="550" border="0">
  <tr> 
    <td width="140">Quantity</td>
    <td width="150">From
      <input type=text name=quantityFrom value=0>      
    </td>
    <td width="140" height="29">Until
      <input type=text name=quantityUntil value=9999>      
    </td>
  </tr>
  <tr> 
    <td width="140">Cart total</td>
    <td width="150">From
      <input type=text name=priceFrom value=0.00>      
    </td>
    <td width="140">Until
      <input type=text name=priceUntil value=999999.00>      
    </td>
  </tr>
</table>
<table width="550" border="0">
  <tr> 
    <td width="170">Fixed price <%=pCurrencySign%>:</td>
    <td width="80">
      <input type=text name=priceToAdd value=0.00>      
    </td>
    <td width="220">Percentage</td>
    <td width="100">          
      <input type=text name=percentageToAdd value=0>      
    </td>
  </tr>
  <tr> 
    <td width="170">&nbsp;</td>
    <td width="80">
      &nbsp;
    </td>
    <td width="220">Redirection URL</td>
    <td width="100"> 
      <input type="text" name="redirectionUrl" value="">
    </td>
  </tr>
</table>
<table width="550" border="0">
  <tr> 
    <td width="150" valign="top">Information to distribute (bank info for Wire Transfer, etc)</td>
    <td width="400"> 
        <textarea name="EmailText" rows="5" cols="50"></textarea>
    </td>
  </tr>
  <tr>
    <td width="150">Customer type</td>
    <td width="100"> 
      <select name="customerType">
	<option selected value="null">Select</option>
	<% 	
	mySQL="SELECT idCustomerType, customerTypeDesc FROM customerTypes"
	
	call getFromDatabase(mySQL, rstemp, "comersus_backoffice_addPaymentForm.asp") 
	
	do while not rstemp.eof %>
	<option value=<%=rstemp("idCustomerType")%>><%=rstemp("customerTypeDesc")%></option>
	<%		
		rstemp.movenext
	loop 		
	call closeDb()
	%>	
        </select>
    </td>
  </tr>
  <tr> 
    <td> 
      <input type="submit" name="Submit" value="Add payment">
    </td>
  </tr>
</table>
</form>
<!--#include file="footer.asp"-->

