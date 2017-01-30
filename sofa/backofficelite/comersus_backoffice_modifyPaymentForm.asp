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
<!--#include file="../includes/stringFunctions.asp"-->  

<%
on error resume next

dim mySQL, conntemp, rstemp

' get settings 
pCurrencySign	 	= getSettingKey("pCurrencySign")

' form parameter 
pIdPayment		= getUserInput(request.querystring("idPayment"),12)

if trim(pIdPayment)="" then
   response.redirect "comersus_backoffice_message.asp?message="&Server.Urlencode("You must specify payment #.")
end if

' get payment details from db
mySQL="SELECT paymentDesc, priceToAdd, percentageToAdd, redirectionUrl, emailText, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil, idCustomerType FROM payments WHERE idPayment=" &pIdPayment
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_modifyProductForm.asp")  	

if rstemp.eof then
   response.redirect "comersus_backoffice_message.asp?message="&Server.Urlencode("Couldn't find payment in database.")
end if

pPaymentDesc 			= rstemp("paymentDesc")
pPriceToAdd				= rstemp("priceToAdd")
pPercentageToAdd	= rstemp("percentageToAdd")
pRedirectionUrl		= rstemp("redirectionUrl")
pEmailText				= rstemp("emailText")
pQuantityFrom			= rstemp("quantityFrom")
pQuantityUntil		= rstemp("quantityUntil")
pWeightFrom				= rstemp("weightFrom")
pWeightUntil			= rstemp("weightUntil")
pPriceFrom				= rstemp("priceFrom")
pPriceUntil				= rstemp("priceUntil")
pIdCustomerType		= rstemp("idCustomerType")

%>

<!--#include file="header.asp"-->
<!--#include file="./includes/settings.asp"-->

<br><b>Modify payment</b><br>
<form method="post" name="ModPay" action="comersus_backoffice_modifyPaymentExec.asp">
<input type="hidden" name=idPayment value="<%=pIdPayment%>">
<table width="550" border="0">
  <tr> 
    <td width="150">Description</td>
    <td width="400">         
        <input type=text name=paymentDesc value="<%=pPaymentDesc%>">        
    </td>
  </tr>
</table>
<table width="550" border="0">
  <tr> 
    <td width="140">Quantity</td>
    <td width="150">From
      <input type=text name=quantityFrom value="<%=pQuantityFrom%>">      
    </td>
    <td width="140" height="29">Until
      <input type=text name=quantityUntil value="<%=pQuantityUntil%>">      
    </td>
  </tr>
  <tr> 
    <td width="140">Cart total</td>
    <td width="150">From
      <input type=text name=priceFrom value="<%=pPriceFrom%>">      
    </td>
    <td width="140">Until
      <input type=text name=priceUntil value="<%=pPriceUntil%>">      
    </td>
  </tr>
</table>
<table width="550" border="0">
  <tr> 
    <td width="170">Fixed price <%=pCurrencySign%>:</td>
    <td width="80">
      <input type=text name=priceToAdd value="<%=pPriceToAdd%>">      
    </td>
    <td width="220">Percentage</td>
    <td width="100">          
      <input type=text name=percentageToAdd value="<%=pPercentageToAdd%>">      
    </td>
  </tr>
  <tr> 
    <td width="170">&nbsp;</td>
    <td width="80">
      &nbsp;
    </td>
    <td width="220">Redirection URL</td>
    <td width="100"> 
      <input type="text" name="redirectionUrl" value="<%=pRedirectionUrl%>">
    </td>
  </tr>
</table>
<table width="550" border="0">
  <tr> 
    <td width="150" valign="top">Information to distribute (bank info for Wire Transfer, etc)</td>
    <td width="400"> 
        <textarea name="EmailText" rows="5" cols="50"><%=pEmailText%></textarea>
    </td>
  </tr>
  <tr>
    <td width="150">Customer type</td>
    <td width="100"> 
      <select name="customerType">
	<option value="null">Select</option>
	<% 	
	mySQL="SELECT idCustomerType, customerTypeDesc FROM customerTypes"
	
	call getFromDatabase(mySQL, rstemp, "comersus_backoffice_addPaymentForm.asp") 
	
	do while not rstemp.eof %>
	<option value="<%=rstemp("idCustomerType")%>"  <%If cstr(pIdCustomerType) = cstr(rstemp("idCustomerType")) Then response.write "selected"%>><%=rstemp("customerTypeDesc")%></option>
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
      <input type="submit" name="Submit" value="Modify payment">
    </td>
  </tr>
</table>
</form>
<!--#include file="footer.asp"-->

