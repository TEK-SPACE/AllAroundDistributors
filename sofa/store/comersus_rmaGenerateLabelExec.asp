<%
' Comersus BackOffice 
' e-commerce ASP Open Source
' CopyRight Comersus Open Technologies LC
' United States - 2005
' http://www.comersus.com 
%>


<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/writeBarcode128B.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->   
<!--#include file="../includes/settings.asp"-->    
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/screenMessages.asp"-->
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/cartFunctions.asp"--> 
<!--#include file="../includes/adSenseFunctions.asp"-->

<% 

' on error resume next 

dim mySQL, conntemp, rstemp, rstemp2, rstemp3

' get settings 
pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanyLogo	 	= getSettingKey("pCompanyLogo")
pCompanyPhone	 	= getSettingKey("pCompanyPhone")
pCompanyAddress	 	= getSettingKey("pCompanyAddress")
pCompanyStateCode 	= getSettingKey("pCompanyStateCode")
pCompanyCity		= getSettingKey("pCompanyCity")
pCompanyZip	 	= getSettingKey("pCompanyZip")
pCompanyCountryCode 	= getSettingKey("pCompanyCountryCode")
pOrderPrefix	 	= getSettingKey("pOrderPrefix")
pIdRmaPrefix	 	= getSettingKey("pIdRmaPrefix")
pIdRma 			= getUserInput(request.querystring("idRma"),20)

pIdCustomer     	= getSessionVariable("idCustomer",0)
pIdCustomerType 	= getSessionVariable("idCustomerType",1)
pIdDbSessionCart	= getSessionVariable("idDbSessionCart",0)


if trim(pIdRma)="" then
  response.redirect "backoffice_message.asp?message="&Server.Urlencode(getMsg(743,"Invalid RMA"))
end if

pIdRma 	= removePrefix(pIdRma, pIdRmaPrefix)

' check if the RMA belongs to the customer
mySQL="SELECT orders.idOrder FROM orders, RMA WHERE orders.idOrder = RMA.idOrder AND RMA.idRma= " & pIdRma & " AND orders.idCustomer="& pIdCustomer
call getFromDatabase(mySQL, rstemp, "comersus_rmaGenerateLabelExec.asp") 

if rstemp.eof <> 0 then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(743,"Invalid RMA"))
end if

' retrieve order data
mySQL="SELECT orders.idorder, name, lastName, phone, email, orders.address AS address, orders.state AS state, orders.StateCode, orders.zip AS zip, orders.city AS city, orders.countryCode AS countryCode, orders.shippingAddress, orders.shippingState, orders.shippingStateCode, orders.shippingZip, orders.shippingCity, orders.shippingCountryCode, taxAmount, orderDate, ShipmentDetails, paymentDetails, obs, total, details, discountDetails FROM orders, customers, RMA WHERE orders.idOrder = RMA.idOrder AND RMA.idRma= " & pIdRma & " AND orders.idcustomer=customers.idcustomer" 
call getFromDatabase(mySQL, rstemp, "comersus_rmaGenerateLabelExec.asp") 

if rstemp.eof <> 0 then
	response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(743,"Invalid RMA"))
end if


   ' contact
   
   pName		= rstemp("name")
   pLastName		= rstemp("lastName")
   pEmail		= rstemp("email")
   pPhone		= rstemp("phone")
      
   ' order
   pOrderDate		= rstemp("orderDate")   
   
   pAddress		= rstemp("address")
   pZip			= rstemp("zip")
   pState		= rstemp("state")
   pStateCode		= rstemp("stateCode")
   pCity		= rstemp("city")
   pCountryCode		= rstemp("countryCode")
   
   pShippingAddress	= rstemp("shippingAddress")
   pShippingZip		= rstemp("shippingZip")
   pShippingState	= rstemp("shippingState")
   pShippingStateCode	= rstemp("shippingStateCode")
   pShippingCity	= rstemp("shippingCity")
   pShippingCountryCode	= rstemp("shippingCountryCode")   

call closeDb()
%>

<HTML>
<title> <%=pCompany%> >> Shipping Label </title>


  <p>&nbsp; </p>
<table width="300" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td> 
      <div align="center">- - - - - - - - - - - - - - - - - - - - - - - - </div>
    </td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>  
  <tr> 
    <td> 
 
        <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>From  
        <%=pName&" "&pLastName%></b></font></div>
    </td>
  </tr>
  <tr> 
    <td> 
      <div align="center">
      <font face="Verdana, Arial, Helvetica, sans-serif" size="1"><i>RMA #<%=pIdRmaPrefix&pIdRma%> Approved</i></font><br><br>
    </td>
  </tr>
  <%if pShippingAddress<>"" then%> 
  <tr> 
    <td> 
      <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=pShippingAddress%></font></div>
    </td>
  </tr>
  <tr> 
    <td> 
      <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=pShippingStateCode&pShippingState&" - " &pShippingCity& " ("&pShippingZip&")"%></font></div>
    </td>
  </tr>
  <tr> 
    <td> 
      <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=pShippingCountryCode%></font></div>
    </td>
  </tr>
  <%else%> 
  <tr> 
    <td> 
      <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=pAddress%></font></div>
    </td>
  </tr>
  <tr> 
    <td> 
      <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=pStateCode&pState&" - " &pCity& " ("&pZip&")"%></font></div>
    </td>
  </tr>
  <tr> 
    <td> 
      <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=pCountryCode%></font></div>
    </td>
  </tr>    
  <%end if%> 
  <tr> 
    <td>&nbsp;</td>
  </tr>
    <%
  pZipBar=pZip
  if pShippingZip<>"" then pZipBar=pShippingZip
  pEncodedBar=pIdRmaPrefix&pIdRma&"|"&pZipBar&"|"&pCompanyZip
  %>
  
  <tr> 
    <td> 
      <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%call writeBarcode128B(pEncodedBar)%><br><%=pEncodedBar%></font></div>
    </td>
  </tr>
  
  <tr> 
    <td> 
      <div align="center">&nbsp</div>
    </td>
  </tr>  
  <tr> 
    <td> 
      <div align="center">- - - - - - - - - - - - - - - - - - - - - - - - </div>
    </td>
  </tr>
  <tr> 
    <td> 
      <div align="center">&nbsp</div>
    </td>
  </tr>
  <tr> 
    <td> 
       <div align="center">      
   
      <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>To  
        <%=pCompany%></b></font></div>
    </td>
  </tr>
  <tr> 
    <td> 
      <div align="center">
      <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=pCompanyAddress%></font><br>
      <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=pCompanyStateCode&" - "&pCompanyCity& " ("&pCompanyZip&")"%></font></div>
    </td>
  </tr>
  <tr> 
    <td> 
      <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=pCompanyCountryCode%></font></div>
    </td>
  </tr>
  <tr> 
    <td> 
      &nbsp;
    </td>
  </tr>  
  <tr> 
    <td> 
      <div align="center">- - - - - - - - - - - - - - - - - - - - - - - - </div>
    </td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td> 
      <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <SCRIPT LANGUAGE="JavaScript">

<!-- Begin
if (window.print) {
document.write('<form>'
+ '<input type=button name=print value="Print the Label" '
+ 'onClick="javascript:window.print()"> </form>');
}
// End -->
</script>
        </font></div>
    </td>
  </tr>
</table>

<br>
</HTML>

