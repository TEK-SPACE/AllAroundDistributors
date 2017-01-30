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
on error resume next 

dim mySQL, conntemp, rstemp

pRunInstallationWizard 	= getSettingKey("pRunInstallationWizard")
pCompanyCountryCode	= getSettingKey("pCompanyCountryCode")
pCurrencySign		= getSettingKey("pCurrencySign")

if pRunInstallationWizard<>"-1" then response.redirect "comersus_backoffice_message.asp?message="&Server.UrlEncode("Sorry, the Installation Wizard is disabled")

pPayPalMerchantEmail		= getUserInput(request.form("pPayPalMerchantEmail"),50)
p2CheckoutSID			= getUserInput(request.form("p2CheckoutSID"),50)
pUseOffLineCreditCards  	= getUserInput(request.form("pUseOffLineCreditCards"),5)
pChecks				= getUserInput(request.form("pChecks"),5)
pChecksComments			= getUserInput(request.form("pChecksComments"),100)
pWireTransfer			= getUserInput(request.form("pWireTransfer"),5)
pWireComments			= getUserInput(request.form("pWireComments"),100)

if pPayPalMerchantEmail<>"" then
  mySQL1="UPDATE settings SET settingValue='" &pPayPalMerchantEmail& "' WHERE settingKey='pPayPalMerchantEmail' AND idStore=" &pIdStore
  call updateDatabase(mySQL1, rstemp, "comersus_backoffice_install.asp") 
  
  mySQL1="INSERT INTO payments (paymentDesc, redirectionUrl, quantityUntil, weightUntil, priceUntil, idCustomerType, idStore) VALUES ('PayPal','comersus_gatewayPayPal.asp',9999,9999,999999,NULL," &pIdStore& ")"
  call updateDatabase(mySQL1, rstemp, "comersus_backoffice_install.asp") 
  
end if

if p2CheckoutSID<>"" then
  mySQL1="UPDATE settings SET settingValue='" &p2CheckoutSID& "' WHERE settingKey='p2CheckoutSID' AND idStore=" &pIdStore
  call updateDatabase(mySQL1, rstemp, "comersus_backoffice_install.asp") 
  
  mySQL1="INSERT INTO payments (paymentDesc, redirectionUrl, quantityUntil, weightUntil, priceUntil, idCustomerType, idStore) VALUES ('Credit Cards with 2Checkout','comersus_gateway2CheckOut.asp',9999,9999,999999,NULL,"&pIdStore&")"
  call updateDatabase(mySQL1, rstemp, "comersus_backoffice_install.asp") 
end if

' other settings

if pUseOffLineCreditCards="-1" then

 ' check if the payment method is already inserted
 mySQL1="SELECT idPayment FROM payments WHERE redirectionURL='comersus_offLinePaymentForm.asp' AND idStore=" &pIdStore
 call getFromDatabase(mySQL1, rstemp, "comersus_backoffice_install.asp") 
 
 if rstemp.eof then
  ' the payment method is not inserted
  mySQL1="INSERT INTO payments (paymentDesc, redirectionUrl, quantityUntil, weightUntil, priceUntil, idCustomerType, idStore) VALUES ('Credit Cards','comersus_offLinePaymentForm.asp',9999,9999,999999,NULL,"&pIdStore&")"
  call updateDatabase(mySQL1, rstemp, "comersus_backoffice_install.asp") 
 end if
 
 ' cards accepted
 pOffLineCardsAccepted=""
 if request.form("vi")="-1" then pOffLineCardsAccepted=pOffLineCardsAccepted&"VI, "
 if request.form("ma")="-1" then pOffLineCardsAccepted=pOffLineCardsAccepted&"MA, "
 if request.form("am")="-1" then pOffLineCardsAccepted=pOffLineCardsAccepted&"AM, "
 if request.form("din")="-1" then pOffLineCardsAccepted=pOffLineCardsAccepted&"DIN, "
 if request.form("dis")="-1" then pOffLineCardsAccepted=pOffLineCardsAccepted&"DIS, "
 if request.form("jc")="-1" then pOffLineCardsAccepted=pOffLineCardsAccepted&"JC, "
 
 if pOffLineCardsAccepted<>"" then
  mySQL1="UPDATE settings SET settingValue='" &pOffLineCardsAccepted& "' WHERE settingKey='pOffLineCardsAccepted' AND idStore=" &pIdStore
  call updateDatabase(mySQL1, rstemp, "comersus_backoffice_install.asp") 
 end if

end if

if pChecks<>"" and pChecksComments<>"" then
  mySQL1="INSERT INTO payments (paymentDesc, emailText, quantityUntil, weightUntil, priceUntil, idCustomerType, idStore) VALUES ('Checks','" &pChecksComments& "',9999,9999,999999,NULL,"&pIdStore&")"
  call updateDatabase(mySQL1, rstemp, "comersus_backoffice_install.asp") 
end if

if pWireTransfer<>"" and pWireComments<>"" then
  mySQL1="INSERT INTO payments (paymentDesc, emailText, quantityUntil, weightUntil, priceUntil, idCustomerType, idStore) VALUES ('Wire Transfer','" &pWireComments& "',9999,9999,999999,NULL, "&pIdStore&")"
  call updateDatabase(mySQL1, rstemp, "comersus_backoffice_install.asp") 
end if

%>
<!--#include file="header.asp"--> 
<!--#include file="./includes/settings.asp"-->

<br><font size="3"><b>Installation Wizard</b></font><br>

<br>Step 5 - Shipping Zones<br><br>
<form method="post" name="install5a" action="comersus_backoffice_install6.asp">
<table width="418" border="0">

    <tr> 
      <td colspan=2>A. By pass shipping calculations</b></td>
    </tr>
    
    
     <tr> 
      <td width="189">Select this option for digital goods and service oriented stores</td>
      <td width="219">        
        <select name="pByPassShipping">
        <option value='-1'>Yes</option>
        <option value='0' selected>No</option>
        </select>
      </td>
    </tr>
    
    <tr> 
      <td colspan=2>B. Do not by pass shipping calculations</b></td>
    </tr>
    
    <tr> 
      <td width="189">Charge this fixed fee for  <%=pCompanyCountryCode%> </td>
      <td width="219">
       <%=pCurrencySign%> <input type=text name="pShippingAmount1" value="0.00">        
      </td>
    </tr>
    
    
    <tr> 
      <td width="189">Charge this fixed fee for all other countries</td>
      <td width="219">
        <%=pCurrencySign%> <input type=text name="pShippingAmount2" value="0.00">        
      </td>
    </tr>
    
    <tr> 
      <td colspan="2"> 
        <br>
        <input type="submit" name="Submit" value="Continue">
      </td>
    </tr>
</table>    
</form>

<br>Additional information 
<br>(1) Here you can configure a basic shipping fee to be charged. You can set detailed shipping rules later.
<!--#include file="footer.asp"-->
<%call closeDb()%>