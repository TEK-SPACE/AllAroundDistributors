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

<% 
on error resume next 

dim mySQL, conntemp, rstemp

pRunInstallationWizard 	= getSettingKey("pRunInstallationWizard")

if pRunInstallationWizard<>"-1" then 
 response.redirect "comersus_backoffice_message.asp?message="&Server.UrlEncode("Sorry, the Installation Wizard is disabled")
end if

' save settings into db

for each field in request.form 
 fieldValue = request.form(field)
   
 if fieldValue<>"" and field<>"Submit" then
  mySQL1="UPDATE settings SET settingValue='" &fieldValue& "' WHERE settingKey='" &field& "' AND idStore=" &pIdStore
  call updateDatabase(mySQL1, rstemp, "comersus_backoffice_install.asp") 
 end if
 
next 
%>
<!--#include file="header.asp"--> 
<!--#include file="./includes/settings.asp"-->

<br><font size="3"><b>Installation Wizard</b></font><br>
<br>Step 4 - Payments<br><br>
<form method="post" name="install4" action="comersus_backoffice_install5.asp">
<table width="418" border="0">

   <tr> 
      <td width="189">PayPal Merchant Email</td>
      <td width="219">              
        <input type="text" name="pPayPalMerchantEmail" value="">  
        <br><A HREF="https://www.paypal.com/us/mrb/pal=NNLGPNK82TUS6" target="_blank"><img src="images/paypalsignup.gif" border=0 alt="Accept credit cards and bank transfers with PayPal Now"></a>                               
      </td>
    </tr>    
    
    <tr> 
      <td width="189">2Checkout Store ID (1)</td>
      <td width="219">          
       <input type="text" name="p2CheckoutSID" value="">     
       <br><a href="http://www.2checkout.com/cgi-bin/aff.2c?affid=55003" target="_blank"><img src="images/2cosignup.gif" border=0 alt="Accept credit cards with 2Checkout Now"></a>          
      </td>
    </tr>
    <tr> 
      <td width="189">&nbsp;</td>
      <td width="219">&nbsp;</td>
    </tr>
    <tr> 
      <td width="189">Use Off Line Credit Cards</td>
      <td width="219">        
        <select name="pUseOffLineCreditCards">
        <option value='-1'>Yes</option>
        <option value='0'>No</option>
        </select>
      </td>
    </tr>    
    <tr> 
      <td width="189">Cards accepted</td>
      <td width="219">        
        Visa <input type="checkbox" name="vi" value=-1 checked> MasterCard <input type="checkbox" name="ma" value=-1 checked> Amex <input type="checkbox" name="am" value=-1 checked> Diners <input type="checkbox" name="din" value=-1 checked> Discover <input type="checkbox" name="dis" value=-1> JCB <input type="checkbox" name="jc" value=-1>
      </td>
    </tr>
    <tr> 
      <td width="189">&nbsp;</td>
      <td width="219">&nbsp;</td>
    </tr>
    <tr> 
      <td width="189">Accept Off Line checks</td>
      <td width="219">        
        <select name="pChecks">
        <option value='-1'>Yes</option>
        <option value='0' selected>No</option>
        </select>
      </td>      
    </tr> 
      
    <tr> 
      <td width="189">Checks information (where to send, pay to, etc)</td>
      <td width="219">        
      <textarea name="pChecksComments" cols="35"></textarea>
    </td>      
    </tr> 
    <tr> 
      <td width="189">&nbsp;</td>
      <td width="219">&nbsp;</td>
    </tr>  
      <tr> 
      <td width="189">Accept Wire Transfer</td>
      <td width="219">        
        <select name="pWireTransfer">
        <option value='-1'>Yes</option>
        <option value='0' selected>No</option>
        </select>
      </td>
    </tr>
    
    <tr> 
      <td width="189">Wire information (Bank, ABA, Acc Number, Beneficiary)</td>
      <td width="219">        
      <textarea name="pWireComments" cols="35"></textarea>
    </td>      
    </tr>
    <tr> 
      <td width="189">&nbsp;</td>
      <td width="219">&nbsp;</td>
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
<br>(1) 2Checkout Payment scripts are included in the Power Pack, however if you sign up using our affiliate link you can get integration scripts for free. <a href="http://www.comersus.com/contact.html" target="_blank">Contact us to request your 2Checkout scripts here...</a>
<%call closeDb()%>
<!--#include file="footer.asp"-->