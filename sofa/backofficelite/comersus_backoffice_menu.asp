<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"--> 
<!--#include file="includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->   
<!--#include file="../includes/databaseFunctions.asp"-->   
<!--#include file="../includes/settings.asp"--> 
<!--#include file="header.asp"-->

<%
dim connTemp, rsTemp

' get settings 
pCompany	 	= getSettingKey("pCompany")
%>

<br>
    <table width="590" border="0">
      <tr>
        <td width="320" height="72"> 
          <div class=title>Management area</div><br>
            <br>
            Today is <%=Date()%><br>
            Store <%=pCompany%>   
            <br>Storefront version: <%=pStoreFrontVersion%> <a href="http://www.comersus.com" target="_blank">Check for updates</a>        
          </td>
        <td width="270" height="72">&nbsp;</td>
      </tr>
      <tr> 
        <td width="320">
          <b>Main menu</b><br><br>
            • Configuration <a href="comersus_backoffice_settingsModifyForm.asp">Settings</a> - <a href="comersus_backoffice_menuShipping.asp">Shipping</a> - <a href="comersus_backoffice_listPayments.asp">Payment</a> - <a href="comersus_backoffice_addTaxPerPlaceForm.asp">Taxes</a><br>
            • <a href="comersus_backoffice_menuProducts.asp">Products</a><br>
            • <a href="comersus_backoffice_menuSales.asp">Sales</a><br>
            • <a href="comersus_backoffice_menuUtilities.asp">Utilities</a><br>            
            • <a href="comersus_backoffice_logoff.asp">Logoff</a>
        </td>
        <td width="270"> 
        <%randomize
        pNumber=int(rnd*7)+1        
        %>
        
        <%if pNumber=1 then%>
          <div align="center"><a href="http://www.comersus.com/store/pppremium.html" target="_blank"> 
            <img src="images/softPowerPack.gif" border="0" alt="Power Pack"></a><br>
            Upgrade to Comersus Power Pack Premium now <br>
            <i>(Includes BackOffice+, real time shipping and more!)</i></div>
        <%end if%>
        <%if pNumber=2 then%>
           <b>Accept online credit cards with 2Checkout</b>           
           <br><br><i>No monthly fee - $0.45 per transaction - 5.5% of transaction amount - Start Selling in 3 minutes</i>
           <a href="https://www.2checkout.com/2co/signup?affiliate=55003"><img src="images/2checkout.gif" alt="Sign Up Now" border="0"></a>
           <br><br>Get 2checkout payment form for free using <a href="https://www.2checkout.com/2co/signup?affiliate=55003">this link to sign-up</a>
         <%end if%>
         <%if pNumber=3 then%> 
         <center><A HREF="http://www.comersus.com/hosting.html" target="_blank"><IMG SRC="images/siteHosting100.gif" BORDER="0" ALT="Get hosting for your cart now"></A><br><br>Get hosting for your cart through <b>Comersus</b>
         <br>Fast mySQL database - ASP components - Support provided by programmers
         </center>
         <%end if%>     
         
         <%if pNumber=4 then%> 
         <center><a href="http://www.ccnow.com/cgi-local/signup.cgi?affiliate=comersus" target="_blank"><IMG SRC="images/CCNOW.gif" BORDER="0" ALT="Accept Credit Cards"></A><br><br><b>Accept credit card payments with CCNow</b>
         <br>Low commission rate of 4.9%+$0.50 per transaction - No need to set up a merchant account 
         - Free Comersus Script
         </center>
         <%end if%> 
         
         <%if pNumber=5 then%> 
         <center><a href="http://www.comersus.com/news-ecommerce-book.html" target="_blank"><IMG SRC="images/HowToCover.jpg" BORDER="0" ALT="Succeed in e-commerce"></A><br><br><b>Book "Succeeding in e-commerce"</b>
         <br>Written by the founder of Comersus - Tips and Secrets to succeed - Only $15.00 
         </center>
         <%end if%> 
         
         <%if pNumber=6 then%> 
         <center><a href="https://secure.nochex.com/apply/merchant_info.aspx?partner_id=172109090" target="_blank"><IMG SRC="images/nochex-logo.gif" BORDER="0" ALT="Accept credit cards now"></A><br><br><b>Accept credit cards with Nochex </b>
         <br>Can start at just 2.9% plus 20p per transaction 
         </center>         
         <%end if%> 
         
	 <%if pNumber=7 then%> 
         <center><a href="http://www.comersus.com/templates.html" target="_blank"><IMG SRC="images/icon-templates.gif" BORDER="0" ALT="Get templates for your store"></A><br><br><b>Get free templates for your store</b>
         <br>Download now <a href="http://www.comersus.com/templates.html">complete zip files</a> to replace default layout.          
         </center>         
         <%end if%> 
                     
        </td>
      </tr>
    </table>
<%call closeDb()%>
<!--#include file="footer.asp"-->