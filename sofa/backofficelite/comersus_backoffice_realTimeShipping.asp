<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"--> 
<!--#include file="header.asp"--> 
    <p><!--#include file="./includes/settings.asp"--> <br>
      <b>Real Time Shipment</span></b><br>
      <br>
      Comersus is compatible with Fedex, UPS, USPS and DHL. You can obtain real 
      time quotes from all the carriers at once. <br>
      <br>
      <img src="images/shippingLogos.gif"> </p>
    <p>Example: 5lbs From Zip: 33166 To Zip: 90210</p>
    <form name="form1" method="post" action="">
      <select name="shipping" size="4">
        <option value="$18.43" selected>Fedex 2nd Day $18.43</option>
        <option value="$16.08" selected>UPS 3 Day $16.08</option>
        <option value="$30.99" selected>DHL Overnight $30.99</option>
        <option value="$27.30" selected>USP Express Mail Addresses $27.30</option>
      </select>
    </form>
            
    <p>More information about <a href="http://www.comersus.com/power-pack.html" target="_blank">Power Packs here...</a></p>
    <!--#include file="footer.asp"-->