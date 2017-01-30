<%
' Comersus BackOffice Lite
' Free Management Utility
' Comersus Open Technologies LC
' April-2003
' http://www.comersus.com 
%>

<!--#include file="includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->   
<!--#include file="../includes/databaseFunctions.asp"-->   
<!--#include file="../includes/settings.asp"--> 
<!--#include file="header.asp"-->
<br><br><b>You are logged off, now see what you are missing</b>
<br><br>This screen was obtained from BackOffice+, the full management suite included in Power Pack.
<br>More information about BackOffice+ and Power Packs at <a href="http://www.comersus.com/power-pack.html">http://www.comersus.com/power-pack.html</a>
<br>  
<%randomize
pNumber=int(rnd*4)+1%>

<%if pNumber=1 then%>          
 <br><img src="images/plusInvoice.gif" width="400" height="354"><br><br><i>Details: invoice obtained from an existing order. Includes Bar Code and a button to be printed.</i><br>
<%end if%>

<%if pNumber=2 then%>          
 <br><img src="images/plusFore.gif" width="400" height="392"><br><br><i>Details: This chart provides a forecasting of next month sales based on previous sales.</i><br><br>
<%end if%>

<%if pNumber=3 then%>          
 <br><img src="images/plusMain.gif" width="400" height="295"><br><br><i>Details: this is the main screen of the BackOffice+ including shortcuts, sales   charts, relaxing Haikus and links to common tasks.</i><br>
<%end if%>

<%if pNumber=4 then%>          
 <br><img src="images/plusSales.gif" width="400" height="395"><br><br><i>Details: this is a Sales and Visits Chart in 3D.</i><br>
<%end if%>

<br><br><b>BackOffice+ features</b><br>
<br>. Multi User with different permission levels
<br>. Each BackOffice User may have different views and permissions
<br>. Wysiwyg HTML Editor for products descriptions
<br>. Traffic Booster Google Friendly (generate HTML product details pages for Search Engines)
<br>. Newsletter distributions
<br>. Export to Microsoft Excel, QuickBooks, CSV, VCF, PriceGrabber and Froogle 
<br>. Reports and charts (stocks, sales, statistics, etc) 
<br>. Sales Forecasting using mathematical theory 
<br>. List orders with direct access to remote SSL credit card data (off-line payment methods)
<br>. Suppliers private login with listing of their sales 
<br>. Enter orders from the BackOffice (ideal for Phone orders)
<br>. Print Invoices and Shipping Labels with BarCodes 
<br>. Print Product Labels with BarCodes 
<br>. Update orders with Shipping Tracking and Transaction Results 
<br><br>More information about Power Pack and services at <a href="http://www.comersus.com/pricing.html">http://www.comersus.com/pricing.html</a>

<br><br><br><img src="images/softPowerPack.gif" alt="Get Power Pack Now">
<br>Get Power Pack <a href="http://www.comersus.com/power-pack.html">here...</a>
<!--#include file="footer.asp"-->
