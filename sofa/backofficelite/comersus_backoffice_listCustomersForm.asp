<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"--> 
<!--#include file="header.asp"-->
<!--#include file="./includes/settings.asp"-->
<br><b>Customers</b><br>
<form method="post" name="listCust" action="comersus_backoffice_listCustomersExec.asp">
<table width="421" border="0">
    <tr> 
      <td width="136">Enter name</td>
      <td width="275">        
        <input type=text name=customerName size=30> 
       </td>
    </tr>
    <tr>
 <td colspan="3"> 
 <input type="submit" name="search" value="Search">
</td>
</tr>
</table>
<!--#include file="footer.asp"-->
</form>