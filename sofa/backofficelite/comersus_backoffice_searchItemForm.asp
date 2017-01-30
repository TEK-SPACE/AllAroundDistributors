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

<br><font size="3"><b><span class="TextBlackSM">Search product</span></font></b>
<br><br>Please enter search criteria in order to modify or delete a product:
<br>

<table width="584" border="0" cellspacing="0" cellpadding="0" height="18">
  <tr> 
    <td width="368" height="47"> 
      <form method="post" action="comersus_backoffice_searchItemExec.asp?idsubcategory=0" name="search">
        <div align="left">  
          <input type="text" name="key" size="20">
          <input type="submit" name="Submit2" value="Search">
          </div>
      </form>
    </td>
  </tr>
</table>
<!--#include file="footer.asp"-->

