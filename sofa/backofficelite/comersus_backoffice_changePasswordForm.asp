<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"-->   
<!--#include file="../includes/settings.asp"-->

<%
on error resume next
dim mySQL, conntemp, rstemp

%>

<!--#include file="header.asp"-->
<!--#include file="./includes/settings.asp"-->

<br><b>Change admin password</b><br>
<br><i>6 to 10 chars, only letters and numbers are valid</i>
<form method="post" name="chPass" action="comersus_backoffice_changePasswordExec.asp" >
  <table width="418" border="0">    
    <tr> 
      <td width="189">New password</td>
      <td width="219">     
        <input type=password name=password size=10 maxlength=10>           
      </td>
    </tr>
     <tr> 
      <td width="189">Confirm new password</td>
      <td width="219">        
      <input type=password name=passwordVerify size=10 maxlength=10>   
      </td>
    </tr>
    <tr> 
      <td width="189">&nbsp;</td>
      <td width="219">&nbsp;</td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <input type="submit" name="Submit" value="Change">
      </td>
    </tr>
  </table>
</form>
<!--#include file="footer.asp"-->
