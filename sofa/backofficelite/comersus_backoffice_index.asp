<%
' Comersus BackOffice Lite
' Free Management Utility
' Comersus Open Technologies LC
' April-2003
' http://www.comersus.com 
%>

<!--#include file="includes/settings.asp"-->
<!--#include file="../includes/settings.asp"-->    
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<% 
on error resume next 

dim mySQL, conntemp, rstemp

pRunInstallationWizard 	= getSettingKey("pRunInstallationWizard")

if pRunInstallationWizard="-1" then response.redirect "comersus_backoffice_install0.asp"
%>
<!--#include file="header.asp"-->  
<font size="3"><b>Please login</b></font>
<br><br>
  <form method="post" name="login" action="comersus_backoffice_login.asp">    
  User<br>      
    
    <%if pBackOfficeDemoMode=-1 then%>
     <input type="text" name="adminName" size="20" value="admin">        
    <%else%>      
     <input type="text" name="adminName" size="20" value="">        
    <%end if%> 
    
    <br>Password<br>    
    
    <%if pBackOfficeDemoMode=-1 then%>
     <input type="password" name="adminpassword" size="20" value="123456">             
    <%else%>      
     <input type="password" name="adminpassword" size="20" value="">        
    <%end if%> 
  <br><br>
  <input type="submit" name="Submit2" value="Login"> <i>(default user is admin)</i>
  </form>
<br><br>
<!--#include file="footer.asp"-->