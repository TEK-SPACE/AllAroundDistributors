<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"-->   
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/stringFunctions.asp"-->    
<!--#include file="../includes/settings.asp"-->    

<%
on error resume next

dim rstemp, conntemp, mysql

pCustomerName = getUserInput(request("customerName"),50)

if pCustomerName="" then
 response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Please enter customer search criteria")
end if

mySql= "SELECT * FROM customers WHERE name LIKE '%" &pCustomerName& "%'"
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_listcustomersexec.asp") 

if rstemp.eof then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("No customers under that search")
end if
%>

<!--#include file="header.asp"-->
<!--#include file="./includes/settings.asp"-->

<br><b>Customers</b><br>

<table border=0 width="595">
  <tr bgcolor="#CCCCCC"> 
    <td>Name</td>
    <td>Address</td>
    <td>Phone</td>
    <td>City</td>
    <td>Country</td>
    <td>Email</td>   
    <td>Customer type</td>
  </tr>

<%do while not rstemp.eof %>
<tr>
<td><a href="comersus_backoffice_modifyCustomerForm.asp?idCustomer=<%=(rstemp("idCustomer"))%>"><%=(rstemp("name")&" "&rstemp("LastName"))%></a></td>
<td><%=(rstemp("address"))%></td>
<td><%=(rstemp("phone"))%></td>
<td><%=(rstemp("city"))%></td>
<td><%=rstemp("countryCode")%></td>
<td><a href="mailto:<%=rstemp("email")%>"><img src="images/mail.gif" width="14" height="10" border="0"></a> </td>
<td><%

mySql= "SELECT customerTypeDesc FROM customerTypes WHERE idCustomerType= " &rstemp("idCustomerType")
call getFromDatabase(mySQL, rstemp2, "comersus_backoffice_listcustomersexec.asp") 
if not rstemp2.eof then
 response.write (rstemp2("customerTypeDesc"))
end if
%></td>
</tr>
<%rstemp.movenext
loop

call closeDb()
%>
</table>
<!--#include file="footer.asp"-->  