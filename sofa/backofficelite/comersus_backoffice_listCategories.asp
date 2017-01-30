<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/settings.asp"-->    
<!--#include file="header.asp"-->
<!--#include file="includes/settings.asp"-->

<br><b>Modify Categories</b><br>

<% 
on error resume next 

dim randomnum, mySQL, conntemp, rstemp

' get categories

mySQL="SELECT * FROM categories WHERE idCategory>1"

call getFromDatabase(mySQL, rstemp, "comersus_backoffice_listcategories.asp") 

do while not rstemp.eof
%> 
 <br><%=rstemp("categoryDesc")%> <a href="comersus_backoffice_modifyCategoryForm.asp?idCategory=<%=rstemp("idCategory")%>">Edit</a> 
 <%
 rstemp.movenext
loop
call closeDb()
%> 
<!--#include file="footer.asp"-->