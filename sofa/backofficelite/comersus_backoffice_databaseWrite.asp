<%
' Comersus Diagnostics
' e-commerce ASP Open Source
' CopyRight Rodrigo S. Alhadeff, Comersus
' September 2004
' http://www.comersus.com 
' Diagnostics for Comersus store
%>
<!--#include file="comersus_backoffice_adminVerify.asp"-->    
<!--#include file="../includes/databaseFunctions.asp"-->  
<!--#include file="../includes/settings.asp"--> 
 
<!--#include file="header.asp"--> 
<b>Database Write</b>
<br>
<%
on error resume next

dim mySQL, conntemp, rstemp, pRdmProduct

pRdmProduct = "Product "&randomNumber(999)

%><br>Trying to insert one new product named: <%=pRdmProduct%><%
 	
mySQL="INSERT INTO products (description, details, price, active, idSupplier) VALUES ('" &pRdmProduct&"', 'Test Product', 100,0,Null)"
	
call updateDatabase(mySql, rstemp, "databaseWrite")

mySQL="SELECT idProduct FROM products WHERE description='" &pRdmProduct& "'"
	
set rsTemp=connTemp.execute(mySQL)

if instr(err.description,"Operation must use an updateable query")<>0  or instr(err.description,"idDbSession was not defined while")>0 then

 %><br><br>You are having permissions troubles over the Database file or folder.<br>Please check Comersus User's Guide to find how to fix the trouble.<%

else	
 if rstemp.eof then
 %>
	 <br>You are having database errors. More information: <%=err.description%>
<%else%>
         <br>Product inserted with ID #: <%=rstemp("idProduct")%>
 <%
 end if
end if

function randomNumber(limit)
 randomize
 randomNumber=int(rnd*limit)+2
end function

call closeDb()	
%>
<br><br>
<!--#include file="footer.asp"--> 