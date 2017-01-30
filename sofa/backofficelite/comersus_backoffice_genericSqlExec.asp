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

<%
on error resume next 

dim mySQL, conntemp, rstemp, counter

mySql=lcase(request.form("sqlsentence"))

if mySql="" then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Please enter a valid sentence")
end if

g=instr(mySQL,"update")
h=instr(mySQL,"insert")
i=instr(mySQL,"delete")
j=instr(mySQL,"alter")
k=instr(mySQL,"drop")

counter=0

if pBackOfficeDemoMode=-1 then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Function disabled in demo mode")
end if

call updateDatabase(mySQL, rstemp, "comersus_backoffice_genericSQLExec.asp") 

if err.number <> 0 then
   response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Your sentence [" &mySQL& "] returned errors: ["&err.Description&"]") 
end if
%>

<!--#include file="./includes/settings.asp"-->

<br><b>SQL Query</b><br>
<br>Query: <%=mySql%><br><br>
<%

if (g+h+i+j) <> 0 then
' modify%>
 <br>Sentence executed
<%else

if rstemp.eof then%>
 <br>Results
<%else%>

<table border=1 WIDTH="100%">
  <%do while not rstemp.eof%>    
   <tr>		
   <%for each whatever in rstemp.fields
       if whatever.name<>"" then
         dbField=whatever.name         
	 %>
	 
	 <td>
	 <%
	 if counter=0 then
	  response.write "<b>" &dbField& "</b>"	  
	 else
	  response.write rstemp(dbField) & "&nbsp;"	  
	 end if	 
	 %>
	 </td>
	 
	 <%
      end if ' not empty
     next%>
     </tr>
  <%
  counter=counter+1
  if counter>1 then
     rstemp.movenext
  end if
 loop
 %>
 </table>
<%
  end if ' no results
 end if  ' modify
call closeDb()%> 