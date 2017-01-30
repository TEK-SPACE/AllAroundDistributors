<%
' Comersus Diagnostics
' e-commerce ASP Open Source
' CopyRight Rodrigo S. Alhadeff, Comersus
' September 2002
' http://www.comersus.com 
%>
<!--#include file="comersus_backoffice_adminVerify.asp"-->    
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/settings.asp"-->    
<!--#include file="../includes/sessionFunctions.asp"--> 

<%
on error resume next

dim mySQL, conntemp, rstemp, categoryDescription, pIdCategory

if pBackOfficeDemoMode=-1 then
 response.redirect "comersus_backoffice_message.asp?message="&Server.Urlencode("Function disabled in demo mode.")
end if

%>
<!--#include file="header.asp"--> 
<b>Import utility</b>
<br><br>
<i>Importing file: </i><%=Server.MapPath("categories.txt")%><br>
<%	     
   Const ForReading = 1, ForWriting = 2
   Dim fso, MyFile, pLine, valuesToInsert
   Set fso 	= CreateObject("Scripting.FileSystemObject")
   Set MyFile 	= fso.OpenTextFile(Server.MapPath("categories.txt"), ForReading)
   pCounter=0   
 	Do While MyFile.AtEndOfStream <> True 	
      		pLine 		= MyFile.ReadLine                		      		      		      		      			 	
	 	if len(pLine)>1 then 
		mySQL="INSERT INTO categories (categoryDesc, idParentCategory, active) VALUES ('" &trim(pline)&"',1,-1)"
	 	call updateDatabase(mySQL, rstemp, "importUtilityExec.asp") 				
		
		response.write "<br>Inserting category " & trim(pLine)
		pCounter=pCounter+1
		end if
   	Loop

call closeDb()

%>
<br><br><%=pCounter%> records were imported.<br><br>
<!--#include file="footer.asp"--> 