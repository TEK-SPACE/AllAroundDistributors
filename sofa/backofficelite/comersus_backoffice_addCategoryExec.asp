<%
' Comersus BackOffice Lite
' Free Management Utility
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

dim mySQL, mySQL1, conntemp, rstemp

' form parameter
pCategoryDesc 		= getUserInput(request.form("categoryDesc"),150)
pIdParentCategory 	= getUserInput(request.form("idParentCategory"),10)

if pCategoryDesc="" or pIdParentCategory="" then
 response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Please enter category description and parent.")	
end if

' insert category in to db
mySQL="INSERT INTO categories (categoryDesc, idParentCategory, active) VALUES ('" &pCategoryDesc& "'," &pIdParentCategory& ",-1)"
call updateDatabase(mySQL, rstemp, "comersus_backoffice_addcategoryexec.asp") 

response.redirect "comersus_backoffice_addCategoryForm.asp?message="& Server.Urlencode(pCategoryDesc& " added")	

call closeDb()
%>