<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"--> 
<!--#include file="./includes/settings.asp"-->
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/stringFunctions.asp"-->    
<!--#include file="../includes/settings.asp"-->    

<% 
on error resume next 

dim mySQL, conntemp, rstemp

pIdcategory		= getUserInput(request.form("idCategory"),4)
pIdParentCategory	= getUserInput(request.form("idParentCategory"),4)
pCategoryDesc		= getUserInput(request.form("categoryDesc"),150)
pBoton			= getUserInput(request.form("modify"),20)


if trim(pBoton)="Modify" then

 ' verifies that there are no products under that idparentcategory 

 mySQL="SELECT * FROM categories_products WHERE idCategory="& pIdParentCategory
 call getFromDatabase(mySQL, rstemp, "comersus_backoffice_modifycategoryexec.asp") 

 if not rstemp.eof then
   response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Category cannot be assigned to a category with items inside")
 end If


 mySQL="UPDATE categories SET categoryDesc='" &pCategoryDesc& "', idParentcategory=" &pIdParentCategory& " WHERE idCategory=" &pIdCategory
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_modifycategoryexec.asp") 

 response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Category modified")

else

' delete option

 ' verify assignment products
 mySQL="SELECT products.idProduct FROM products, categories_products WHERE products.idProduct=categories_products.idProduct AND idCategory=" &pIdCategory
 call getFromDatabase(mySQL, rstemp, "comersus_backoffice_modifycategoryexec.asp") 

 if not rstemp.eof then
   response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("The category has products inside. Delete products first or assign them to another category")
 end if

 ' verify categories assigned
 mySQL="SELECT idCategory from categories WHERE idParentCategory=" &pIdCategory
 call getFromDatabase(mySQL, rstemp, "comersus_backoffice_modifycategoryexec.asp") 

 if not rstemp.eof then
   response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("The category has other categories inside. Delete those categories first.")
 end if

 ' delete from categories_products
 mySQL="DELETE FROM categories_products WHERE idCategory=" &pIdCategory
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_modifycategoryexec.asp") 

 ' delete from categories
 mySQL="DELETE FROM categories WHERE idCategory=" &pIdCategory
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_modifycategoryexec.asp") 

  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Category deleted")

end if ' delete
%>
