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
<!--#include file="../includes/itemFunctions.asp"-->    
<!--#include file="../includes/settings.asp"-->    

<% 
on error resume next 

dim mySQL, conntemp, rstemp, rsTemp1

pIdCategory	= getUserInput(request.querystring("idCategory"),12)

if trim(pIdCategory)="" then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Enter category ID")
end if

' get data of the category to modify
mySQL="SELECT * FROM categories WHERE idCategory=" &pIdCategory
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_modifycategoryform.asp") 

pCategoryDesc		= rstemp("categoryDesc")
pIdParentCategory	= rstemp("idParentCategory")

' get all parent categories (not assigned to one product)
mySQL="SELECT idCategory, categoryDesc FROM categories WHERE idCategory<>" &pIdCategory
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_modifycategoryform.asp") 

if rstemp.eof then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Cannot get parent categories")
end If
%>

<!--#include file="header.asp"-->
<!--#include file="./includes/settings.asp"-->

<br><b>Edit category</b><br>
<form method="post" name="modifyCateg" action="comersus_backoffice_modifyCategoryExec.asp">
<input type="hidden" name="idCategory" size="60" value='<%=pIdCategory%>'>
  <table width="359" border="0">
    <tr> 
      <td width="194">Description</td>
      <td width="310">        
        <input type=text name=categoryDesc size=40 value="<%=pCategoryDesc%>">
        </td>
    </tr>
    <tr> 
      <td width="189">Parent category</td>
      <td width="219">        
      <select name="idParentCategory">
         <%do  while not rstemp.eof
          if itHasNoItemsInside(rstemp("idCategory")) then%> 
	          <option value='<%=rstemp("idCategory")%>'
	          <%if pIdParentCategory=rstemp("idCategory") then
	            response.write "selected"
	          end if%>
	          ><%=rstemp("categoryDesc")%></option>
          <%end if
          rstemp.movenext
          loop%>
      </select>
      </td>
    </tr>    
    <tr><td><input type="submit" name="modify" value="Modify"></td>
    <td><input type="submit" name="delete" value="Delete"></td></tr>
    <tr><td colspan=2> </td></tr>
  </table>
<!--#include file="footer.asp"-->
</form>