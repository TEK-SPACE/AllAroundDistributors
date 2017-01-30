<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"-->   
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/itemFunctions.asp"-->    
<!--#include file="../includes/settings.asp"-->    
<!--#include file="../includes/stringFunctions.asp"--> 

<%
on error resume next

dim mySQL, conntemp, rstemp

' get all categories
mySQL="SELECT idCategory, categoryDesc, idParentCategory, displayOrder FROM categories ORDER BY categoryDesc"

call getFromDatabase (mySql, rsTemp, "comersus_backoffice_addCategoryForm.asp")

if  rstemp.eof then 
  response.redirect "comersus_backoffice_supportError.asp?error="& Server.Urlencode("Error in addcategoryform: you need at least one root category") 
end if

arrCategoriesIndex=0

' get categories with no items inside (there you can insert other categories)
dim arrCategories(1000,2)
do while not rstemp.eof    
   
   if itHasNoItemsInside(rstemp("idCategory")) then
    arrCategories(arrCategoriesIndex,0) = rstemp("idCategory")
    arrCategories(arrCategoriesIndex,1) = rstemp("categoryDesc")
    arrCategoriesIndex		        = arrCategoriesIndex+1   
   end if
      
   rstemp.moveNext
loop

%>

<!--#include file="header.asp"-->
<!--#include file="./includes/settings.asp"-->

<br><b><span class="TextBlackSM">Add category</span></b><br>
<br><%=getScreenMessage(request.querystring ("message"),200)%>
<form method="post" name="addCateg" action="comersus_backoffice_addCategoryExec.asp">
  <table width="418" border="0">
    <tr> 
      <td width="189">Description</td>
      <td width="219">        
        <input type=text name="categoryDesc">
      </td>
    </tr>
     <tr> 
      <td width="189">Parent category</td>
      <td width="219">        
        <select name="idParentCategory">        
          <%for f=0 to arrCategoriesIndex-1%> 
           <option value='<%=arrCategories(f,0)%>' <%if arrCategories(f,0)=1 then response.write "selected"%>><%=arrCategories(f,1)%></option>
          <%next%>  
        </select>
      </td>
    </tr>
    <tr> 
      <td width="189">&nbsp;</td>
      <td width="219">&nbsp;</td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <input type="submit" name="Submit" value="Add category">
      </td>
    </tr>
  </table>
</form>
<%call closeDb()%>
<!--#include file="footer.asp"-->
