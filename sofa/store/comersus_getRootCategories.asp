<% 
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: get root categories
%>

<%

pShowCategories	= getSettingKey("pShowCategories")

if pShowCategories="-1" then

%>
             	
<%
 pIdCategoryStart=getCategoryStart(pIdStore)
 
 dim pCountCategories
 pCountCategories=0
 
 ' get categories
 mySQL="SELECT idCategory, categoryDesc FROM categories WHERE idParentCategory=" &pIdCategoryStart& " AND active=-1 AND idCategory<>1 ORDER BY displayOrder"

 call getFromDatabase (mySql, rsTempRoot, "getRootCategories")

 do while not rsTempRoot.eof and pCountCategories<8%>  
  
  <div class="b01">&nbsp;<img src="images/arrow.gif"> &nbsp;<a href="comersus_listOneCategory.asp?idCategory=<%=rsTempRoot("idCategory")%>"><%=rsTempRoot("categoryDesc")%></a></div>  
  
  <%
  pCountCategories=pCountCategories+1
  rsTempRoot.movenext
loop

%>
<div class="b01">&nbsp;<img src="images/arrow.gif"> &nbsp;<a href="comersus_listCategories.asp"><%=getMsg(94,"All")%>&nbsp;<%=getMsg(12,"Ctg")%></a></div>  	
<%
end if%>

