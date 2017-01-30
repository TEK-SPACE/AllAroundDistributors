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
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/stringFunctions.asp"--> 
<!--#include file="../includes/currencyFormat.asp"-->

<% 

on error resume next

dim mySQL, conntemp, rstemp

' get settings 
pDefaultLanguage 	= getSettingKey("pDefaultLanguage")
pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pStoreLocation	 	= getSettingKey("pStoreLocation")

pKey 			= getUserInput(request("key"),20)

pKey 			= replace(pKey,"'","")

if (trim(pKey)="") then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Please enter search criteria")
end if
  
mySQL = "SELECT * FROM products WHERE details LIKE '%" &pKey& "%' OR description LIKE '%" &pKey& "%'"

call getFromDatabase(mySQL, rstemp, "comersus_backoffice_searchItemExec.asp")  	

if  rstemp.eof then 
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("There are no items under your search")
end if

%>

<!--#include file="header.asp"-->
<!--#include file="./includes/settings.asp"-->

<br><b><span class="TextBlackSM">Product selection</span></b><br>
<br><i>Search key: </i><font color="#CC0000"><b><%=pKey%></b></font><br><br>

<table>
<tr bgcolor="#CCCCCC">
 <td>SKU</td>
 <td>Description</td>
 <td>Price</td>
 <td>Cost</td>
 <td>Active</td>
 <td>Clearance</td>
 <td>In Home</td>
 <td>Actions</td>
</tr>
<%

' set Count equal to zero
 Count = 0
 do while not rstemp.eof 

   pIdProduct		= rstemp("idProduct")
   pDescription		= rstemp("description")
   pPrice		= rstemp("price")
   pCost		= rstemp("cost")
   pDetails		= rstemp("details")
   pListPrice		= rstemp("listPrice")
   pSku			= rstemp("sku")   
   pActive		= rstemp("active")   
   pHotDeal		= rstemp("hotDeal")   
   pShowInHome		= rstemp("showInHome")   
   %>
<tr>
 <td><%=pSku%></td>
 <td><%=pDescription%></td>
 <td><%=pCurrencySign & money(pPrice)%></td>
 <td><%=pCurrencySign & money(pCost)%></td>
 <td><%
 if pActive="-1" then 
  response.write "Yes"
 else
  response.write "No"
 end if
 %></td>
  <td><%
 if pHotDeal="-1" then 
  response.write "Yes"
 else
  response.write "No"
 end if
 %></td>
   <td><%
 if pShowInHome="-1" then 
  response.write "Yes"
 else
  response.write "No"
 end if
 %></td>
 <td>
 <a href="comersus_backoffice_modifyProductForm.asp?idProduct=<%=pIdProduct%>">Edit</a>
 <a href="comersus_backoffice_deleteItemExec.asp?idProduct=<%=pIdProduct%>">Delete</a>
 </td>
</tr>      

 <%  
   count = count + 1
   rstemp.MoveNext
 loop
%> 
</table>
<br><i>Full BackOffice included with Power Pack includes pagination (Page x of y...)</i>
<!--#include file="footer.asp"--> 