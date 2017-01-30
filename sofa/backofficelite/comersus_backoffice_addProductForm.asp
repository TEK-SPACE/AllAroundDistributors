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
<!--#include file="../includes/getSettingKey.asp"-->  
<!--#include file="../includes/settings.asp"-->    

<%
on error resume next

dim mySQL, conntemp, rstemp, rsTemp1

' get settings 
pCurrencySign	 	= getSettingKey("pCurrencySign")

%>

<!--#include file="./includes/settings.asp"-->
<!--#include file="header.asp"-->

<br><b>Add product</b><br><br>
<i>Place product image files inside /comersus/store folder </i>
<form method="post" name="addproduct" action="comersus_backoffice_addProductExec.asp"> 
        <table width="650" border="0">
          <tr> 
            <td width="84">SKU (*)</td>
            <td width="187"> 
              <input type="text" name="sku" value="" size=8>
            </td>
            <td colspan="2">Description (*)</td>
            <td width="253"> 
              <input type="text" name="description" size="40" value="">
            </td>
          </tr>
          <tr> 
            <td width="84">Details (*)</td>
            <td colspan=4> 
              <textarea name="details" rows="5" cols="60"></textarea>
            </td>
          </tr>
          <tr> 
            <td width="84">Price (*)<%=pCurrencySign%>:</td>
            <td width="187"> 
              <input type="text" name="price" value="0.00">
            </td>
            <td colspan="2"  >List price <%=pCurrencySign%>:</td>
            <td width="253"  > 
              <input type="text" name="listPrice" value="0.00">
            </td>
          </tr>
          <tr> 
            <td width="84">Wholesale price <%=pCurrencySign%>:</td>
            <td width="187"> 
              <input type="text" name="bToBprice" value="0.00">
            </td>
            <td colspan="2">Cost <%=pCurrencySign%>:</td>
            <td width="253"> 
              <input type="text" name="cost" value="0.00">
            </td>
          </tr>
          <tr> 
            <td width="84">Image file name</td>
            <td width="187"> 
              <input type="text" name="imageUrl" value="">
            </td>
            <td colspan="2">Thumbnail file name</td>
            <td width="253"> 
              <input type="text" name="smallImageUrl" value="">
            </td>
          </tr>
          <tr> 
            <td width="84">Weight</td>
            <td width="187"> 
              <input type="text" name="weight" value="0">
            </td>
            <td colspan="2">Stock</td>
            <td width="253"> 
              <input type="text" name="stock" value="0">
            </td>
          </tr>
          <tr> 
            <td width="84">Availability</td>
            <td width="187"> 
              <input type="text" name="deliveringTime" value="0">
            </td>
            <td colspan="2">Info to distribute (download link, serial, etc)</td>
            <td width="253"> 
              <input type="text" name="emailtext" value="">
            </td>
          </tr>
          <tr> 
            <td width="84">Supplier</td>
            <td width="187"> <%

mySQL="SELECT * FROM suppliers"
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_addproductform.asp") 

if  rstemp.eof then 
 response.redirect "comersus_backoffice_supportError.asp?error="& Server.Urlencode("No suppliers defined") 
end if %> 
              <select name="idSupplier">
                <%
do  until rstemp.eof 
   pIdSupplier		= rstemp("idSupplier")
   pSupplierName	= rstemp("supplierName")%> 
                <option value='<%=pIdSupplier%>'><%=pSupplierName%></option>
                <%
   rstemp.movenext
loop
%> 
              </select>
            </td>
            <%
' get leaf categories
dim arrCategories(500,2)

mySQL="SELECT * FROM categories WHERE idCategory>1"
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_addproductfrom.asp") 

if  rstemp.eof then 
  response.redirect "comersus_backoffice_supportError.asp?error="& Server.Urlencode("No categories defined") 
end if

arrCategoriesIndex=0
do  until rstemp.eof 
  if isCategoryLeaf(rstemp("idCategory")) then
  	arrCategories(arrCategoriesIndex,0) = rstemp("idCategory")
  	arrCategories(arrCategoriesIndex,1) = rstemp("categoryDesc")
  	arrCategoriesIndex=arrCategoriesIndex+1
  end if	
  rstemp.movenext
loop   
%> 
            <td colspan="2">Category 1</td>
            <td width="253"> 
              <select name="idCategory1">
                <%for f=0 to arrCategoriesIndex-1%> 
                <option value='<%=arrCategories(f,0)%>'><%=arrCategories(f,1)%></option>
                <%next%> 
              </select>
            </td>
          </tr>
          <tr> 
            <td width="84">Category 2</td>
            <td width="187"> 
              <select name="idCategory2">
                <option value='' selected>None</option>
                <%for f=0 to arrCategoriesIndex-1%> 
                <option value='<%=arrCategories(f,0)%>'><%=arrCategories(f,1)%></option>
                <%next%> 
              </select>
            </td>
            <td colspan="2">Category 3</td>
            <td width="253"> 
              <select name="idCategory3">
                <option value='' selected>None</option>
                <%for f=0 to arrCategoriesIndex-1%> 
                <option value='<%=arrCategories(f,0)%>'><%=arrCategories(f,1)%></option>
                <%next%> 
              </select>
            </td>
          </tr>
          <tr> 
            <td width="84">Drop down quantity</i></td>
            <td width="187"> 
              <input type=text name=formQuantity value=10 size=4>
            </td>
            <td colspan="2">Active</td>
            <td width="253"> 
              <input type="checkbox" name="active" value="-1" checked>
            </td>
          </tr>
          <tr> 
            <td width="84">Clearance</td>
            <td width="187"> 
              <input type="checkbox" name="hotDeal" value="-1" >
            </td>
            <td colspan="2">Hidden for search</td>
            <td width="253"> 
              <input type="checkbox" name="listhidden" value="-1">
            </td>
          </tr>
          <tr> 
            <td width="84">Show in home</td>
            <td colspan=4> 
              <input type="checkbox" name="showInHome" value="-1" checked>
            </td>
          </tr>
          <tr> 
            <td width="84">Variation 1</td>
            <td colspan="4">Drop down name 
              <input type="text" name="optionGroupDescrip1" value="" size=20>
            </td>
          </tr>
          <tr> 
            <td width="84"></td>
            <td width="187"><b>Description</b></td>
            <td colspan="2"><b>Price</b></td>
            <td width="253"><b>Percentage</b></td>
          </tr>
          <tr> 
            <td width="84"></td>
            <td width="187"> 
              <input type="text" name="optionDescrip1" value="" size=20>
            </td>
            <td colspan="2"> <%=pCurrencySign%> 
              <input type="text" name="priceToAdd1" value="0.00" size=6>
            </td>
            <td width="253"> % 
              <input type="text" name="percentageToAdd1" value="0.00" size=6>
            </td>
          </tr>
          <tr> 
            <td width="84"></td>
            <td width="187"> 
              <input type="text" name="optionDescrip1" value="" size=20>
            </td>
            <td colspan="2"> <%=pCurrencySign%> 
              <input type="text" name="priceToAdd1" value="0.00" size=6>
            </td>
            <td width="253">% 
              <input type="text" name="percentageToAdd1" value="0.00" size=6>
            </td>
          </tr>
          <tr> 
            <td width="84"></td>
            <td width="187"> 
              <input type="text" name="optionDescrip1" value="" size=20>
            </td>
            <td colspan="2"> <%=pCurrencySign%> 
              <input type="text" name="priceToAdd1" value="0.00" size=6>
            </td>
            <td width="253">% 
              <input type="text" name="percentageToAdd1" value="0.00" size=6>
            </td>
          </tr>
          <tr> 
            <td colspan=5>&nbsp;</td>
          </tr>
          <tr> 
            <td width="84">Variation 2</td>
            <td colspan="4">Drop down name 
              <input type="text" name="optionGroupDescrip2" value="" size=20>
            </td>
          </tr>
          <tr> 
            <td width="84"></td>
            <td width="187"><b>Description</b></td>
            <td colspan="2"><b>Price</b></td>
            <td width="253"><b>Percentage</b></td>
          </tr>
          <tr> 
            <td width="84"></td>
            <td colspan=2>
<input type="text" name="optionDescrip2" value="" size=20>
            </td>
            <td width="104"> <%=pCurrencySign%> 
              <input type="text" name="priceToAdd2" value="0.00" size=6>
            </td>
            <td width="253">% 
              <input type="text" name="percentageToAdd2" value="0.00" size=6>
            </td>
          </tr>
          <tr> 
            <td width="84"></td>
            <td colspan=2> 
              <input type="text" name="optionDescrip2" value="" size=20>
            </td>
            <td width="104"> <%=pCurrencySign%> 
              <input type="text" name="priceToAdd2" value="0.00" size=6>
            </td>
            <td width="253">% 
              <input type="text" name="percentageToAdd2" value="0.00" size=6>
            </td>
          </tr>
          <tr> 
            <td width="84"></td>
            <td colspan=2> 
              <input type="text" name="optionDescrip2" value="" size=20>
            </td>
            <td width="104"> <%=pCurrencySign%> 
              <input type="text" name="priceToAdd2" value="0.00" size=6>
            </td>
            <td width="253">% 
              <input type="text" name="percentageToAdd2" value="0.00" size=6>
            </td>
          </tr>
          <tr> 
            <td colspan=5>&nbsp;</td>
          </tr>
          <tr> 
            <td colspan=5> 
              <input type="submit" name="Submit" value="Add product">
            </td>
          </tr>
        </table>
</form>
<!--#include file="footer.asp"-->

