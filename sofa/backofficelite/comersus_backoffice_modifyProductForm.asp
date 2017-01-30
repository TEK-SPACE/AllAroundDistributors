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
<!--#include file="../includes/getSettingKey.asp"-->     
<!--#include file="../includes/settings.asp"-->    
<!--#include file="../includes/currencyFormat.asp" --> 

<% 
on error resume next

dim mySQL, conntemp, rstemp, rsTemp1

' get settings 
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign   		= getSettingKey("pDecimalSign")
pMoneyDontRound		= getSettingKey("pMoneyDontRound")
pChangeDecimalPoint	= getSettingKey("pChangeDecimalPoint")

' form parameter 
pIdProduct		= getUserInput(request.querystring("idProduct"),12)

if trim(pIdProduct)="" then
   response.redirect "comersus_backoffice_message.asp?message="&Server.Urlencode("You must specify item #.")
end if

' get item details from db
mySQL="SELECT idProduct, idSupplier, description, price, listPrice, bToBPrice, cost, details, imageUrl, smallImageUrl, deliveringTime FROM products WHERE products.idProduct=" &pIdProduct
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_modifyProductForm.asp")  	

' charge rscordset data into local variables
pIdProduct		= rstemp("idProduct")
pIdSupplier		= rstemp("idSupplier")
pDescription		= rstemp("description")
pPrice			= rstemp("price")
pListPrice		= rstemp("listPrice")
pBToBPrice		= rstemp("bToBPrice")
pCost			= rstemp("cost")
pDetails		= rstemp("details")
pImageUrl		= rstemp("imageUrl")
pSmallImageUrl		= rstemp("smallImageUrl")
pDeliveringTime		= rstemp("deliveringTime")

' second query 
mySQL="SELECT listHidden, hotDeal, active, weight, listPrice, sku, formQuantity, emailText, showInHome FROM products WHERE products.idProduct=" &pIdProduct
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_modifyProductForm.asp")  	

' charge rscordset data into local variables
pListhidden		= rstemp("listhidden")
pHotDeal		= rstemp("hotDeal")
pActive			= rstemp("active")
pShowInHome		= rstemp("showInHome")
pWeight			= rstemp("weight")
pListPrice		= rstemp("listPrice")
pSku			= rstemp("sku")
pFormQuantity		= rstemp("formQuantity")
pEmailText		= rstemp("emailText")

' end second query

pStock			= getStock(pIdProduct)

' change , with .
if pChangeDecimalPoint="-1" then
 pPrice		= replace(pPrice, ",",".")
 pListPrice	= replace(pListPrice, ",",".")
 pBToBPrice	= replace(pBtoBPrice, ",",".")
 pCost		= replace(pCost, ",",".")
end if

' load variations
dim pHiddenIdOptions, pOptionDescrip1, pOptionDescrip2, pIdOptionGroup1, pIdOptionGroup2
dim arrayOptions1(3,3)
dim arrayOptions2(3,3)

call loadProductVariations(pIdProduct, arrayOptions1, arrayOptions2, pHiddenIdOptions, pOptionDescrip1, pOptionDescrip2, pIdOptionGroup1, pIdOptionGroup2)
%>

<!--#include file="header.asp"-->
<!--#include file="includes/settings.asp"-->

<br><b>Modify product</b><br>
<form method="post" name="modifyProduct" action="comersus_backoffice_modifyProductExec.asp"> 
  Item #: <%=pIdProduct%> 
  
  <input type="hidden" name="idProduct"       value="<%=pIdProduct%>">
  <input type="hidden" name="idOptionGroup1"  value="<%=pIdOptionGroup1%>">
  <input type="hidden" name="idOptionGroup2"  value="<%=pIdOptionGroup2%>">
  <input type="hidden" name="hiddenIdOptions" value="<%=pHiddenIdOptions%>">


  <table width="650" border="0">
    <tr> 
      <td width="105">SKU</td>
      <td width="250"><input type="text" name="sku" value="<%=pSku%>"></td>
      <td width="95">Description</td>
      <td width="200"><input type="text" name="description" size="40" value="<%=pDescription%>"></td>
    </tr>
    
    <tr>
      <td>Details</td>
      <td colspan=3><textarea name="details" rows="5" cols="60"><%=pDetails%></textarea></td>
    </tr>
    
    <tr> 
      <td width="120">Price <%=pCurrencySign%>:</td>
      <td width="180"><input type="text" name="price" value="<%=money(pPrice)%>"> </td>
      <td width="150">List price <%=pCurrencySign%>:</td>
      <td width="200"><input type="text" name="listPrice" value="<%=money(pListPrice)%>"> </td>
    </tr>    
    <tr> 
      <td width="120">Wholesale price <%=pCurrencySign%>:</td>
      <td width="180"><input type="text" name="bToBprice" value="<%=money(pBtoBPrice)%>"></td>
      <td width="120">Cost <%=pCurrencySign%>:</td>
      <td width="200"><input type="text" name="cost" value="<%=money(pCost)%>"></td>
    </tr>  
    <tr> 
      <td width="120">Image file name</td>
      <td width="180"><input type="text" name="imageUrl" value="<%=pImageUrl%>"></td>
      <td width="120">Thumbnail file name</td>
      <td width="200"><input type="text" name="smallImageUrl" value="<%=pSmallImageUrl%>"></td>
    </tr>
    <tr> 
      <td width="120">Weight</td>
      <td width="180"> <input type="text" name="weight" value="<%=pWeight%>"></td>
      <td width="120">Stock</td>
      <td width="180"><input type="text" name="stock" value="<%=pStock%>"></td>
    </tr>
    <tr> 
      <td width="120">Availability (days)</td>
      <td width="180"><input type="text" name="deliveringTime" value="<%=pDeliveringTime%>"></td>
      <td width="120">Text to distribute (links, serials, etc)</td>
      <td width="200"><input type="text" name="emailtext" value="<%=pEmailText%>"></td>      
    </tr>

    <tr> 
      <td width="120">Supplier</td>
      <td width="200"> 
<%

mySQL="SELECT idSupplier, supplierName FROM suppliers"

call getFromDatabase(mySQL, rstemp, "comersus_backoffice_modifyProductForm.asp")  	

if  rstemp.eof then 
    response.redirect "comersus_backoffice_supporterror.asp?error="& Server.Urlencode("No suppliers defined") 
end if %> 
        <select name="idSupplier">
          <%
do  until rstemp.eof 
   pIdSupplier2		= rstemp("idSupplier")
   pSupplierName	= rstemp("supplierName")%> 
          <option value='<%=pIdSupplier2%>'
          <%if pIdSupplier2=pIdSupplier then
              response.write "selected"
          end if%>
          ><%=pSupplierName%></option>
          <%
   rstemp.movenext
loop

' get categories assigned
mySQL="SELECT * FROM categories_products WHERE idProduct=" &pIdProduct

call getFromDatabase(mySQL, rstemp, "comersus_backoffice_modifyProductForm.asp")  	

categoryCounter=1
do while not rstemp.eof
   select case categoryCounter
   case 1
     pIdCategory1=rstemp("idCategory")
   case 2  
     pIdCategory2=rstemp("idCategory")
   case 3
     pIdCategory3=rstemp("idCategory")
   case 4
     pIdCategory4=rstemp("idCategory")
   end select 
 categoryCounter=categoryCounter+1
 rstemp.movenext
loop

' get leaf categories
dim arrCategories(500,2)

mySQL="SELECT * FROM categories WHERE idCategory>1" 
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_modifyProductForm.asp")  	

if  rstemp.eof then 
  response.redirect "comersus_backoffice_supportError.asp?error="& Server.Urlencode("No categories defined") 
end if

arrCategoriesIndex=0

do  until rstemp.eof 
	if isLeafOfTheTree(rstemp("idCategory")) then
  		arrCategories(arrCategoriesIndex,0) = rstemp("idCategory")
  		arrCategories(arrCategoriesIndex,1) = rstemp("categoryDesc")
  		arrCategoriesIndex		    = arrCategoriesIndex+1  		
  	end if	
  	rstemp.movenext
loop  

%> 
      <td width="120">Category 1</td>
      <td width="200"> 
<select name="idCategory1">  
<%for f=0 to arrCategoriesIndex-1       
   ' detect sucbcategory assigned
   if arrCategories(f,0)=pidCategory1 then%> 
      <option value='<%=arrCategories(f,0)%>' selected><%=arrCategories(f,1)%></option>
   <%else%> 
      <option value='<%=arrCategories(f,0)%>'><%=arrCategories(f,1)%></option>
   <%end if
  next%> 
  </select>
        </td>
    </tr>
    <tr> 
      <td width="120">Category 2</td>
      <td width="200"> 
 <select name="idCategory2">  
 <option value='' selected>None</option>
<%for f=0 to arrCategoriesIndex-1       
   ' detect sucbcategory assigned
   if arrCategories(f,0)=pidCategory2 then%> 
      <option value='<%=arrCategories(f,0)%>' selected><%=arrCategories(f,1)%></option>
   <%else%> 
      <option value='<%=arrCategories(f,0)%>'><%=arrCategories(f,1)%></option>
   <%end if
  next%> 
  </select>
        </td>
      <td width="120">Category 3</td>
      <td width="200"> 
<select name="idCategory3">  
<option value='' selected>None</option>
<%for f=0 to arrCategoriesIndex-1       
   ' detect sucbcategory assigned
   if arrCategories(f,0)=pIdCategory3 then%> 
      <option value='<%=arrCategories(f,0)%>' selected><%=arrCategories(f,1)%></option>
   <%else%> 
      <option value='<%=arrCategories(f,0)%>'><%=arrCategories(f,1)%></option>
   <%end if
  next%> 
  </select>
        </td>
    </tr>
    <tr>
      <td width="140">Form drop down quantity</i></td>
      <td width="60"><input type="text" name="formQuantity" size="5" value="<%=pFormQuantity%>"></td>
      <td>Active</td>
      <td width="40" align="left">
        <input type="checkbox" name="active" value=-1
        <%
        if pactive=-1 then
         response.write "checked"
        end if
        %> 
        >
      </td>
    </tr>
    <tr>
      <td>Clearance</td>
      <td>
        <input type="checkbox" name="hotDeal" value=-1
        <%
        if photDeal=-1 then
         response.write "checked"
        end if
        %>
        ></td>
      <td>Hidden for listing</td>
      <td>
        <input type="checkbox" name="listhidden" value=-1 <%
        if plisthidden=-1 then
         response.write "checked"
        end if
        %> >
      </td>
   </tr>
   <tr>
      <td>Show in home</td>
      <td colspan=3>
        <input type="checkbox" name="showInHome" value=-1 <%
        if pShowInHome=-1 then
         response.write "checked"
        end if
        %> >
      </td>
    </tr>
    
   <tr>
            <td width="84">Variation 1</td>
            <td colspan="4">Drop down name 
              <input type="text" name="optionGroupDescrip1" value="<%=pOptionDescrip1%>" size=20>
            </td>
          </tr>

             <tr> 
            <td width="187"><b>Description</b></td>
            <td colspan="2"><b>Price</b></td>
            <td width="253"><b>Percentage</b></td>
          </tr>
          <tr> 
            <td width="187"> 
              <input type="text" name="optionDescrip1" value="<%=arrayOptions1(0,0)%>" size=20>
            </td>
            <td colspan="2"> <%=pCurrencySign%> 
              <input type="text" name="priceToAdd1" value="<%=money(arrayOptions1(1,0))%>" size=6>
            </td>
            <td width="253"> % 
              <input type="text" name="percentageToAdd1" value="<%=money(arrayOptions1(2,0))%>" size=6>
            </td>
          </tr>
          <tr> 
            <td width="187"> 
              <input type="text" name="optionDescrip1" value="<%=arrayOptions1(0,1)%>" size=20>
            </td>
            <td colspan="2"> <%=pCurrencySign%> 
              <input type="text" name="priceToAdd1" value="<%=money(arrayOptions1(1,1))%>" size=6>
            </td>
            <td width="253">% 
              <input type="text" name="percentageToAdd1" value="<%=money(arrayOptions1(2,1))%>" size=6>
            </td>
          </tr>
          <tr> 
            <td width="187"> 
              <input type="text" name="optionDescrip1" value="<%=arrayOptions1(0,2)%>" size=20>
            </td>
            <td colspan="2"> <%=pCurrencySign%> 
              <input type="text" name="priceToAdd1" value="<%=money(arrayOptions1(1,2))%>" size=6>
            </td>
            <td width="253">% 
              <input type="text" name="percentageToAdd1" value="<%=money(arrayOptions1(2,2))%>" size=6>
            </td>
          </tr>

    
        <tr>
            <td width="84">Variation 1</td>
            <td colspan="4">Drop down name 
              <input type="text" name="optionGroupDescrip2" value="<%=pOptionDescrip2%>" size=20>
            </td>
         </tr>

    
             <tr> 
            <td width="187"><b>Description</b></td>
            <td colspan="2"><b>Price</b></td>
            <td width="253"><b>Percentage</b></td>
          </tr>
          <tr> 
            <td width="187"> 
              <input type="text" name="optionDescrip2" value="<%=arrayOptions2(0,0)%>" size=20>
            </td>
            <td colspan="2"> <%=pCurrencySign%> 
              <input type="text" name="priceToAdd2" value="<%=money(arrayOptions2(1,0))%>" size=6>
            </td>
            <td width="253"> % 
              <input type="text" name="percentageToAdd2" value="<%=money(arrayOptions2(2,0))%>" size=6>
            </td>
          </tr>
          <tr> 
            <td width="187"> 
              <input type="text" name="optionDescrip2" value="<%=arrayOptions2(0,1)%>" size=20>
            </td>
            <td colspan="2"> <%=pCurrencySign%> 
              <input type="text" name="priceToAdd2" value="<%=money(arrayOptions2(1,1))%>" size=6>
            </td>
            <td width="253">% 
              <input type="text" name="percentageToAdd2" value="<%=money(arrayOptions2(2,1))%>" size=6>
            </td>
          </tr>
          <tr> 
            <td width="187"> 
              <input type="text" name="optionDescrip2" value="<%=arrayOptions2(0,2)%>" size=20>
            </td>
            <td colspan="2"> <%=pCurrencySign%> 
              <input type="text" name="priceToAdd2" value="<%=money(arrayOptions2(1,2))%>" size=6>
            </td>
            <td width="253">% 
              <input type="text" name="percentageToAdd2" value="<%=money(arrayOptions2(2,2))%>" size=6>
            </td>
          </tr>

    
    <tr> 
      <td colspan=4> 
        <input type="submit" name="Submit" value="Modify">
      </td>
    </tr>
  </table>
  </form>
<%call closeDb()%>  
<!--#include file="footer.asp"-->
<%
Private Function isLeafOfTheTree (idCategory)
isLeafOfTheTree=0
mySQL="Select idCategory from categories where idParentCategory="&idCategory
call getFromDatabase(mySQL, rsTemp1, "comersus_backoffice_addproductform.asp") 
if rsTemp1.eof then
	isLeafOfTheTree=1
end if	
end function
%>