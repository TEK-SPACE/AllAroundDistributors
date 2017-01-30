<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"--> 
<!--#include file="comersus_backoffice_functions.asp"-->   
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/itemFunctions.asp"-->    
<!--#include file="../includes/settings.asp"-->    
<!--#include file="../includes/stringFunctions.asp"-->    
<% 
if pBackOfficeDemoMode=-1 then
 response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Function disabled in demo mode")
end if

' on error resume next

dim mySQL, conntemp, rstemp

pIdProduct		= getUserInput(request.form("idProduct"),12)
pSku			= getUserInput(request.form("sku"),20)
pDescription		= getUserInput(request.form("description"),100)
pDetails		= getUserInput(request.form("details"),2000)
pIdCategory1		= getUserInput(request.form("idCategory1"),4)
pIdCategory2		= getUserInput(request.form("idCategory2"),4)
pIdCategory3		= getUserInput(request.form("idCategory3"),4)
pPrice			= getUserInput(request.form("price"),20)
pListPrice		= getUserInput(request.form("listPrice"),20)
pBtoBPrice		= getUserInput(request.form("btoBPrice"),20)
pCost			= getUserInput(request.form("cost"),20)
pImageUrl		= getUserInput(request.form("imageUrl"),150)
pSmallImageUrl		= getUserInput(request.form("smallImageUrl"),150)
pIdSupplier		= getUserInput(request.form("idSupplier"),4)
pListHidden		= getUserInput(request.form("listHidden"),2)
pHotDeal		= getUserInput(request.form("hotDeal"),2)
pWeight			= getUserInput(request.form("weight"),12)
pShowInHome		= getUserInput(request.form("showInHome"),2)
pStock			= getUserInput(request.form("stock"),12)
pDeliveringTime		= getUserInput(request.form("deliveringTime"),4)
pActive			= getUserInput(request.form("active"),2)
pEmailText		= getUserInput(request.form("emailText"),250)
pFormQuantity		= getUserInput(request.form("formQuantity"),4)

pOptionGroupDescrip1    = getUserInput(request.form("optionGroupDescrip1"), 100)
pOptionGroupDescrip2    = getUserInput(request.form("optionGroupDescrip2"), 100)

pOptionDescrip1		= getUserInput(request.form("optionDescrip1"), 500)
pPriceToAdd1		= getUserInput(request.form("priceToAdd1"), 500)  
pPercentageToAdd1	= getUserInput(request.form("percentageToAdd1"), 500)  

pOptionDescrip2		= getUserInput(request.form("optionDescrip2"), 500)
pPriceToAdd2		= getUserInput(request.form("priceToAdd2"), 500)  
pPercentageToAdd2	= getUserInput(request.form("percentageToAdd2"), 500)  

pHiddenIdOptions	= getUserInput(request.form("hiddenIdOptions"), 100)
pIdOptionGroup1		= getUserInput(request.form("idOptionGroup1"), 10) 
pIdOptionGroup2   	= getUserInput(request.form("idOptionGroup2"), 10)


if pListHidden<>"-1" then
  plistHidden="0"
end if

if pHotDeal<>"-1" then
  pHotDeal="0"
end if

if pActive="" then
  pActive="0"
end if

if pShowInHome="" then
  pShowInHome="0"
end if

' required fields
if pDescription="" or pDetails="" then
 response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Please enter description and details field.")
end if

' numbers
if isNumeric(pPrice)=false or isNumeric(pListPrice)=false or isNumeric(pbtobPrice)=false or isNumeric(pStock)=false or isNumeric(pDeliveringTime)=false or isNumeric(pWeight)=false then
 response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Please enter valid numbers for numeric values")
end if
  
' fill idOptions array for cleaning
arrayIdOptions  = split(pHiddenIdOptions,",") 

pDetails	= TwoSingleQ(pdetails)
pDescription	= TwoSingleQ(pdescription)

pPrice		= replace(pPrice, ",",".")
pListPrice	= replace(pListPrice, ",",".")
pBtoBPrice	= replace(pBtoBPrice, ",",".")
pCost		= replace(pCost, ",",".")

mySQL="UPDATE products SET sku='" &pSku& "', description='" &pDescription& "', details='" &pDetails& "', price=" &pPrice& ", listPrice=" &pListPrice& ", cost=" &pCost& ", imageUrl='" &pImageUrl& "', weight=" &pWeight& ", listHidden=" &pListHidden& ", hotDeal=" &pHotDeal& ",active=" &pActive& ", showInHome=" &pShowInHome& ", idSupplier= "&pIdSupplier& ", emailText='" &pEmailText& "', bToBPrice=" &pBToBPrice& ", deliveringTime=" &pDeliveringTime& ", formQuantity=" &pFormQuantity&", smallImageUrl='" &pSmallImageUrl& "' WHERE idProduct=" &pIdProduct
call updateDatabase(mySQL, rstemp, "comersus_backoffice_modifyProductExec.asp")  	

' replace stock

call replaceStock(pIdProduct, pStock)

' update categories

' delete previously assigned
mySQL="DELETE FROM categories_products WHERE idProduct=" &pIdProduct
call updateDatabase(mySQL, rstemp, "comersus_backoffice_modifyProductExec.asp")  	

' assign new ones
mySQL="INSERT INTO categories_products (idProduct, idCategory) VALUES (" &pIdProduct& "," &pIdCategory1& ")"
call updateDatabase(mySQL, rstemp, "comersus_backoffice_modifyProductExec.asp")  	

if pIdCategory2<>"" then
  mySQL="INSERT INTO categories_products (idProduct, idCategory) VALUES (" &pIdProduct& "," &pIdCategory2& ")"
  call updateDatabase(mySQL, rstemp, "comersus_backoffice_modifyProductExec.asp")  	
end if

if pIdCategory3<>"" then
 mySQL="INSERT INTO categories_products (idProduct, idCategory) VALUES (" &pIdProduct& "," &pIdCategory3& ")" 
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_modifyProductExec.asp")  	
end if

' remove old variations
call removeVariations(pIdProduct, pIdOptionGroup1, pIdOptionGroup2, arrayIdOptions)

' assign new variations

call addVariations(pOptionGroupDescrip1, pOptionDescrip1, pPriceToAdd1, pPercentageToAdd1, pIdProduct)
call addVariations(pOptionGroupDescrip2, pOptionDescrip2, pPriceToAdd2, pPercentageToAdd2, pIdProduct)

call closeDb()

response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Item modified")
%>
