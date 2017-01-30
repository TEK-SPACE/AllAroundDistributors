<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
pPageLevel=Cint(0)
%>

<!--#include file="comersus_backoffice_adminVerify.asp"-->   
<!--#include file="comersus_backoffice_functions.asp"-->   
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/dateFunctions.asp"--> 
<!--#include file="../includes/settings.asp"-->    
<!--#include file="../includes/stringFunctions.asp"--> 
<!--#include file="../includes/miscFunctions.asp"--> 
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 

<% 
' on error resume next

dim mySQL, conntemp, rstemp

pDateFormat 		= getSettingKey("pDateFormat")

' form parameters
pSku			= getUserInput(request.form("sku"),20)
pDescription		= getUserInput(request.form("description"),100)
pDetails		= getUserInput(request.form("details"),2000)
pIdcategory1		= getUserInput(request.form("idCategory1"),4)
pIdcategory2		= getUserInput(request.form("idCategory2"),4)
pIdcategory3		= getUserInput(request.form("idCategory3"),4)
pPrice			= getUserInput(request.form("price"),12)
pListPrice		= getUserInput(request.form("listPrice"),12)
pBtoBPrice		= getUserInput(request.form("btoBPrice"),12)
pCost			= getUserInput(request.form("cost"),12)
pImageUrl		= getUserInput(request.form("imageUrl"),100)
pSmallImageUrl		= getUserInput(request.form("smallImageUrl"),100)
pIdSupplier		= getUserInput(request.form("idSupplier"),4)
pActive			= getUserInput(request.form("active"),2)
pShowInHome		= getUserInput(request.form("showInHome"),2)
pListhidden		= getUserInput(request.form("listHidden"),2)
pHotDeal		= getUserInput(request.form("hotDeal"),2)
pWeight			= getUserInput(request.form("weight"),12)
pStock			= getUserInput(request.form("stock"),12)
pDeliveringTime		= getUserInput(request.form("deliveringTime"),10)
pEmailText		= getUserInput(request.form("emailText"),250)
pFormQuantity		= getUserInput(request.form("formQuantity"),4)
pOptionGroupDescrip1	= getUserInput(request.form("optionGroupDescrip1"),50)
pOptionGroupDescrip2	= getUserInput(request.form("optionGroupDescrip2"),50)
pOptionDescrip1		= getUserInput(request.form("optionDescrip1"), 500)
pPriceToAdd1		= getUserInput(request.form("priceToAdd1"), 500)  
pPercentageToAdd1		= getUserInput(request.form("percentageToAdd1"), 500)  

pOptionDescrip2		= getUserInput(request.form("optionDescrip2"), 500)
pPriceToAdd2		= getUserInput(request.form("priceToAdd2"), 500)  
pPercentageToAdd2		= getUserInput(request.form("percentageToAdd2"), 500)  


if pListHidden<>"-1" then pListhidden="0"
if pHotDeal<>"-1" then pHotDeal="0"
if pActive<>"-1" then pActive="0"
if pShowInHome<>"-1" then pShowInHome="0"

' create date 
pDateAdded=formatDate(Date())

' required fields
if pDescription="" or pDetails="" or pSku="" then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Please enter sku, description and details field")
end if

' numbers
if isNumeric(pPrice)=false or isNumeric(pListPrice)=false or isNumeric(pbtobPrice)=false or isNumeric(pStock)=false or isNumeric(pDeliveringTime)=false or isNumeric(pWeight)=false then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Invalid values for numeric fields")
end if

' check categories
if (pIdCategory1=pIdCategory2) or (pIdCategory1=pIdCategory3) or (pIdCategory2<>"" and pIdCategory2=pIdCategory3) then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("You cannot select the same category several times")
end if  
  
' replace comma
pPrice			= replace(pPrice,",",".")
pListPrice		= replace(pListPrice,",",".")
pBtoBPrice		= replace(pBToBPrice,",",".")
pCost			= replace(pCost,",",".")

' remove quotes
pSku			= formatForDb(pSku)
pDetails		= formatForDb(pDetails)
pDescription		= formatForDb(pDescription)
pOptionGroupDescrip1 	= formatForDb(pOptionGroupDescrip1)
pOptionGroupDescrip2 	= formatForDb(pOptionGroupDescrip2)

' insert product in to db
mySQL="INSERT INTO products (sku, description, details, price, listPrice, bToBPrice, cost, imageUrl, listhidden, weight, active, idSupplier, hotDeal, emailText, deliveringTime, formQuantity, smallImageUrl, showInHome, dateAdded, isBundleMain, idStore) VALUES ('" &pSku& "','" &pDescription& "','" & pDetails& "'," &pPrice& "," &pListPrice& "," &pBtoBPrice& "," &pCost& ",'" &pImageUrl& "'," &pListhidden& "," &pWeight& "," &pActive& "," &pIdSupplier& "," &pHotDeal& ",'" &pEmailText& "'," &pDeliveringTime& "," &pFormQuantity& ",'" &pSmallImageUrl& "'," &pShowInHome& ",'" &pDateAdded& "',0," &pIdStore& ")"

call  updateDatabase(mySQL, rstemp, "comersus_backoffice_addproductexec.asp") 

' get product inserted
mySQL="SELECT idProduct FROM products WHERE sku='" &pSku& "' AND price=" &pPrice& " AND idStore=" &pIdStore& " ORDER BY idProduct DESC"
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_addproductexec.asp") 

if rstemp.eof then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Cannot locate inserted item")
end if

pIdProduct = rstemp("idProduct")

' insert categories for the item
mySQL="INSERT INTO categories_products (idProduct, idCategory) VALUES (" &pIdProduct& "," &pIdCategory1& ")"
call updateDatabase(mySQL, rstemp, "comersus_backoffice_addproductexec.asp") 

if pIdCategory2<>"" then
  mySQL="INSERT INTO categories_products (idProduct, idCategory) VALUES (" &pIdProduct& "," &pIdCategory2& ")"
  call updateDatabase(mySQL, rstemp, "comersus_backoffice_addproductexec.asp") 
end if

if pIdCategory3<>"" then
  mySQL="INSERT INTO categories_products (idProduct, idCategory) VALUES (" &pIdProduct& "," &pIdCategory3& ")" 
  call updateDatabase(mySQL, rstemp, "comersus_backoffice_addproductexec.asp") 
end if

' create stock record
call createStock(pIdProduct, pStock)

' insert variations
call addVariations(pOptionGroupDescrip1, pOptionDescrip1, pPriceToAdd1, pPercentageToAdd1, pIdProduct)
call addVariations(pOptionGroupDescrip2, pOptionDescrip2, pPriceToAdd2, pPercentageToAdd2, pIdProduct)

response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Item added.")

call closeDb()

%>