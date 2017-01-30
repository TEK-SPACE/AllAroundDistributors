<%
' Comersus BackOffice 
' e-commerce ASP Open Source
' CopyRight Rodrigo S. Alhadeff, Comersus
' February 2006
' http://www.comersus.com 
' Import products tool for Comersus store
%>
<!--#include file="comersus_backoffice_adminVerify.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/itemFunctions.asp"-->    
<!--#include file="../includes/settings.asp"-->    
<!--#include file="../includes/stringFunctions.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 

<!--#include file="header.asp"--> 

<b>Products import</b>
<br><br>
<%	

' on error resume next

if pBackOfficeDemoMode=-1 then
 response.redirect "comersus_backoffice_message.asp?message="&Server.Urlencode("Function disabled in demo mode.")
end if

Const ForReading = 1, ForWriting = 2

Dim mySQL, conntemp, rstemp, categoryDescription, pIdCategory, pNewSupplierId, pNewCategoryId

call openDb()    

pIdStore    = getUserInput(request.Form("idStore"),3)
pActive		= getUserInput(request.Form("active"),2)

if pActive ="on" then
	pActive = "-1"
else	
	pActive = "0"
end if

mySQL="SELECT idCategoryStart FROM stores WHERE idStore="&pIdStore
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_importProductsExec.asp") 

if not rstemp.eof then
    pIdCategoryStart=rstemp("idCategoryStart")	
end if


pNewSupplierId = 0
pNewCategoryId = 0
pCounter=0 

pFields =  "(idStore, active, sku, description, details, price, btobPrice, imageUrl, smallImageUrl , listHidden, weight, cost, sales , length, width, height, emailText, keywords, hotDeal, isDonation, hasPersonalization, idSupplier)" 

pCounterLine = 1

if request.Form("InfoToImport")="" then
   response.redirect "comersus_backoffice_message.asp?message="&Server.Urlencode("You must specify the information to be imported.")
end if
pArrayLine = split(request.Form("InfoToImport"),vbCrlf) 
response.write "<br><br><b>Trying to Import from text field.</b>"
For x = 0 to UBOUND(pArrayLine)

    pLine   = pArrayLine(x)
	pArray = split (pLine, vbTab)

    response.write "<br><br>Trying to Import line#"&pCounterLine&"..."
	if Ubound(pArray)= 20 then
        pSku        = "'" & getUserInput(pArray(0),50)& "'"

        pDesc       = "'" & formatForDb(pArray(1)) & "'"
        pDetails    = "'" & formatForDb(pArray(2)) & "'"
        
        pPrice      = formatNumberForDb(pArray(3))     
        pBtobPrice  = formatNumberForDb(pArray(4))  

        pImage      = "'" & getUserInput(pArray(5),150) & "'"
        pSmallImage = "'" & getUserInput(pArray(6),150) & "'"

        pListHidden = pArray(7)
        pCost       = formatNumberForDb(pArray(8))     
        pSales      = getUserInput(pArray(9),30)     
        pWeight     = getUserInput(pArray(10),30)     
        pLength     = getUserInput(pArray(11),30)     
        pWidth      = getUserInput(pArray(12),30) 
        pHeight     = getUserInput(pArray(13),30) 
		pCategory   = getUserInput(pArray(14),200) 
		pSupplier	= getUserInput(pArray(15),200) 
		
		pIdCategory = getCategory(pCategory)
		
		pIdSupplier = getSupplier(pSupplier)
		
        pEmailText  = "'" & getUserInput(pArray(16),150) & "'"
        pKeywords   = "'" & getUserInput(pArray(17),150) & "'"
        pClearance  = pArray(18)
        pIsDonation	= pArray(19)
        pHasPersonalization = "'" & getUserInput(pArray(20),150) & "'"

		if pPrice = "" then pPrice = "0"
		if pBtoBPrice = "" then pBtoBPrice = "0"
			
		if pListhidden<>"-1" then pListhidden="0"
		if pClearance<>"-1" then pClearance="0"
		if pHasPersonalization<>"-1" then pHasPersonalization="0"
		if pShowInHome<>"-1" then pShowInHome="0"
		if pFreeShipping<>"-1" then pFreeShipping="0"
		if pIsDonation<>"-1" then pIsDonation="0"

		pFields =  "(idStore, active, sku, description, details, price, btobPrice, imageUrl, smallImageUrl , listHidden, weight, cost, sales , length, width, height, emailText, searchKeywords, hotDeal, isDonation, hasPersonalization, idSupplier)" 
        pValues =  pIdStore & "," & pActive & "," & pSku & "," & pDesc & "," & pDetails  & "," & pPrice  & "," & pBtobPrice  & "," & pImage  & "," & pSmallImage & "," & pListHidden & "," & pWeight & "," & pCost & "," & pSales & "," & pLength & "," & pWidth & "," & pHeight & "," & pEmailText & "," & pKeywords & "," & pClearance & "," & pIsDonation & "," & pHasPersonalization & "," & pIdSupplier 
        
        mySQL = "INSERT INTO products " & pFields & " VALUES (" & pValues & ")"
		call updateDatabase(mySQL, rstemp, "comersus_backoffice_importProductsExec.asp") 
		
		pCounter = pCounter + 1
		
		pIdProduct = GetMaxProduct()

		' create stock 
		call createStock(pIdProduct, 0)
		
		' insert category - product relationship

		mySQL = "INSERT INTO categories_products (idCategory, idProduct) VALUES (" & pIdCategory &", " & pIdProduct &")"
		call updateDatabase(mySQL, rstemp, "comersus_backoffice_importProductsExec.asp") 
       
 		if err.number <> 0 then 		 	  			
  			Response.write " <br>-Error found!"
        end if			

	else
        response.write " Wrong fields quantity. Check the file format."
    end if       
	pCounterLine = pCounterLine + 1
Next


call closeDb()

function getCategory(pCategory)
    if pCategory="" then 
		' create a new category
	    pCategoryAdded  = createCategory("", pIdCategoryStart)
	    pIdCategory     = pCategoryAdded 
    else
        mySQL="SELECT idCategory FROM categories where categoryDesc = '" & pCategory &"'"
        call getFromDatabase(mySQL, rstemp, "comersus_backoffice_importProductsExec.asp") 

        if not rstemp.eof then
            pIdCategory = rstemp("idCategory")
            
            if isCategoryLeaf(pIdCategory) = 0 then
            	'the category has subcategories. Add new category 
            	pIdCategory     = createCategory("", pIdCategory)
            end if

        else
            ' create a new category
            pIdCategory     = createCategory(pCategory, pIdCategoryStart)
        end if
    end if
    getCategory = pIdCategory
end function

function getSupplier (pSupplier)
    ' verify supplier to insert.
    if pSupplier = "''" or pSupplier = "" then 
		getSupplier = createSupplier("")
    else
        ' search supplier
        mySQL="SELECT idSupplier FROM suppliers where supplierName = '" & pSupplier & "'"
        call getFromDatabase(mySQL, rstemp, "comersus_backoffice_importProductsExec.asp") 

        if not rstemp.eof then
            getSupplier  = rstemp("idSupplier")
        else
            'insert new supplier.
            getSupplier = createSupplier(pSupplier)
        end if
    end if            
end function

function createSupplier(pSupplier)
    if pNewSupplierId <> 0 and (pSupplier = "''" or pSupplier = "") then
    	createSupplier = pNewSupplierId
    else	
		if pSupplier = "''" or pSupplier = "" then 
	        mySQL = "INSERT INTO suppliers (supplierName) VALUES ('Import Supplier"&randomNumber(9999)&"')"
		else
			mySQL = "INSERT INTO suppliers (supplierName) VALUES ('" & pSupplier & "')"
		end if
	 	set rsTemp=connTemp.execute(mySQL) 
	 	mySQL = "SELECT MAX(idSupplier) AS SupplierID FROM suppliers "
	 	set rsTemp=connTemp.execute(mySQL) 

		pNewSupplierId = rsTemp("SupplierID") 
		createSupplier = rsTemp("SupplierID") 
	end if	
end function

function createCategory(categoryDescription, pIdCategoryStart)
    if pNewCategoryId <> 0 and categoryDescription = "" then
		createCategory 	= pNewCategoryId   
	else
		' insert a new category for products to be imported
		
		if categoryDescription="" then 
			categoryDescription = "Import Category "& randomNumber(9999)
		end if	
		
		mySQL="INSERT INTO categories (categoryDesc, idParentCategory, active) VALUES ('" &categoryDescription& "',"&pIdCategoryStart&",-1)"	
		call updateDatabase(mySQL, rstemp, "comersus_backoffice_importProductsExeca.asp") 
		
		' get idCategory for inserted category
		
		mySQL="SELECT Max(idCategory) as IdCategoryAdded FROM categories WHERE categoryDesc='" &categoryDescription & "'"
		
		call getFromDatabase(mySQL, rstemp, "comersus_backoffice_importProductsExec.asp") 
		
		pIdCategory  = rstemp("IdCategoryAdded")
		
		pNewCategoryId = pIdCategory
		createCategory=pIdCategory	
	end if
end function

function GetMaxProduct()
	mySQL="SELECT MAX(idProduct) AS idMax FROM products "
	call getFromDatabase(mySQL, rsMax, "comersus_backoffice_importProductsExec.asp") 				
	GetMaxProduct=rsMax("idMax")
end function

%>

<br><br><%=pCounter%> records imported.<br>
<br><br>
<!--#include file="footer.asp"--> 