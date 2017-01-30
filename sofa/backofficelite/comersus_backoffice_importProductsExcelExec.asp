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

<%

dim mySQL, conntemp, rstemp

call openDb()    

pIdStore    = 1

mySQL="SELECT idCategoryStart FROM stores WHERE idStore="&pIdStore
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_importProductsExec.asp") 

if not rstemp.eof then
    pIdCategoryStart=rstemp("idCategoryStart")	
end if

%><br>Starting...<br><%

writeExcelData()

%>


<!--#include file="footer.asp"--> 

<%


function writeExcelData()

Dim rs,sql,i

	
pListPrice=0
pImageUrl=""
pListhidden=0
pWeight=0
pActive=-1
pHotDeal=0
pEmailText=""
pDeliveringTime=0
pFormQuantity=10
pSmallImageUrl=""
pShowInHome=-1
pDateAdded=Date

sql = "SELECT * FROM [Import$];"

if runsql(sql,rs) then

 Do While Not rs.EOF

		      
 	For I = 0 To rs.Fields.Count - 1
			
		if rs.Fields.Item(I).Name="sku" then
		 pSku=rs.Fields.Item(I).Value		
		end if
		
		if rs.Fields.Item(I).Name="description" then
		 pDescription=rs.Fields.Item(I).Value	
		 if pDescription<>"" then	
		  pDescription=replace(pDescription,"'","")
		 end if
		end if							     		
                
                pDetails=pDescription
                				
		if rs.Fields.Item(I).Name="price" then
		 pPrice=rs.Fields.Item(I).Value		 
		 
		 if pPrice<>"" then
		  pPrice=replace(pPrice,",",".")
		 end if
		end if	 
		
	        if rs.Fields.Item(I).Name="btobprice" then
		 pBtoBprice=rs.Fields.Item(I).Value
		 if pBtoBprice<>"" then
		  pBtoBprice=replace(pBtoBprice,",",".")	 
		 end if
		end if
		
	        if rs.Fields.Item(I).Name="cost" then
		 pCost=rs.Fields.Item(I).Value		
		 if pCost<>"" then
		  pCost=replace(pCost,",",".")	 
		 end if
		end if
	
		if rs.Fields.Item(I).Name="sales" then
		 pSales=rs.Fields.Item(I).Value	
		 if pSales<>"" then	
		  pSales=replace(pSales,",",".")
		 end if
		end if
		
		if rs.Fields.Item(I).Name="category" then
		 pCategoryDesc=rs.Fields.Item(I).Value		 
		end if
		
		if rs.Fields.Item(I).Name="supplier" then
		 pSupplierName=rs.Fields.Item(I).Value		 
		end if
		
		if rs.Fields.Item(I).Name="keywords" then
		 pKeywords=rs.Fields.Item(I).Value		 
		end if
		         
	Next		        
	
	if pSku<>"" then
	 
	response.write "<br>Record: "&	pSku&" "&pDescription        
	
	 pIdCategory = getCategory(pCategoryDesc)
		
	 pIdSupplier = getSupplier(pSupplierName)


 	' insert item
	mySQL="SELECT idProduct FROM products WHERE sku='"&pSku&"'"

	call  getFromDatabase(mySQL, rstemp2, "comersus_backoffice_addproductexec.asp") 
	
	if rstemp2.eof then
			
	  ' insert item
	 mySQL="INSERT INTO products (sku, description, details, price, listPrice, bToBPrice, cost, imageUrl, listhidden, weight, active, idSupplier, hotDeal, emailText, deliveringTime, formQuantity, smallImageUrl, showInHome, dateAdded, isBundleMain, idStore) VALUES ('" &pSku& "','" &pDescription& "','" & pDetails& "'," &pPrice& "," &pListPrice& "," &pBtoBPrice& "," &pCost& ",'" &pImageUrl& "'," &pListhidden& "," &pWeight& "," &pActive& "," &pIdSupplier& "," &pHotDeal& ",'" &pEmailText& "'," &pDeliveringTime& "," &pFormQuantity& ",'" &pSmallImageUrl& "'," &pShowInHome& ",'" &pDateAdded& "',0," &pIdStore& ")"

	 call  updateDatabase(mySQL, rstemp, "comersus_backoffice_addproductexec.asp") 
	 
 	  pIdProduct = GetMaxProduct()

	  ' create stock 
	  call createStock(pIdProduct, 0)
	
	   '  insert category - product relationship

  	  mySQL = "INSERT INTO categories_products (idCategory, idProduct) VALUES (" & pIdCategory &", " & pIdProduct &")"
	  call updateDatabase(mySQL, rstemp, "comersus_backoffice_importProductsExec.asp") 
	 end if ' item is not there
	
	else
	 response.write "<br>Invalid SKU in row" 	
	end if ' sku valid

        rs.MoveNext
    Loop

    rs.Close

end if

Set rs = Nothing

end function




function runSQL(SQL,rs)

	on error resume next

	dim myrs

	set myRs = createobject("ADODB.recordset")
	myRs.Open SQL,"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("comersus.xls") & ";Extended Properties=""Excel 8.0;IMEX=1;""", 1, 3
	set rs = myRs
	if err then
		runSQL = false

		response.write err.description

	else
		runSQL = true
	end if

end function


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