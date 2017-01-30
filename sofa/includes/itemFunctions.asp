<%
function isRental(idProduct)
 
 dim rsTempRental
 
 mySQL="SELECT rental FROM products WHERE idProduct="& pIdProduct 

 call getFromDatabase(mySQL, rsTempRental, "isRental")
 
 isRental=0
 
 if not rsTempRental.eof then
  if rsTempRental("rental")=-1 then isRental=-1
 end if

end function

function isDonation(idProduct)
 
 dim rsTempDon
 
 mySQL="SELECT isDonation FROM products WHERE idProduct="& pIdProduct 

 call getFromDatabase(mySQL, rsTempDon, "isRental")
 
 isDonation=0
 
 if not rsTempDon.eof then
  if rsTempDon("isDonation")=-1 then isDonation=-1
 end if

end function


function itHasDiscountPerQty(idProduct)
 
 dim rsTempDPQ
 
 mySQL="SELECT idDiscountperquantity FROM discountsPerQuantity WHERE idProduct=" &pIdProduct

 call getFromDatabase (mySql, rsTempDPQ, "itHasDiscountPerQty")

 if not rsTempDPQ.eof then
  itHasDiscountPerQty=-1
 else
  itHasDiscountPerQty=0
 end if

end function

function getIdAuction(pIdProduct)

 ' check for auctions

 dim rsTempAuction
 getIdAuction=0

 mySQL="SELECT idAuction FROM auctions WHERE active=-1 AND idProduct=" &pIdProduct

 call getFromDatabase (mySql, rsTempAuction, "ViewItem")

 if not rsTempAuction.eof then  
  getIdAuction=rsTempAuction("idAuction")
 else
  getIdAuction=0
 end if

end function

function isBundleMain(idProduct)
 
 dim rsTempIBM
 
 mySQL="SELECT isBundleMain FROM products WHERE idProduct=" &pIdProduct

 call getFromDatabase (mySql, rsTempIBM, "itHasDiscountPerQty")

 if not rsTempIBM.eof then
  isBundleMain=rsTempIBM("isBundleMain")
 else
  isBundleMain=0
 end if

end function

function getRateReview(pIdProduct)
  
  dim mysql, rsTempRv
  
  mySQL="SELECT SUM(stars) AS sumStars, COUNT(*) AS countReviews FROM reviews WHERE idProduct=" &pIdProduct& " AND active=-1"

  call getFromDatabase (mySql, rsTempRv, "viewItem")
		
  if not rsTempRv.eof then
  
     if Cint(rsTempRv("countReviews"))<>0 then	   
	   getRateReview          	= cSng(rsTempRv("sumStars"))/cInt(rsTempRv("countReviews"))
      else
	   getRateReview		= 0
     end if
      
   else
     getRateReview			= 0
  end if
  
end function

Function getOptionsGroups(pIdProduct)
		
	dim pIdOptionGroup, htmlToPrint, rsTempGEO1, rsTempGEO2
 	
 	htmlToPrint=""
 	
 	pProductPrice = getPrice(pIdProduct, pIdCustomerType, pIdCustomer)
 	
	' get optionsGroups assigned
	
	mySQL = "SELECT idOptionGroup FROM optionsGroups_products WHERE idProduct=" & pIdProduct
	
	call getFromDatabase (mySql, rsTempGEO1, "itemFunctions")
	
	' iterate through optionsGroups
	do while not rsTempGEO1.eof
		pType=""
		pIdOptionGroup = rsTempGEO1("idOptionGroup")
		pOptionsGroupsLink=""
		
		' get options inside current optionGroup
		mySQL = "SELECT options.priceToAdd, options.percentageToAdd, options.optionDescrip, options.idOption, optionsGroups.optionGroupDesc, type, options.imageUrl, optionsGroups.optionLink FROM optionsGroups, options_optionsGroups, options WHERE optionsGroups.idOptionGroup=" & Cstr(pIdOptionGroup) & " AND optionsGroups.idOptionGroup=options_optionsGroups.idOptionGroup AND options.idOption=options_optionsGroups.idOption ORDER BY options.optionDescrip"
		
		call getFromDatabase (mySql, rsTempGEO2, "itemFunctions")		        	
		if not rsTempGEO2.eof then
    		pType		 	=lcase(rsTempGEO2("type"))											
    		pOptionGroupDesc 	=rsTempGEO2("optionGroupDesc")		
    		pOptionsGroupsLink	=rsTempGEO2("optionLink")		    		            
            
            if pOptionsGroupsLink<>"" then
             htmlToPrint=htmlToPrint&"<br><b><a href='"&pOptionsGroupsLink&"'>"& pOptionGroupDesc&"</a></b><br>"             
             else
             htmlToPrint=htmlToPrint&"<br><b>"& pOptionGroupDesc&"</b><br>"
            end if
                        
            
        end if		
	
		if pType="d" then
		 htmlToPrint = htmlToPrint & "<select name='idOptions" & pIdOptionGroup & "'>" & vbCrLf
		end if			
		
		if pType="m" then
		 htmlToPrint = htmlToPrint & "<select name='idOptions" & pIdOptionGroup & "' multiple>" & vbCrLf
		end if
			
					
		 ' drop down optionals				
									
		 do while not rsTempGEO2.eof
		 
		 	pImageUrl	 =rsTempGEO2("imageUrl")
		 	
		        if pType="d" or pType="m" then
			 htmlToPrint=htmlToPrint &"<option value='" & rsTempGEO2("idOption")  & "'>"
			end if			
			
			if pType="r" then
			 htmlToPrint=htmlToPrint &"<input type=radio name=idOptions" & pIdOptionGroup & " value=" &rsTempGEO2("idOption")& ">"
			end if
			
			if pType="c" then
			    htmlToPrint=htmlToPrint &"<input type=checkbox name=idOptions" & pIdOptionGroup & " value=" &rsTempGEO2("idOption")& ">"
			end if
			
		 	htmlToPrint=htmlToPrint & rsTempGEO2("optionDescrip")			
				
			' show price only if >0
				
			if rsTempGEO2("priceToAdd") > 0 Then
			 htmlToPrint=htmlToPrint &" "&pCurrencySign & money(rsTempGEO2("priceToAdd"))
			else
			    if rsTempGEO2("percentageToAdd") > 0 Then
			        htmlToPrint=htmlToPrint &" "&pCurrencySign & money(rsTempGEO2("percentageToAdd")*pProductPrice/100)
			    end if 
			End If												
			
			if pType="d" or pType="m" then
			 htmlToPrint=htmlToPrint &"</option>" & vbCrLf
			end if
			
			if pImageUrl<>"" then
			 htmlToPrint=htmlToPrint & "<img src='catalog/" &pImageUrl& "'>"
			end if
			
			if pType="r" or pType="c" then
			 htmlToPrint=htmlToPrint &"<br>" & vbCrLf
			end if
			
		  rsTempGEO2.movenext
		 loop		
		
         if pType="d" or pType="m" then			
    		 htmlToPrint=htmlToPrint &vbCrLf & "</select>"					
	     end if
		
		 if pType<>"" then
		    htmlToPrint=htmlToPrint &"<br>"
		 end if
         rsTempGEO1.movenext
	loop			        
	        
	getOptionsGroups=htmlToPrint
End Function

function getSupplierName(idSupplier)
 
 dim rsTempGS
 
 mySQL="SELECT supplierName FROM suppliers WHERE idSupplier="& pIdSupplier 

 call getFromDatabase(mySQL, rsTempGS, "isRental")
 
 if not rsTempGS.eof then
  getSupplierName=rsTempGS("supplierName")
 else
  getSupplierName="-"
 end if

end function

function getStock(pIdProduct)

 dim rsTempGS
 
 getStock=0
 
 mySQL="SELECT stock.stock FROM products, stock WHERE products.idStock=stock.idStock AND products.idProduct="& pIdProduct
 call getFromDatabase(mySQL, rsTempGS, "getStock")

 if not rsTempGS.eof then
  getStock=rsTempGS("stock")
 end if
end function

function getCompleteStock(pIdStore2)

 dim rsTempGS
 
 getCompleteStock=0
 
 mySQL="SELECT SUM(stock.stock) AS sumStock FROM products, stock WHERE products.idStock=stock.idStock AND products.idStore="& pIdStore2
 call getFromDatabase(mySQL, rsTempGS, "getStock")

 if not rsTempGS.eof then
  getCompleteStock=rsTempGS("sumStock")
 end if
 
end function

function getCompleteStockForSupplier(pIdStore2, pIdSupplier)

 dim rsTempGS
 
 getCompleteStockForSupplier=0
 
 mySQL="SELECT SUM(stock.stock) AS sumStock FROM products, stock WHERE products.idStock=stock.idStock AND products.idStore="& pIdStore2& " AND products.idSupplier="&pIdSupplier
 call getFromDatabase(mySQL, rsTempGS, "getStock")

 if not rsTempGS.eof then
  getCompleteStockForSupplier=rsTempGS("sumStock")
 end if
 
end function

function updateStock(pIdProduct, pQuantity)

 dim rsTempGS, pIdStock  
 
 pIdStock=0
 
 ' get idStock
 mySQL="SELECT idStock FROM products WHERE idProduct="& pIdProduct
 call getFromDatabase(mySQL, rsTempGS, "getStock")

 if not rsTempGS.eof then
  pIdStock=rsTempGS("idStock")
 end if
 
 ' update
 
 mySQL="UPDATE stock SET stock=stock+ " &pQuantity& " WHERE idStock="& pIdStock
 call updateDatabase(mySQL, rsTempGS, "getStock")

end function


sub replaceStock(pIdProduct, pQuantity)

 dim rsTempGS, pIdStock  
 
 pIdStock=0
 
 ' get idStock
 mySQL="SELECT idStock FROM products WHERE idProduct="& pIdProduct
 call getFromDatabase(mySQL, rsTempGS, "getStock")

 if not rsTempGS.eof then
  pIdStock=rsTempGS("idStock")
 end if
 
 ' update
 
 mySQL="UPDATE stock SET stock=" &pQuantity& " WHERE idStock="& pIdStock
 call updateDatabase(mySQL, rsTempGS, "getStock")

end sub


sub createStock(pIdProduct, pStock)
  dim rstemp
    
  ' insert stock record
  mySQL="INSERT INTO stock (idProductMain, stock) VALUES (" &pIdProduct& "," &pStock& ")"    
   
  call updateDatabase(mySQL, rstemp, "createStock") 
  
  ' retrieve ID
  mySQL="SELECT MAX(idStock) AS maxIdStock FROM stock WHERE idProductMain=" &pIdProduct
  
  call getFromDatabase(mySQL, rstemp, "comersus_backoffice_addproductexec.asp") 
  
  pIdStock = rstemp("maxIdStock")
  
  ' update product record
  mySQL="UPDATE products SET idStock=" &pIdStock&" WHERE idProduct="&pIdProduct 
  call updateDatabase(mySQL, rstemp, "createStock") 
  
end sub

sub assignStockId(pIdProduct, pIdProductStock)
  dim rstemp
    
  ' get stock ID
  mySQL="SELECT idStock FROM stock WHERE idProductMain=" &pIdProductStock  
  call getFromDatabase(mySQL, rstemp, "comersus_backoffice_addproductexec.asp") 
  
  pIdStock = rstemp("idStock")
    
  mySQL="UPDATE products SET idStock="&pIdStock&" WHERE idProduct="&pIdProduct   
   
  call updateDatabase(mySQL, rstemp, "createStock")     
    
end sub

sub getImage(pImageUrl, pImageUrl2, pImageUrl3, pImageUrl4, pImage)

if pImageUrl<>"" then
    
    if pImage=1 then
     %>
     
     <%if pZoomItemImage="-1" then%>        
        <div><img id="myimg" alt="<%=pDescription%>" src="catalog/<%=pImageUrl%>"><br />
        <br><br>
	<a href="#" onclick="zoom('-','myimg',92/66)"><img src="images/ZoomOut.gif" border=0></a> 
	<a href="#" onclick="zoom('+','myimg',92/66)"><img src="images/ZoomIn.gif" border=0></a>
        </div> 
     <%else%>
      <img src='catalog/<%=pImageUrl%>' alt='<%=pDescription%>'>
     <%end if%>
     
     <%
    end if
    
    if pImage=2 then
     %>
     
     <%if pZoomItemImage="-1" then%>        
        <div><img id="myimg" alt="<%=pDescription%>" src="catalog/<%=pImageUrl2%>"><br />
        <br><br>
	<a href="#" onclick="zoom('-','myimg',92/66)"><img src="images/ZoomOut.gif" border=0></a> 
	<a href="#" onclick="zoom('+','myimg',92/66)"><img src="images/ZoomIn.gif" border=0></a>
        </div> 
     <%else%>
      <img src='catalog/<%=pImageUrl2%>' alt='<%=pDescription%>'>
     <%end if%>
          
     <%
    end if
    
    if pImage=3 then
     %>
     
     <%if pZoomItemImage="-1" then%>        
        <div><img id="myimg" alt="<%=pDescription%>" src="catalog/<%=pImageUrl3%>"><br />
        <br><br>
	<a href="#" onclick="zoom('-','myimg',92/66)"><img src="images/ZoomOut.gif" border=0></a> 
	<a href="#" onclick="zoom('+','myimg',92/66)"><img src="images/ZoomIn.gif" border=0></a>
        </div> 
     <%else%>
      <img src='catalog/<%=pImageUrl3%>' alt='<%=pDescription%>'>
     <%end if%>
     
     
     <%
    end if
    
    if pImage=4 then
     %>
     
     <%if pZoomItemImage="-1" then%>        
        <div><img id="myimg" alt="<%=pDescription%>" src="catalog/<%=pImageUrl4%>"><br />
        <br><br>
	<a href="#" onclick="zoom('-','myimg',92/66)"><img src="images/ZoomOut.gif" border=0></a> 
	<a href="#" onclick="zoom('+','myimg',92/66)"><img src="images/ZoomIn.gif" border=0></a>
        </div> 
     <%else%>
      <img src='catalog/<%=pImageUrl4%>' alt='<%=pDescription%>'>
     <%end if%>          
     
     <%
    end if
    
    if pImageUrl2<>"" then
     %><br><br><%=getMsg(672,"switch")%>&nbsp;<a href="comersus_viewItem.asp?idProduct=<%=pIdProduct%>&image=">1</a>&nbsp;<a href="comersus_viewItem.asp?idProduct=<%=pIdProduct%>&image=2">2</a><%
    end if
    if pImageUrl3<>"" then
     %>&nbsp;<a href="comersus_viewItem.asp?idProduct=<%=pIdProduct%>&image=3">3</a><%
    end if
    if pImageUrl4<>"" then
     %>&nbsp;<a href="comersus_viewItem.asp?idProduct=<%=pIdProduct%>&image=4">4</a><%
    end if
 else
    %><img src='catalog/imageNA.gif'><%              
 end if

end sub

function getCategoryLevel(idCategory, level)
  
   dim mySql2, rstempLevel      
   
  ' get parent
   mySql2="SELECT idParentCategory FROM categories WHERE idCategory=" &idCategory
 
   call getFromDatabase (mySql2, rstempLevel, "getCategoryLevel")      
   
   if rstempLevel("idParentCategory")>1 then
    level=level+1
    call getCategoryLevel(rstempLevel("idParentCategory"), level)
   else
    level=level+1        
   end if
    
   getCategoryLevel=level
         
end function

function itHasNoItemsInside(pIdCategory)
  
   dim mySql2, rstemp2
   
   itHasNoItemsInside=0
   
   mySql2="SELECT * FROM categories_products WHERE idCategory=" &pIdCategory

   call getFromDatabase (mySql2, rsTemp2, "itHasNoItemsInside")      
   
   if rstemp2.eof then
     itHasNoItemsInside=-1
   end if  
      
end function

function getCategoryId(pCategoryDesc)
  
   dim mySql2, rstemp2      
   
   mySql2="SELECT idCategory FROM categories WHERE categoryDesc='" &pCategoryDesc&"'"

   call getFromDatabase (mySql2, rsTemp2, "itHasNoItemsInside")      
   
   if not rstemp2.eof then
     getCategoryId=rstemp2("idCategory")
   else
     getCategoryId=0
   end if  
      
end function

function getCategoryPath(pIdCategory, pCategoryString)

 dim rsTemp1

 mySQL="SELECT c2.idCategory, c2.categoryDesc FROM categories c1, categories c2 WHERE c1.idParentCategory=c2.idCategory AND c1.idCategory="&pIdCategory&" AND c2.idCategory>1"
 call getFromDatabase(mySQL, rsTemp1, "itemFunctions.asp") 

 if not rsTemp1.eof  then
   
   pCategoryString=pCategoryString &"|"& rstemp1("categoryDesc")
   
   if rstemp1("idCategory")>1 then
    call getCategoryPath(rstemp1("idCategory"),pCategoryString)
   end if
 end if	

end function

function isCategoryLeaf(idCategory)

 isCategoryLeaf=0
 mySQL="SELECT idCategory FROM categories WHERE idParentCategory="&idCategory
 call getFromDatabase(mySQL, rsTemp1, "comersus_backoffice_addproductform.asp") 

 if rsTemp1.eof then
   isCategoryLeaf=1
 end if	

end function

sub addVariations(pOptionGroupDescrip, pOptionDescrip, pPriceToAdd, pPercentageToAdd, pIdProduct)

 dim rstemp, pIndex  
 
 if pOptionGroupDescrip<>"" then
     
  arrayOptions  	=split(pOptionDescrip,",")
  arrayPrices   	=split(pPriceToAdd,",")         
  arrayPercentage   	=split(pPercentageToAdd,",")         

  ' verification
  if arrayOptions(0)="" then
    response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("If you need variations please fill every variation in order")
  end if
  
   '  add option group 
   mySQL="INSERT INTO optionsGroups (optionGroupDesc, type) VALUES ('" &pOptionGroupDescrip & "','D')"
   call  updateDatabase(mySQL, rstemp, "itemFunctions.asp")  
   
   ' retrieve optionGroupId
    
   mySQL="SELECT MAX(idOptionGroup) AS maxIdOptionGroup FROM optionsGroups WHERE optionGroupDesc='" & pOptionGroupDescrip& "'"
   call  getFromDatabase(mySQL, rstemp, "itemFunctions.asp") 
   
   if rstemp.eof then        
     response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Cannot get option group ID")
   end if  
   
   pIdOptionGroup=rstemp("maxIdOptionGroup")       
   
   ' add options
   for f=0 to 2
    
   if arrayOptions(f)<>"" and (arrayPrices(f) <> "" or arrayPercentage(f) <> "") then
	    mySQL="INSERT INTO options (optionDescrip, priceToAdd, percentageToAdd) VALUES ('" & arrayOptions(f) & "'," & arrayPrices(f)& "," & arrayPercentage(f) & ")"     
	    call  updateDatabase(mySQL, rstemp, "itemFunctions.asp")  
	    
	    ' get idOption
	   
	    mySQL="SELECT MAX(idOption) AS maxIdOption FROM options WHERE optionDescrip='" & arrayOptions(f) & "'"
	     
	    call  getFromDatabase(mySQL, rstemp, "itemFunctions.asp") 
	    
	    if rstemp.eof then        
	         response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Cannot get option ID")
	    end if     
	    
	    pIdOption=rstemp("maxIdOption")   
        
		' assign to group
		mySQL="INSERT INTO options_optionsGroups (idOption, idOptionGroup) VALUES (" & pIdOption & "," & pIdOptionGroup & ")"
		 
		call  updateDatabase(mySQL, rstemp, "itemFunctions.asp")  
        
   end if
    
   next
   
   ' assign group to product
    mySQL="INSERT INTO optionsGroups_products (idOptionGroup, idProduct) VALUES (" & pIdOptionGroup & "," & pIdProduct & ")"
     
    call  updateDatabase(mySQL, rstemp, "itemFunctions.asp")  
    
  end if ' filled
  
end sub


sub loadProductVariations(pIdProduct, arrayOptions1, arrayOptions2, pHiddenIdOptions, pOptionDescrip1, pOptionDescrip2, pIdOptionGroup1, pIdOptionGroup2)
 
  dim rstemp
  
  ' retrieve option groups
  
  mySQL="SELECT optionsGroups.idOptionGroup, optionGroupDesc FROM optionsGroups_products, optionsGroups WHERE optionsGroups.idOptionGroup=optionsGroups_products.idOptionGroup AND idProduct=" &pIdProduct
  call getFromDatabase(mySQL, rstemp, "getOptions") 
  
  pIdOptionGroup1=0
  pIdOptionGroup2=0  
  
  do while not rstemp.eof 
    
    if pIdOptionGroup1=0 then
      pOptionDescrip1=rstemp("optionGroupDesc")
      pIdOptionGroup1=rstemp("idOptionGroup")
    else
      pOptionDescrip2=rstemp("optionGroupDesc")
      pIdOptionGroup2=rstemp("idOptionGroup")
    end if
    
  rstemp.movenext
  loop  
  
  ' fill option arrays with default values
  for f=0 to 2
   arrayOptions1(0,f)=""
   arrayOptions1(1,f)=0
   arrayOptions1(2,f)=0
   arrayOptions2(0,f)=""
   arrayOptions2(1,f)=0   
   arrayOptions2(2,f)=0
  next
  
   
  ' retrieve options
  pHiddenIdOptions=""
  
  if pIdOptionGroup1<>0 then
      
   mySQL="SELECT optionDescrip, priceToAdd, percentageToAdd, options.idOption FROM options, options_optionsGroups WHERE idOptionGroup=" & pIdOptionGroup1 & " AND options.idOption=options_optionsGroups.idOption"
   call getFromDatabase(mySQL, rstemp, "getOptions") 
   f=0
   
   do while not rstemp.eof
     arrayOptions1(0,f)	= trim(rstemp("optionDescrip"))
     arrayOptions1(1,f)	= rstemp("priceToAdd")
     arrayOptions1(2,f)	= rstemp("percentageToAdd")
     pHiddenIdOptions	= pHiddenIdOptions & rstemp("idOption") &"," 
    f=f+1
   rstemp.movenext
   loop   
   
  end if
  
  if pIdOptionGroup2<>0 then
      
   mySQL="SELECT optionDescrip, priceToAdd, percentageToAdd, options.idOption FROM options, options_optionsGroups WHERE idOptionGroup=" & pIdOptionGroup2 & " AND options.idOption=options_optionsGroups.idOption"
   call getFromDatabase(mySQL, rstemp, "getOptions") 
   
   f=0
   do while not rstemp.eof
     arrayOptions2(0,f)	= trim(rstemp("optionDescrip"))
     arrayOptions2(1,f)	= rstemp("priceToAdd")
     arrayOptions2(2,f)	= rstemp("percentageToAdd")
     pHiddenIdOptions	= pHiddenIdOptions & rstemp("idOption") &"," 
    f=f+1
   rstemp.movenext
   loop   
   
  end if

end sub

sub removeVariations(pIdProduct, pIdOptionGroup1, pIdOptionGroup2, arrayIdOptions)

 dim rstemp
 
  mySQL = "DELETE FROM optionsGroups_products WHERE idProduct=" &pIdProduct
  call updateDatabase(mySQL, rstemp, "ItemFunctions, modify variations")  	    
   
  mySQL = "DELETE FROM options_optionsGroups WHERE idOptionGroup=" & pIdOptionGroup1 &" OR idOptionGroup=" & pIdOptionGroup2
  call updateDatabase(mySQL, rstemp, "ItemFunctions, modify variations")  	    
   
  mySQL = "DELETE FROM optionsGroups WHERE idOptionGroup=" & pIdOptionGroup1 &" OR idOptionGroup=" & pIdOptionGroup2
  call updateDatabase(mySQL, rstemp, "ItemFunctions, modify variations")  	    
   
  mySQL = "DELETE FROM options WHERE idOption=99999999"
  
  for f=0 to uBound(arrayIdOptions)
   if arrayIdOptions(f)<>"" then
    mySQL=mySQL & " OR idOption=" & arrayIdOptions(f)    
   end if 
  next
  
  call updateDatabase(mySQL, rstemp, "ItemFunctions, modify variations")  	
  
end sub

function getCategoryStart(pIdStore)
 
 dim mysql, rsTempGC
 
 ' get category start from stores
 mySQL="SELECT idCategoryStart FROM stores WHERE idStore=" &pIdStore
 call getFromDatabase (mySql, rsTempGC, "itemFunctions")

 if not rsTempGC.eof then
  getCategoryStart=rsTempGC("idCategoryStart")
 else
  getCategoryStart=1
 end if

end function

function getCategoryDescription(pIdCategory)
 
 dim mysql, rsTempGC
 
 ' get category start from stores
 mySQL="SELECT categoryDesc FROM categories WHERE idCategory=" &pIdCategory
 call getFromDatabase (mySql, rsTempGC, "itemFunctions")

 if not rsTempGC.eof then
  getCategoryDescription=rsTempGC("categoryDesc")
 else
  getCategoryDescription="N/A"
 end if

end function


function getPrice(pIdProduct, pIdCustomerType, pIdCustomer)
 
 dim pPrice, pBtoBPrice
 
 pPrice=0
 pBtoBPrice=0
 
 ' get price
 mySQL="SELECT price, bToBPrice FROM products WHERE idProduct=" &pIdProduct
 
 call getFromDatabase(mySQL, rstempSPrice, "specialPrices")  
 
 if not rstempSPrice.eof then
  pPrice	=rstempSPrice("price")
  pBtoBPrice	=rstempSPrice("bToBPrice")
 end if
 
 if pIdCustomer<>0 then
 
  ' check if the customer has a special price for that product 
  mySQL="SELECT specialPrice FROM customer_specialPrices WHERE idProduct=" &pIdProduct& " AND idCustomer="&pIdCustomer
 
  call getFromDatabase(mySQL, rstempSPrice, "specialPrices")  

  if not rstempSPrice.eof then
   pPrice		=rstempSPrice("specialPrice")
   pBtoBPrice		=rstempSPrice("specialPrice") 
  end if  
   
 end if

 if pIdCustomerType=1 then
   getPrice=pPrice
 else  
   getPrice=pBtoBPrice  
 end if
   
end function

function getPriceByQty(pIdProduct, pIdCustomerType, pIdCustomer, quantity)
 
 dim pPrice, pBtoBPrice
 
 pPrice=0
 pBtoBPrice=0
 
 ' get price
 mySQL="SELECT price, bToBPrice FROM products WHERE idProduct=" &pIdProduct
 
 call getFromDatabase(mySQL, rstempSPrice, "specialPrices")  
 
 if not rstempSPrice.eof then
  pPrice	=rstempSPrice("price")
  pBtoBPrice	=rstempSPrice("bToBPrice")
 end if
 
 ' discount per qty
 mySQL="SELECT discountPerUnit FROM discountsPerQuantity WHERE idProduct=" &pIdProduct& " AND quantityFrom<=" &quantity& " AND quantityUntil>=" &quantity
  
 call getFromDatabase(mySQL, rsTempDPQ, "cartRecalculate")          	  

 if not rsTempDPQ.eof then  	
  	pDiscountPerUnit 	= rsTempDPQ("discountPerUnit")    
  	pPrice			= pPrice-pDiscountPerUnit
 	pBtoBPrice		= pBtoBPrice-pDiscountPerUnit
 
 else
  	pDiscountPerUnit  	= 0
 end if	
  
 ' special price
 if pIdCustomer<>0 then
 
  ' check if the customer has a special price for that product 
  mySQL="SELECT specialPrice FROM customer_specialPrices WHERE idProduct=" &pIdProduct& " AND idCustomer="&pIdCustomer
 
  call getFromDatabase(mySQL, rstempSPrice, "specialPrices")  

  if not rstempSPrice.eof then
   pPrice		=rstempSPrice("specialPrice")
   pBtoBPrice		=rstempSPrice("specialPrice") 
  end if  
   
 end if

 if pIdCustomerType=1 then
   getPriceByQty=pPrice
 else  
   getPriceByQty=pBtoBPrice  
 end if
   
end function

%>