<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' USA - 2006
' http://www.comersus.com 
' Details: cart functions
%>
<%
' findProduct
function findProduct(pIdDbSessionCart, pIdProduct)  
 
 ' check if the product is inside the cart
 mySQL="SELECT idProduct FROM cartRows WHERE idProduct=" &pIdProduct& " AND idDbSessionCart="&pIdDbSessionCart
 
 call getFromDatabase(mySQL, rstemp, "cartFunctions")  

 if not rstemp.eof then
  findProduct=-1
 else
  findProduct=0
 end if  
 
end function

function findProductPrice(pIdDbSessionCart, pIdCustomerType, pIdProduct)  
 
 ' check price of the product in the cart
 
 mySQL="SELECT quantity, unitPrice FROM cartRows WHERE idProduct=" &pIdProduct& " AND idDbSessionCart="&pIdDbSessionCart
 call getFromDatabase(mySQL, rsTemp10, "cartFunctions")  
 
 findProductPrice = Cdbl(rsTemp10("quantity") * rsTemp10("unitPrice"))                  

end function

function findProductQuantity(pIdDbSessionCart, pIdProduct)  
 
 ' retrieve quantity of the product in the cart
 
 mySQL="SELECT quantity FROM cartRows WHERE idProduct=" &pIdProduct& " AND idDbSessionCart="&pIdDbSessionCart
 call getFromDatabase(mySQL, rsTemp10, "cartFunctions")  

 findProductQuantity = rsTemp10("quantity")

end function

' count cart Rows
function countCartRows(pIdDbSessionCart)
 
 dim mySQL, rsTemp2
 
 ' check if the product is inside the cart
 mySQL="SELECT COUNT(idCartRow) AS howMany FROM cartRows WHERE idDbSessionCart="&pIdDbSessionCart

 call getFromDatabase(mySQL, rsTemp2, "cartFunctions")  

 if rsTemp2.eof then
  countCartRows=0
 else
  countCartRows=Cint(rsTemp2("howMany"))
 end if  
 
end function

' Cart total
function calculateCartTotal(pIdDbSessionCart)       
 
 dim mySql, rsTempCartTotal, rsTempCartTotal2, pTotal, pOptionsTotal
 
 pTotal=Cdbl(0)
 
 ' iterate through cart rows
 
 mySQL="SELECT idCartRow, quantity, unitPrice FROM cartRows WHERE idDbSessionCart="&pIdDbSessionCart  
 
 call getFromDatabase(mySQL, rsTempCartTotal, "cartFunctions")  
  
 do while not rsTempCartTotal.eof
 
  pUnitPrice		= Cdbl(rsTempCartTotal("unitPrice"))  
  pQuantity		= rsTempCartTotal("quantity")
  pIdCartRow		= rsTempCartTotal("idCartRow")    
     
  ' get Optionals value for that row    
  
  pOptionsTotal=Cdbl(0)
  
  mySQL="SELECT SUM(cartRowsOptions.priceToAdd) AS optionsTotal FROM options, cartRowsOptions WHERE options.idOption=cartRowsOptions.idOption AND cartRowsOptions.idCartRow="&pIdCartRow     
  
  call getFromDatabase(mySQL, rsTempCartTotal2, "cartFunctions")  
    
  if rsTempCartTotal2.eof then
   pOptionsTotal=Cdbl(0)   
  else
   if isNull(rsTempCartTotal2("optionsTotal")) then
    pOptionsTotal=Cdbl(0)   
   else
    pOptionsTotal=Cdbl(rsTempCartTotal2("optionsTotal"))
   end if
  end if                    
  
  pTotal=pTotal+pQuantity*(pUnitPrice+pOptionsTotal)      
  
  rsTempCartTotal.movenext
 loop   
  
 calculateCartTotal=pTotal
 
end function

' Cart Weight
function calculateCartWeight(pIdDbSessionCart)   
 
 ' sum the total weight of the cart
 mySQL="SELECT SUM(quantity*unitWeight) AS total FROM cartRows WHERE idDbSessionCart="&pIdDbSessionCart 
 
 call getFromDatabase(mySQL, rsTemp, "cartFunctions")    

 if rstemp.eof then
  calculateCartWeight=0
 else
 
  if not isNull(rstemp("total")) then
    calculateCartWeight=Cdbl(rstemp("total"))
  else
    calculateCartWeight=0
  end if
    
 end if  
   
 ' sum the weight of the optionals, if any
 
 mySQL="SELECT SUM(cartRows.quantity*options.weight) AS total FROM cartRows, cartRowsOptions, options WHERE cartRows.idCartRow=cartRowsOptions.idCartRow AND cartRowsOptions.idOption=options.idOption AND idDbSessionCart=" &pIdDbSessionCart  
 
 call getFromDatabase(mySQL, rsTemp, "cartFunctions")    
 
 pTotalCartWeight=rstemp("total") 
 
 if not rstemp.eof and not isNull(pTotalCartWeight) then
    calculateCartWeight=calculateCartWeight+Cdbl(pTotalCartWeight)  
 end if  
   
end function

' Cart Product Quantity
function calculateCartQuantity(pIdDbSessionCart)
 
 dim mySQL, rstemp2  
     
 ' sum
 mySQL="SELECT SUM(quantity) AS qTotal FROM cartRows WHERE idDbSessionCart="&pIdDbSessionCart 
 
 call getFromDatabase(mySQL, rsTemp2, "cartFunctions")        

 if rstemp2.eof then
  calculateCartQuantity=0
 else    
  if isNull(rstemp2("qTotal")) then
   calculateCartQuantity=0
  else
   calculateCartQuantity=Cdbl(rstemp2("qTotal"))
  end if
 end if     
end function

' Check session lost
function checkSessionLost() 
 if pIdDbSession="" or isNull(pIdDbSession) then
  checkSessionLost = 1
 else
  checkSessionLost = 0
 end if
end function


' Cart Weight
function calculateCartFreeShippingWeight(pIdDbSessionCart)   
 
 ' sum the total weight of the FS products 
 mySQL="SELECT SUM(quantity*unitWeight) AS total FROM cartRows, products WHERE idDbSessionCart="&pIdDbSessionCart& " AND cartRows.idProduct=products.idProduct AND freeShipping=-1"

 call getFromDatabase(mySQL, rsTemp, "cartFunctions")    

 if rstemp.eof then
  calculateCartFreeShippingWeight=0
 else
 
  if not isNull(rstemp("total")) then
    calculateCartFreeShippingWeight=Cdbl(rstemp("total"))
  else
    calculateCartFreeShippingWeight=0
  end if
 
 end if
   
 ' sum the weight of the optionals, if any
 
 mySQL="SELECT SUM(cartRows.quantity*options.weight) AS total FROM cartRows, cartRowsOptions, options, products WHERE cartRows.idCartRow=cartRowsOptions.idCartRow AND cartRowsOptions.idOption=options.idOption AND idDbSessionCart="&pIdDbSessionCart& " AND cartRows.idProduct=products.idProduct AND freeShipping=-1"

 call getFromDatabase(mySQL, rsTemp, "cartFunctions")    
 
 pTotalCartWeight=rstemp("total") 
 
 if not rstemp.eof and not isNull(pTotalCartWeight) then
    calculateCartFreeShippingWeight=calculateCartFreeShippingWeight+Cdbl(pTotalCartWeight)  
 end if  
   
end function

' Cart Free Shipping Quantity
function calculateCartFreeShippingQuantity(pIdDbSessionCart)
 
 dim mySQL, rstemp2  
     
 ' sum
 mySQL="SELECT SUM(quantity) AS qTotal FROM cartRows, products WHERE idDbSessionCart="&pIdDbSessionCart& " AND cartRows.idProduct=products.idProduct AND freeShipping=-1" 
 
 call getFromDatabase(mySQL, rsTemp2, "cartFunctions")        

 if rstemp2.eof then
  calculateCartFreeShippingQuantity=0
 else    
  if isNull(rstemp2("qTotal")) then
   calculateCartFreeShippingQuantity=0
  else
   calculateCartFreeShippingQuantity=Cdbl(rstemp2("qTotal"))
  end if
 end if     
end function

' Cart total
function calculateFreeShippingTotal(pIdDbSessionCart)       
 
 dim mySql, rsTempCartTotal, rsTempCartTotal2, pTotal, pOptionsTotal  
 
 pTotal=Cdbl(0)
 
 ' iterate through FS cart rows
 
 mySQL="SELECT idCartRow, quantity, unitPrice FROM cartRows, products WHERE idDbSessionCart="&pIdDbSessionCart& " AND cartRows.idProduct=products.idProduct AND freeShipping=-1"   
 
 call getFromDatabase(mySQL, rsTempCartTotal, "cartFunctions")  
  
 do while not rsTempCartTotal.eof
 
  pUnitPrice		= Cdbl(rsTempCartTotal("unitPrice"))  
  pQuantity		= rsTempCartTotal("quantity")
  pIdCartRow		= rsTempCartTotal("idCartRow")    
     
  ' get Optionals value for that row    
  
  pOptionsTotal=Cdbl(0)
  
  mySQL="SELECT SUM(cartRowsOptions.priceToAdd) AS optionsTotal FROM options, cartRowsOptions WHERE options.idOption=cartRowsOptions.idOption AND cartRowsOptions.idCartRow="&pIdCartRow     
  
  call getFromDatabase(mySQL, rsTempCartTotal2, "cartFunctions")  
    
  if rsTempCartTotal2.eof then
   pOptionsTotal=Cdbl(0)   
  else
   if isNull(rsTempCartTotal2("optionsTotal")) then
    pOptionsTotal=Cdbl(0)   
   else
    pOptionsTotal=Cdbl(rsTempCartTotal2("optionsTotal"))
   end if
  end if                    
  
  pTotal=pTotal+pQuantity*(pUnitPrice+pOptionsTotal)        
  
  rsTempCartTotal.movenext
 loop   
   
 calculateFreeShippingTotal=pTotal
 
end function

sub calculateCartVolume(pIdDbSessionCart, pMaxLength, pMaxHeight, pSumWidth)
 
 pInterWidth		= getSettingKey("pInterWidth")
 
 dim mySQL, rstemp2  
     
 ' get max length
 mySQL="SELECT MAX(length) AS maxLength FROM cartRows, products WHERE idDbSessionCart="&pIdDbSessionCart& " AND cartRows.idProduct=products.idProduct" 
 
 call getFromDatabase(mySQL, rsTemp2, "cartFunctions")        

 if rstemp2.eof then
  pMaxLength=0
 else    
  if isNull(rstemp2("maxLength")) then
   pMaxLength=0
  else
   pMaxLength=Cdbl(rstemp2("maxLength"))
  end if
 end if     
 
 ' get max height
 mySQL="SELECT MAX(height) AS maxHeight FROM cartRows, products WHERE idDbSessionCart="&pIdDbSessionCart& " AND cartRows.idProduct=products.idProduct" 
 
 call getFromDatabase(mySQL, rsTemp2, "cartFunctions")        

 if rstemp2.eof then
  pMaxHeight=0
 else    
  if isNull(rstemp2("maxHeight")) then
   pMaxHeight=0
  else
   pMaxHeight=Cdbl(rstemp2("maxHeight"))
  end if
 end if     
 
 ' get sum width
 mySQL="SELECT width, quantity FROM cartRows, products WHERE idDbSessionCart="&pIdDbSessionCart& " AND cartRows.idProduct=products.idProduct" 
 
 call getFromDatabase(mySQL, rsTemp2, "cartFunctions")        
 
 pSumWidth=0
    
 ' sum width and sum inter space from settings    
 do while not rstemp2.eof   
   pSumWidth=pSumWidth+(rstemp2("quantity")*rstemp2("width")) + (rstemp2("quantity")*Cdbl(pInterWidth))
   rstemp2.movenext
 loop
 
 'response.write "<br>Max Lenght:"&pMaxLength 
 'response.write "<br>Max Height:"&pMaxHeight
 'response.write "<br>Sum Width:"&pSumWidth
end sub

Function getIdCartRowForAnItem(pIdDbSessionCart, pIdProduct, pIdOptions)
  
  dim pIdCartRow, pIdCartRowCandidate, pOptionFailed, pHasPersonalization, f       
  
  pIdCartRow		=0
  pHasPersonalization	=0
  
  mySQL="SELECT hasPersonalization FROM products WHERE idProduct=" &pIdProduct 
  call getFromDatabase(mySQL, rsTempCF1, "cartFunctions")        
  
  if not rstemp.eof then
   pHasPersonalization=rsTempCF1("hasPersonalization")
  end if
  
  if pHasPersonalization=0 then
  
  if pIdOptions="" then
    
    ' check for no options        
  	
    ' get all cartRows with that idProduct
    mySQL = "SELECT DISTINCT idCartRow  FROM cartRows WHERE idProduct=" & pIdProduct & " AND idDbSessionCart=" & pIdDbSessionCart
	
    call getFromDatabase(mySQL, rsTempCF1, "cartFunctions")        
	
    do while not rsTempCF1.eof  
      
      pIdCartRowCandidate = rsTempCF1("idCartRow")
      
      ' now check if the candidate has no cartRowsOptions entry
      
      mySQL = "SELECT COUNT(*) AS cartRowsEntries FROM cartRowsOptions WHERE cartRowsOptions.idCartRow=" & pIdCartRowCandidate             
      
      call getFromDatabase(mySQL, rsTempCF2, "cartFunctions")        
      
      if not rsTempCF2.eof then
        if cInt(rsTempCF2("cartRowsEntries"))=0 then pIdCartRow=pIdCartRowCandidate
      end if            
      
    rsTempCF1.movenext
    loop
    
    getIdCartRowForAnItem=pIdCartRow          
    
   else
   
    ' check with options    
    
    arrayIdOptions=split(pIdOptions,",")        
          	
    ' get all cartRows with that idProduct
    mySQL = "SELECT DISTINCT idCartRow FROM cartRows WHERE idProduct=" & pIdProduct & " AND idDbSessionCart=" & pIdDbSessionCart
	
    call getFromDatabase(mySQL, rsTempCF1, "cartFunctions")        

    ' lines moved outside the loop	
    pOptionFailed  = 0
    pIdCartRow	   = 0
      
    do while not rsTempCF1.eof
      
      pIdCartRowCandidate = rsTempCF1("idCartRow")
            
      ' check for all options  
      f=0
      do while f<=ubound(arrayIdOptions)
             
       mySQL = "SELECT COUNT(idOption) AS countIdOption FROM cartRowsOptions WHERE idCartRow=" & pIdCartRowCandidate & " AND idOption=" & arrayIdOptions(f)              
              
       call getFromDatabase(mySQL, rsTempCF2, "cartFunctions")        
       
       ' cannot find with that option
       if Cint(rsTempCF2("countIdOption"))=0 then         
         pOptionFailed=-1
       end if       
       
       f=f+1
      loop
    
     if pOptionFailed=0 and pIdCartRow=0 then
       pIdCartRow=pIdCartRowCandidate
     end if
    
     rsTempCF1.movenext
    loop
    
    getIdCartRowForAnItem=pIdCartRow    
    
   end if ' get with options
   else
    getIdCartRowForAnItem=0
  end if ' personalization
    
End Function

' get qty for a product in the cart
Function getQtyForACartProduct(pIdDbSessionCart, pIdProduct)      
  
  dim rsTempQFC
    	
  ' check if the product is inside the cart
  mySQL = "SELECT SUM(quantity) AS sumQty FROM cartRows WHERE idProduct=" &pIdProduct& " AND idDbSessionCart=" &pIdDbSessionCart
	
  call getFromDatabase(mySQL, rsTempQFC, "cartFunctions")        
    
  getQtyForACartProduct = 0
  	
  if not rsTempQFC.eof then 
	if not isNull(rsTempQFC("sumQty")) then getQtyForACartProduct = rsTempQFC("sumQty")  
  End If    
  	
End Function

' get qty for a product in the cart
Function getQtyForACartRow(pIdDbSessionCart, pIdCartRow)      
  
  dim rsTempQFC
    	
  ' check if the product is inside the cart
  mySQL = "SELECT SUM(quantity) AS sumQty FROM cartRows WHERE idCartRow=" &pIdCartRow& " AND idDbSessionCart=" &pIdDbSessionCart
	
  call getFromDatabase(mySQL, rsTempQFC, "cartFunctions")        
    
  getQtyForACartRow = 0
  	
  if not rsTempQFC.eof then 
	if not isNull(rsTempQFC("sumQty")) then getQtyForACartRow = rsTempQFC("sumQty")  
  End If    
  	
End Function

function getCartRowOptionals(pIdCartRow)
 
 dim rstempGCRO, pHtmlToReturn
 
 pHtmlToReturn=""
 
 ' get optionals for current cart row
 mySQL="SELECT options.idOption as myIdOption, cartRowsOptions.optionDescrip, cartRowsOptions.priceToAdd FROM options, cartRowsOptions WHERE options.idOption=cartRowsOptions.idOption AND idCartRow="&pIdCartRow	  	  	  	  
	  
 call getFromDatabase(mySQL, rstempGCRO, "getCartRowOptionals") 	  	  
         
 do while not rstempGCRO.eof
         
          pIdOption		= rstempGCRO("myIdOption")
	  pOptionDescrip	= rstempGCRO("optionDescrip")
	  pPriceToAdd		= rstempGCRO("priceToAdd")
	  
	  pHtmlToReturn=pHtmlToReturn&pOptionDescrip
	  
	  if pPriceToAdd>0 then
	     pHtmlToReturn=pHtmlToReturn&" "&pCurrencySign & money(pPriceToAdd)
	  end if	  	  
	  
	  pHtmlToReturn=pHtmlToReturn&"<br>"	  	  
          
          rstempGCRO.movenext
 loop                   
 
 if pHtmlToReturn="" then
   pHtmlToReturn="-"
 end if	
        
 getCartRowOptionals=pHtmlToReturn
 
end function

function getCartRowPrice(pIdCartRow)
 
 dim rstempGCRP, pUnitPrice, pQuantity, pSumPriceToAdd
 
 getCartRowPrice=0
 
 ' get unit price * qty
 mySQL="SELECT unitPrice, quantity FROM cartRows WHERE idCartRow="&pIdCartRow	  	  	  	  
	  
 call getFromDatabase(mySQL, rstempGCRP, "getCartRowOptionals") 	  	  
 
 if not rstempGCRP.eof then 
  pUnitPrice	=rstempGCRP("unitPrice")
  pQuantity	=rstempGCRP("quantity")
 end if  
 
 ' get optionals price for the row
 mySQL="SELECT SUM(priceToAdd) AS sumPriceToAdd FROM cartRowsOptions WHERE idCartRow="&pIdCartRow	  	  		  	  
 call getFromDatabase(mySQL, rstempGCRP, "getCartRowOptionals") 	  	  
 
 pSumPriceToAdd=0
         
 if not isNull(rstempGCRP("sumPriceToAdd")) then
  pSumPriceToAdd=rstempGCRP("sumPriceToAdd")         
 end if                
 
 getCartRowPrice=pQuantity*(pUnitPrice+pSumPriceToAdd)
 
end function
%>
