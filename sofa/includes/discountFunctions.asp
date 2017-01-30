<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' USA - 2006
' http://www.comersus.com 
' Details: discount functions
%>
<%
sub getDiscount(pDiscountCode, pDiscountDesc, pDiscountTotal)

 dim pSubTotal, pCartTotalWeight, pCartQuantity, rstemp3, mySQL
 
 ' calculate total price of the order, total weight and product total quantities
 pSubTotal  		= Cdbl(calculateCartTotal(pIdDbSessionCart))
 pCartTotalWeight 	= Cdbl(calculateCartWeight(pIdDbSessionCart))
 pCartQuantity  	= Cdbl(calculateCartQuantity(pIdDbSessionCart))
 
 if pDiscountCode="" then

   ' user did not entered discount code 
   pDiscountDesc = getMsg(460,"not entered")
 
 else

  ' get discounts details 
  mySQL="SELECT idProduct, discountDesc, priceToDiscount, percentageToDiscount FROM discounts WHERE discountCode='" &pDiscountCode& "' AND active=-1 AND used=0 AND idStore=" &pIdStore& " AND " &pCartQuantity& ">=quantityFrom AND "  &pCartQuantity& "<=quantityUntil AND " &pCartTotalWeight&" >=weightFrom AND " &pCartTotalWeight& "<=weightUntil AND " &pSubTotal& ">=priceFrom AND " &pSubTotal& "<=priceUntil"

  call getFromDatabase(mySQL, rsTemp3, "orderVerify")   

  if rstemp3.eof then
   
    ' invalid discount code entered
    pDiscountDesc 	= getMsg(461,"invalid")
    pDiscountTotal 	= 0
    
  else          
    
   ' check if the discount is specific for one product      
   
   if rstemp3("idProduct")>0 then         
   
    ' find out if the product is in the cart
    
    if findProduct(pIdDbSessionCart, rstemp3("idProduct"))=0 then      
    
      pDiscountTotal				= 0
      pDiscountDesc 				= getMsg(462,"missing")
     else            
      
      ' get price and qty to calculate percentage discount      
      pProductPriceForPercentageDiscount	= findProductPrice(pIdDbSessionCart, pIdCustomerType, rstemp3("idProduct")) 
      pDiscountDesc				= rsTemp3("discountDesc")
      pDiscountTotal				= (rsTemp3("percentageToDiscount")*(pProductPriceForPercentageDiscount)/100) + rsTemp3("priceToDiscount")             
        
     end if ' findProduct
   
   else
     
     ' calculate amount for a generic discount
     pDiscountDesc	= rsTemp3("discountDesc")
     pDiscountTotal	= (rsTemp3("percentageToDiscount")*(pSubTotal)/100)  + rsTemp3("priceToDiscount")
      
   end if   ' is for specific product
      
 end if
 
 end if ' pDiscountCode=""
 
 ' if the discount is more than subtotal, limit
 if pDiscountTotal>pSubTotal then
  pDiscountTotal=pSubTotal
 end if
 
end sub

function markDiscountAsUsed(pDiscountCode)
 
 dim mysql, rstemp3, rstemp
 
 mySQL="SELECT idDiscount, oneTime FROM discounts WHERE discountcode='" &pDiscountCode& "' AND active=-1 AND used=0"
 call getFromDatabase(mySQL, rsTemp3, "markDiscountAsUsed")    

 if not rsTemp3.eof then
   
    pIdDiscount				= rsTemp3("idDiscount")
    pOneTime				= rsTemp3("oneTime")    
    
    ' update discount used to true
    if pOneTime=-1 then         
     mySQL="UPDATE discounts SET used=-1 WHERE idDiscount="& pIdDiscount
     call updateDatabase(mySQL, rsTemp, "markDiscountAsUsed")        
    end if        
   
 end if ' discount not found

end function

%>