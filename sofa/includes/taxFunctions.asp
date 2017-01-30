<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' USA - 2006
' http://www.comersus.com 
' Details: Taxes functions
%>

<%

function calculateTax(pShipmentCharge, pSubTotal, pVatNumber)

 dim pIdDbSession, pIdDbSessionCart, rstemp, mysql

 pIdDbSession	 	= checkSessionData()
 pIdDbSessionCart 	= checkDbSessionCartOpen()

 pIdCustomer		= getSessionVariable("idCustomer",0)
 pIdCustomerType 	= getSessionVariable("idCustomerType",1)
 
 calculateTax		=0
 
 ' retrieve billing location

 mySQL="SELECT stateCode, countryCode, zip, shippingStateCode, shippingCountryCode, shippingZip FROM customers WHERE idCustomer=" &pIdCustomer

 call getFromDatabase(mySQL, rstemp, "calculateTax") 

if not rstemp.eof then
 
 pTaxStateCode		= rstemp("stateCode")
 pTaxCountryCode		= rstemp("countryCode")
 pTaxZip			= rstemp("zip")
 
 if rstemp("shippingStateCode")<>"" then
  ' use shipping for tax calculation
  pTaxStateCode		= rstemp("shippingStateCode")
  pTaxCountryCode		= rstemp("shippingCountryCode")
  pTaxZip			= rstemp("shippingZip")
 end if
 
end if

' if customer use anotherState, insert a dummy state code to simplify SQL sentence
if pTaxStateCode="" then
   pTaxStateCode="**"
end if

 ' tax per location segment (and the same customer type)

 mySQL="SELECT taxPerPlace FROM taxPerPlace WHERE ((stateCode='" &pTaxStateCode& "' AND stateCodeEq=-1) OR (stateCode IS NULL) OR (stateCode<>'" &pTaxStateCode& "' AND stateCodeEq=0)) AND ((countryCode='"&pTaxCountryCode&"' AND countryCodeEq=-1) OR (countryCode IS NULL) OR (countryCode<>'" &pTaxCountryCode& "' AND countryCodeEq=0)) AND ((zip='" &pTaxZip& "' AND zipEq=-1) OR (zip IS NULL) OR (zip<>'" &pTaxZip& "' AND zipEq=0)) AND (idCustomerType IS NULL OR idCustomerType=" &pIdCustomerType& ") AND idStore=" &pIdStore

 call getFromDatabase(mySQL, rstemp, "calculateTax") 

 if  rstemp.eof then 
  ' there are no taxes defined for that zone
  pTaxPerPlace=0
 end if

 do until rstemp.eof 
    pTaxPerPlace = pTaxPerPlace+rstemp("taxPerPlace")      
    rstemp.movenext
 loop

' tax per product segment

' iterate through products in the cart

mySQL="SELECT idCartRow, cartRows.idProduct, quantity, unitPrice FROM cartRows, products WHERE cartRows.idProduct=products.idProduct AND cartRows.idDbSessionCart="&pIdDbSessionCart

call getFromDatabase(mySQL, rsTemp2, "calculateTax") 	

if rsTemp2.eof then 
   response.redirect "comersus_supportError.asp?error="&Server.Urlencode("There are no products in the cart at calculateTax.  for idDbSessionCart="&pIdDbSessionCart) 
end if

do while not rstemp2.eof
	  
 pIdCartRow		= rsTemp2("idCartRow")
 pIdProduct		= rsTemp2("idProduct")
 pOptionGroupsTotal	= 0
	          
 ' get optionals for current cart row and sum the unit option price
 
 mySQL="SELECT priceToAdd FROM cartRowsOptions WHERE idCartRow="&pIdCartRow	  	  	  	  
		  
 call getFromDatabase(mySQL, rsTemp3, "calculateTax") 
	         
 do while not rstemp3.eof
	         	  
   pPriceToAdd		= rstemp3("priceToAdd")	  	  	  	  
   pOptionGroupsTotal	= pOptionGroupsTotal+ pPriceToAdd        
	          
   rstemp3.movenext
 loop                   
 
 mySQL="SELECT taxPerProduct FROM taxPerProduct WHERE ((stateCode='" &pTaxStateCode& "' AND stateCodeEq=-1) OR (stateCode IS NULL) OR (stateCode<>'" &pTaxStateCode& "' AND stateCodeEq=0)) AND ((countryCode='"&pTaxCountryCode&"' AND countryCodeEq=-1) OR (countryCode IS NULL) OR (countryCode<>'" &pTaxCountryCode& "' AND countryCodeEq=0)) AND ((zip='" &pTaxZip& "' AND zipEq=-1) OR (zip IS NULL) OR (zip<>'" &pTaxZip& "' AND zipEq=0)) AND idProduct=" &pIdProduct

 call getFromDatabase(mySQL, rstemp, "calculateTax")  
 
 do until rstemp.eof 
   ' example (21/100 * row price)   
   pTaxPerProductAmount=pTaxPerProductAmount+( rstemp("taxPerProduct")/100 * ( rstemp2("quantity") * (rstemp2("unitPrice")+pOptionGroupsTotal) ))          
   rstemp.movenext
loop

 rstemp2.movenext
loop

' tax calculations. Exclude Shipping charge and Payment surcharge 
if pSubTotal>0 and pVatNumber="" then 
 
 ' calculate tax considering shipping
 if pShippingTaxable="0" then
  pTaxPerPlaceAmount 	= (pSubTotal  + pTaxPerProductAmount)* pTaxPerPlace/100
 else
  pTaxPerPlaceAmount 	= (pSubTotal  + pShipmentCharge + pTaxPerProductAmount)* pTaxPerPlace/100
 end if
 
 calculateTax         	= pTaxPerProductAmount + pTaxPerPlaceAmount
 
end if

end function

function getTaxIncludedPercentage()

 dim rstemp, mysql
 
 ' get tax rate of the first country, it should be the same for all

 mySQL="SELECT taxPerPlace FROM taxPerPlace WHERE idStore="&pIdStore
 call getFromDatabase(mySQL, rstemp, "calculateTax") 
 
 if not rstemp.eof then
  getTaxIncludedPercentage=rstemp("taxPerPlace")
 else
  getTaxIncludedPercentage=0
 end if
 
end function
%>