<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com 
' Details: add item to the shopping cart
%>

<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->
<!--#include file="../includes/screenMessages.asp"-->
<!--#include file="../includes/cartFunctions.asp"--> 
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/miscFunctions.asp"--> 
<!--#include file="../includes/customerTracking.asp"-->  

<% 

on error resume next

dim mySQL, connTemp, rsTemp, rsTemp2, pTotalQuantity, lineNumber, pProductHasOptionals, pIdCartRow, pRowPrice, f, pOptionDescrip, pPriceToAdd, pDefaultLanguage, pStoreFrontDemoMode, pCurrencySign, pDecimalSign, pMoneyDontRound, pCompany, pCompanyLogo, pHeaderKeywords, pAuctions, pListBestSellers, pNewsLetter, pPriceList, pStoreNews, pOneStepCheckout, pForceSelectOptionals, pMaxAddCartQuantity, pCartQuantityLimit, pUnderStockBehavior, pDateFormat, pCookiesDetection, pChangeDecimalPoint, pIdDbSession, pIdDbSessionCart, pIdCustomer, pIdCustomerType, pIdProduct, p1StepCheckout, pOptionsQuantity, pQuantity, pPersonalizationDesc, pFrom, pUntil, pIdProductFree, pStock, pHowManyOptionals, pHowManyOptionalsSelected, pThereAreSelectedOptionals, pUnitPrice, pUnitBToBPrice, pUnitWeight, pUnitCost, pSpecialPrice, pCartRowPrice, pNewQuantity, pOldQuantity

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pForceSelectOptionals	= getSettingKey("pForceSelectOptionals")
pMaxAddCartQuantity	= getSettingKey("pMaxAddCartQuantity")
pCartQuantityLimit	= getSettingKey("pCartQuantityLimit")
pUnderStockBehavior	= lcase(getSettingKey("pUnderStockBehavior"))
pDateFormat 		= getSettingKey("pDateFormat")
pCookiesDetection	= getSettingKey("pCookiesDetection")
pChangeDecimalPoint	= getSettingKey("pChangeDecimalPoint")
pTemplateStore		= getSettingKey("pTemplateStore")

if pCookiesDetection="-1" and readCookie()=0 then 
 response.redirect "comersus_cookiesInformation.asp"
end if

pTotalQuantity 			= Cdbl(0)
pHowManyOptionalsSelected 	= 0

pIdDbSession	 		= checkSessionData()
pIdDbSessionCart 		= checkDbSessionCartOpen()

pIdCustomer      		= getSessionVariable("idCustomer",0)
pIdCustomerType  		= getSessionVariable("idCustomerType",1)

if pIdDbSessionCart=0 then
 pIdDbSessionCart=createNewDbSessionCart() 
end if

' get data from viewitem form
pIdProduct	 	= getUserInput(request("idProduct"),10)
p1StepCheckout	 	= getUserInput(request("1StepCheckout"),3)
pQuantity	 	= getUserInput(request("quantity"),10)
pPersonalizationDesc 	= getUserInput(request("personalizationDesc"),150)

' get optionsGroups assigned

mySQL="SELECT idOptionGroup FROM optionsGroups_products WHERE idProduct=" & pIdProduct
call getFromDatabase (mySql, rsTempGEO1, "itemFunctions")

' iterate through optionsGroups, compile array with options

do while not rsTempGEO1.eof
	
	pIdCurrentOption	= trim(getUserInputL(request("idOptions"&rsTempGEO1("idOptionGroup")),80))
	
	if pIdCurrentOption<>"" then
	 pIdOptions		= pIdOptions & pIdCurrentOption & ","
	end if
	
	rsTempGEO1.MoveNext
loop

' remove last comma
if pIdOptions<>"" then
	pIdOptions		= left(pIdOptions, len(pIdOptions)-1)
end if

' options are sent comma delimited, create an array with each option
if pIdOptions<>"" then
   
   arrayIdOptions=split(pIdOptions,",")        
      
   ' check if there are optionals selected
   pHowManyOptionalsSelected=ubound(arrayIdOptions)+1   
   
end if  
	          
' check how many optionals are assigned to the product
mySQL="SELECT COUNT(idOptionGroup) AS howManyOptionals FROM optionsGroups_products WHERE idProduct=" &pIdProduct

call getFromDatabase(mySQL, rstemp, "addItem")

pHowManyOptionals = 0

if not rstemp.eof then
 pHowManyOptionals = rstemp("howManyOptionals")
end if

' if the product has optionals, is configured to reject and the user has not selected all optionals, then send to message
if pForceSelectOptionals="-1" and Cint(pHowManyOptionalsSelected)<Cint(pHowManyOptionals) then

 if request.querystring("idProduct")="" then
  ' came from viewItem
  response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(35,"Select optionals"))  
 else
 ' came from listing
   response.redirect "comersus_viewItem.asp?idProduct="&pIdProduct
 end if
 
end if

' redirect to donation screen
if isDonation(pIdProduct) and request.querystring("idProduct")<>"" then
 response.redirect "comersus_viewItemDonation.asp?idProduct="&pIdProduct
end if

' if cannot get quantity, reject
if pQuantity="" then
    response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(36,"Select qty"))
end if

' get price
pUnitPrice	= getPrice(pIdProduct, pIdCustomerType, pIdCustomer)

' get item details 
mySQL="SELECT weight, cost, listPrice FROM products WHERE idProduct=" &pIdProduct& " AND active=-1"

call getFromDatabase(mySQL, rstemp, "addItem")

if  rstemp.eof then 
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(37,"Cannot get details"))       
end if

pUnitWeight		= rstemp("weight")
pUnitCost		= rstemp("cost")
pStock			= getStock(pIdProduct)
   
pUnitPrice	= formatNumberForDb(pUnitPrice)
pUnitWeight	= formatNumberForDb(pUnitWeight)
pUnitCost	= formatNumberForDb(pUnitCost)


' check stock 
if (pStock-pQuantity-getQtyForACartProduct(pIdDbSessionCart, pIdProduct))<0 and pUnderStockBehavior="dontadd" then
    response.redirect "comersus_message.asp?message="&Server.UrlEncode(getMsg(38,"Stock restrictions"))
end if

' check backorder
if (pStock-pQuantity-getQtyForACartProduct(pIdDbSessionCart, pIdProduct))<0 and pUnderStockBehavior="backorder" then
    response.redirect "comersus_viewItem.asp?idProduct="&pIdProduct
end if

call customerTracking("comersus_addItem.asp", left(request.form,50))
  
' check if the item is already in the basket
pIdCartRow=getIdCartRowForAnItem(pIdDbSessionCart, pIdProduct, pIdOptions)    

' get old quantity
mySQL = "SELECT quantity FROM cartRows WHERE idCartRow="& pIdCartRow   
call getFromDatabase(mySQL, rstemp, "addItem")

if not rstemp.eof then
 pOldQuantity=rstemp("quantity")
else
 pOldQuantity=0
end if


if pIdCartRow=0 then
 
 ' insert new row line, is not in the cart   

 mySQL="INSERT INTO cartRows (idProduct, quantity, unitPrice, unitCost, unitWeight, idDbSessionCart, personalizationDesc) VALUES (" &pIdProduct& "," &pQuantity& "," &pUnitPrice& "," &pUnitCost& "," &pUnitWeight& "," &pIdDbSessionCart& ",'" &pPersonalizationDesc& "')"
 
 call updateDatabase(mySQL, rstemp, "addItem")

 ' retrieve ID
 
 mySQL="SELECT MAX(idCartRow) AS maxIdCartRow FROM cartRows WHERE idDbSessionCart=" & pIdDbSessionCart  
 call getFromDatabase(mySQL, rstemp, "addItem")
 
 if not rstemp.eof then
  pIdCartRow=rstemp("maxIdCartRow") 
 end if
     
 ' insert optionals 
 
 for f=0 to pHowManyOptionalsSelected-1

  if arrayIdOptions(f)<>"" then   
  
  ' get the price and optionDescrip (for historic purposes is also saved)
   mySQL="SELECT priceToAdd, percentageToAdd, optionDescrip FROM options WHERE idOption="&arrayIdOptions(f)
   call getFromDatabase(mySQL, rstemp, "addItem")      
   
   if not rstemp.eof then
        pOptionDescrip	=rsTemp("optionDescrip")
        if cSng(rsTemp("priceToAdd"))> 0 then
            pPriceToAdd		=rsTemp("priceToAdd")
        else
            pPriceToAdd		=pUnitPrice * rsTemp("percentageToAdd")/100
        end if        
   else
        pOptionDescrip	="-"
        pPriceToAdd	=0
   end if
   
   pPriceToAdd=replace(pPriceToAdd,",",".")
   
   mySQL="INSERT INTO cartRowsOptions (idCartRow, idOption, optionDescrip, priceToAdd) VALUES (" &pIdCartRow& "," &arrayIdOptions(f)& ",'" &pOptionDescrip& "',"  &pPriceToAdd& ")"          
   call updateDatabase(mySQL, rstemp, "addItem")      
   
  end if ' <>""   
 next
 
 pNewQuantity = Cdbl(pQuantity)
 
else

   ' item is already in the cart
   pNewQuantity = pOldQuantity + Cdbl(pQuantity)
    
   ' reset unit price before discounts     
   
   pUnitPrice 		= replace(pUnitPrice,",",".")   

   mySQL="UPDATE cartRows SET quantity=" &pNewQuantity& ", unitPrice=" &pUnitPrice& " WHERE idCartRow=" &pIdCartRow
   call updateDatabase(mySQL, rstemp, "addItem")            

end if

 ' load pRowPrice 
 
 pRowPrice=pUnitPrice
  
' get discount per quantity for old or just inserted line

mySQL="SELECT discountPerUnit FROM discountsPerQuantity WHERE idProduct=" &pIdProduct& " AND quantityFrom<=" &pNewQuantity& " AND quantityUntil>=" &pNewQuantity

call getFromDatabase(mySQL, rstemp, "addItem")

if not rstemp.eof then
 
 ' there are quantity discounts defined for that quantity 
 pDiscountPerUnit 	= rstemp("discountPerUnit")   
 pRowPrice 		= pRowPrice - pDiscountPerUnit  
 
 ' format for SQL
 pRowPrice 		= replace(pRowPrice,",",".") 

 ' update price
 
 mySQL="UPDATE cartRows SET unitPrice=" &pRowPrice& " WHERE idCartRow=" &pIdCartRow
 
 call updateDatabase(mySQL, rstemp, "addItem")            
  
end if		 
 
 ' clean session data
mySQL="UPDATE dbSession set sessionData='|' WHERE idDbSession=" &pIdDbSession
call updateDatabase(mySQL, rstemp, "orderVerify")

call closeDB()

response.redirect "comersus_goToShowCart.asp"
%>


