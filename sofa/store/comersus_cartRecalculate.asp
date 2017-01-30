<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: recalculate quant for an item inside the cart
%>

<!--#include file="../includes/screenMessages.asp"--> 
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->
<!--#include file="../includes/cartFunctions.asp"--> 
<!--#include file="../includes/stringFunctions.asp"--> 

<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/customerTracking.asp"-->  

<%

on error resume next

dim mySQL, connTemp, rsTemp3, rstemp2, f, total, pNewQuantity, pIdProduct

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pForceSelectOptionals	= getSettingKey("pForceSelectOptionals")
pMaxAddCartQuantity	= getSettingKey("pMaxAddCartQuantity")
pCartQuantityLimit	= getSettingKey("pCartQuantityLimit")
pUnderStockBehavior	= lcase(getSettingKey("pUnderStockBehavior"))

total	     	 = 0

pIdDbSession	 = checkSessionData()
pIdDbSessionCart = checkDbSessionCartOpen()

pIdCustomerType = getSessionVariable("idCustomerType",1)
pIdCustomer  	= getSessionVariable("idCustomer",0)

if Cint(countCartRows(pIdDbSessionCart))=0 then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(40,"Empty cart")) 
end if

call customerTracking("comersus_cartRecalculate.asp", request.form)

' iterate through the cart in order to identify wich items have changed the quantity

mySQL="SELECT idCartRow, cartRows.idProduct, quantity FROM cartRows, products WHERE cartRows.idProduct=products.idProduct AND cartRows.idDbSessionCart="&pIdDbSessionCart

call getFromDatabase(mySQL, rsTemp3, "cartRecalculate") 

do while not rsTemp3.eof
    
    
    pNewQuantity=Cdbl(request.form(getUserInput(Cstr(rsTemp3("idCartRow")),8)))
    
    ' update items with different qty   
    if pNewQuantity<>Cdbl(rsTemp3("quantity")) then           
          
      pIdProduct	= rsTemp3("idProduct")
      pIdCartRow	= rsTemp3("idCartRow")       
      pStock		= getStock(pIdProduct)
      
      pPrice		= getPriceByQty(pIdProduct, pIdCustomerType, pIdCustomer, pNewQuantity)
                     
      ' replace , by .
      pPrice		= replace(pPrice,",",".")                  
            
      pOldRowQty	=getQtyForACartRow(pIdDbSessionCart, pIdCartRow)
      pCartQty		=getQtyForACartProduct(pIdDbSessionCart, pIdProduct)
      
      ' check stock 
      if (pStock-pCartQty-pNewQuantity+pOldRowQty)<0 and pUnderStockBehavior="dontadd" then       
         response.redirect "comersus_message.asp?message="&Server.UrlEncode(getMsg(38,"Stock restrictions"))      
      end if
                        
      ' update cart quantity with price recalculated due to disc per qty
      mySQL="UPDATE cartRows SET quantity=" &pNewQuantity& ", unitPrice=" &pPrice& " WHERE idCartRow="&pIdCartRow     
     
      call updateDatabase(mySQL, rstemp2, "cartRecalculate")          
      
    end if  ' that product has changed the quantity       
 rsTemp3.movenext
loop    

' clean session data
mySQL="UPDATE dbSession set sessionData='|' WHERE idDbSession=" &pIdDbSession
call updateDatabase(mySQL, rstemp, "orderVerify")

call closeDb()

response.redirect "comersus_goToShowCart.asp"
%>