<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: final screen with order # and confirmation
%>

<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/screenMessages.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/stringFunctions.asp"--> 
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/cartFunctions.asp"-->
<!--#include file="../includes/miscFunctions.asp"-->
<!--#include file="../includes/adSenseFunctions.asp"-->
<%
on error resume next 

dim connTemp, rsTemp, rsTemp3, pRowsShown, pMaxPopularity

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanySlogan	 	= getSettingKey("pCompanySlogan")
pOrderPrefix		= getSettingKey("pOrderPrefix")
pAuctions		= getSettingKey("pAuctions")
pListBestSellers 	= getSettingKey("pListBestSellers")
pNewsLetter 		= getSettingKey("pNewsLetter")
pPriceList 		= getSettingKey("pPriceList")
pStoreNews 		= getSettingKey("pStoreNews")
pOneStepCheckout	= getSettingKey("pOneStepCheckout")
pEmailAdmin		= getSettingKey("pEmailAdmin")
pExitSurvey		= getSettingKey("pExitSurvey")
pShowNews		= getSettingKey("pShowNews")
pGoogleAnalytics	= getSettingKey("pGoogleAnalytics")

pIdCustomer	 	= getSessionVariable("idCustomer",0)
pIdCustomerType	 	= getSessionVariable("idCustomerType",1)
pIdDbSessionCart	= getSessionVariable("idDbSessionCart",0)

pIdOrder		= getUserInput(request.querystring("idOrder"),20)
pRealIdOrder 		= removePrefix(pIdOrder,pOrderPrefix)

if pIdOrder="" then
 response.redirect "comersus_message.asp?message="&Server.Urlencode("Invalid Order")  
end if

%>
<!--#include file="header.asp"-->
<br><br> <%call displayAdConfirmation()%>
<br><br><b><%=getMsg(582,"order confirmation")%></b><br>
<br><img src="images/invoice.jpg" alt=Invoice>
<br><br><%=getMsg(583,"thanks, order")%>&nbsp;<a href="comersus_customerShowOrderDetails.asp?idOrder=<%=pIdOrder%>"><%=pIdOrder%></a>
<br><%=getMsg(584,"concerns")%>&nbsp; <a href="comersus_customerContactAdminForm.asp"><%=getMsg(585,"here")%></a> 

<%if pRealIdOrder>30 then%>
<br><br>(!) This store is powered by <a href="http://www.comersus.com">Comersus Shopping Cart Application</a>
<%end if%>

<%if pStoreFrontDemoMode="-1" then%>
 <br><br><%=getMsg(497,"thanks for testing")%> &nbsp;<%=getMsg(498,"visit")%>&nbsp; <a href="http://www.comersus.com/backOfficeTest/backOfficePlus"><%=getMsg(499,"admin demo")%></a>
<%end if%>

<%if pGoogleAnalytics<>"" and pGoogleAnalytics<>"0" then%>

<%
mySQL="SELECT state, stateCode, city, countryCode, total, taxAmount FROM orders WHERE orders.idOrder=" &pRealIdOrder&" AND idCustomer="&pIdCustomer

call getFromDatabase(mySQL, rstemp, "orderConfirmation") 

if  rstemp.eof then 
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(518,"Cannot locate the order inside your history"))
end if 

pState			= rstemp("state")
pStateCode		= rstemp("stateCode")
pCity			= rstemp("city")
pCountryCode		= rstemp("countryCode")
pTotal			= rstemp("total")
pTaxAmount		= rstemp("taxAmount")

pCountryName		= getCountryName(pCountryCode)
pStateName		= getStateName(pStateCode)
%>

<script type="text/javascript">
var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
</script>

<script type="text/javascript">
  var pageTracker = _gat._getTracker("<%=pGoogleAnalytics%>");
  pageTracker._initData();
  pageTracker._trackPageview();

  pageTracker._addTrans(
    "<%=pIdOrder%>",                        // Order ID
    "1234",                            		// Affiliation
    "<%=money(pTotal)%>",                       // Total
    "<%=money(pTaxAmount)%>",                   // Tax
    "<%=pShipmentAmount%>",                     // Shipping
    "<%=pCity%>",                               // City
    "<%=pStateName%>",                          // State
    "<%=pCountryName%>"                         // Country
  );
    <%
    
' get cartRow contents
    
mySQL="SELECT cartRows.idCartRow, products.idProduct, products.sku, products.description, cartRows.quantity, cartRows.unitPrice, emailText, personalizationDesc FROM cartRows, dbSessionCart, products WHERE dbSessionCart.idDbSessionCart=cartRows.idDbSessionCart AND cartRows.idProduct=products.idProduct AND dbSessionCart.idOrder=" &pRealIdOrder

call getFromDatabase(mySQL, rstemp2, "orderConfirmation") 

pSubTotal = 0
i	  = 0

do while not rstemp2.eof

	pIdCartRow		= rstemp2("idCartRow")
	pDescription		= rstemp2("description")				
	pSku			= rstemp2("sku")
	pIdProduct		= rstemp2("idProduct")
        ' get optionals                    
          
        pOptionGroupsTotal	= 0                    
          
        ' get optionals for current cart row
        mySQL="SELECT optionDescrip, priceToAdd FROM cartRowsOptions WHERE idCartRow="&pIdCartRow	  	  	  	  	  
	  
	call getFromDatabase(mySQL, rstemp3, "orderConfirmation") 
         
         do while not rstemp3.eof
         
	  pOptionDescrip	= rstemp3("optionDescrip")
	  pPriceToAdd		= rstemp3("priceToAdd")
	  
	  pDescription=pDescription&" " &pOptionDescrip&" "
	  
	  if pPriceToAdd>0 then
	   pDescription=pDescription& pCurrencySign&money(pPriceToAdd)
	  end if
	  	  
	  pOptionGroupsTotal	= pOptionGroupsTotal+ pPriceToAdd        
          
          rstemp3.movenext
         loop 
                                   
        pUnitPrice=0
        
	
	pUnitPrice= Cdbl(rstemp2("unitPrice")) + pOptionGroupsTotal		
        
        pCategoryDesc=""
        ' get category name
        mySQL="SELECT categoryDesc FROM categories, categories_products WHERE categories.idCategory=categories_products.idCategory AND idProduct="&pIdProduct	  	  	  	  	  	  
	call getFromDatabase(mySQL, rstemp4, "orderConfirmation")
        if not rstemp4.eof then
         pCategoryDesc=rstemp4("categoryDesc")
        end if
        
        pDescription=replace(pDescription,"'","")
    %>   
     pageTracker._addItem(
    "<%=pIdOrder%>",                                     
    "<%=pSku%>",                                     
    "<%=pDescription%>",                                  
    "<%=pCategoryDesc%>",                             
    "<%=money(pUnitPrice)%>",                           
    "<%=rstemp2("quantity")%>"                                    
  );

  pageTracker._trackTrans();
</script> 

<%  
 rstemp2.movenext
loop
%>  

<%end if%>

<br><br>                    
<!--#include file="footer.asp"-->
<%call closeDb()%>
