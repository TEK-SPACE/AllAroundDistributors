<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: view item details for bundle products
%>

<!--#include file="../includes/screenMessages.asp"-->
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/stringFunctions.asp"--> 
<!--#include file="../includes/adSenseFunctions.asp"-->

<% 
on error resume next

dim connTemp, rsTemp, mySql, pIdProduct, pDescription, pPrice, pDetails, pListPrice, pImageUrl, pWeight, pIdProduct2

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanySlogan 		= getSettingKey("pCompanySlogan")
pHeaderKeywords 	= getSettingKey("pHeaderKeywords")

pMoneyDontRound	 	= getSettingKey("pMoneyDontRound")
pCurrencyConversion	= lCase(getSettingKey("pCurrencyConversion"))
pShowBtoBPrice		= getSettingKey("pShowBtoBPrice")
pShowStockView		= getSettingKey("pShowStockView")
pProductReviews		= getSettingKey("pProductReviews")
pUnderStockBehavior	= lCase(getSettingKey("pUnderStockBehavior"))
pEmailToFriend		= getSettingKey("pEmailToFriend")
pWishList		= getSettingKey("pWishList")
pRelatedProducts 	= getSettingKey("pRelatedProducts")
pGetRelatedProductsLimit= getSettingKey("pGetRelatedProductsLimit")
pAffiliatesStoreFront   = getSettingKey("pAffiliatesStoreFront")
pRealTimeShipping	= lCase(getSettingKey("pRealTimeShipping"))
pCompareWithAmazon	= getSettingKey("pCompareWithAmazon")
pAuctions		= getSettingKey("pAuctions")
pListBestSellers 	= getSettingKey("pListBestSellers")
pNewsLetter 		= getSettingKey("pNewsLetter")
pPriceList 		= getSettingKey("pPriceList")
pStoreNews 		= getSettingKey("pStoreNews")
pExitSurvey		= getSettingKey("pExitSurvey")
pTemplateStore		= getSettingKey("pTemplateStore")
pDonationStringId	= getSettingKey("pDonationStringId")
pAdSenseClient 		= getSettingKey("pAdSenseClient")
pShowNews 		= getSettingKey("pShowNews")
pProductCustomField1	= getSettingKey("pProductCustomField1")

' session
pIdCustomer     	= getSessionVariable("idCustomer",0)
pIdCustomerType 	= getSessionVariable("idCustomerType",1)

pIdProduct 		= getUserInput(request.queryString("idProduct"),10)
pIdAffiliate		= getUserInput(request.querystring("idAffiliate"),4)

' set affiliate
if isNumeric(pIdAffiliate)then
   session("idAffiliate")= pIdAffiliate
end if

if trim(pIdProduct)="" or IsNumeric(pIdProduct)=false then
   response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(17,"Invalid")) 
end if

' redirect to rental screen
if isRental(pIdProduct) then 
 response.redirect "comersus_optRentalEnterInterval.asp?idProduct="&pIdProduct
end if

' if is not a bundle main, redirect to regular viewItem
if not isBundleMain(pIdProduct) then 
 response.redirect "comersus_viewItem.asp?idProduct="&pIdProduct
end if

' gets item details from db

mySQL="SELECT description, details, sku, searchKeywords, imageUrl, user1 FROM products WHERE idProduct=" &pIdProduct& " AND active=-1 AND idStore=" &pIdStore  
 
call getFromDatabase (mySql, rsTemp, "ViewItem")

if  rsTemp.eof then 
  response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(580,"Cannot get product details. Please contact us to request more information about item "&pIdProduct))  
end if

pDescription		= rsTemp("description")
pDetails		= rsTemp("details")
pStock			= getStock(pIdProduct)
pSku			= rsTemp("sku")
pSearchKeywords   	= rsTemp("searchKeywords")
pImageUrl		= rsTemp("imageUrl")
pUser1			= rsTemp("user1")

' for meta tags
pTitle			= pDescription 

%>
<!--#include file="header.asp"--> 

<br><br>    
<table width="460" border="0" align="left">
<tr><td colspan=4><b><%=pSku& " " &pDescription%></b></td></tr>
<tr> <td colspan=4><%=pDetails%></td></tr>                             
<tr><td colspan=4>

<%if pImageUrl<>"" then%>
  <img src='catalog/<%=pImageUrl%>'>
<%end if%>

<%if pProductReviews="-1" then
                            
   pRateReview=getRateReview(pIdProduct)
		
       if Cdbl(pRateReview)>0 then%>
         
           <br><a href="comersus_optReviewReadExec.asp?idProduct=<%=pIdProduct%>&description=<%= Server.UrlEncode(pDescription)%>"><%=getMsg(23,"Rated")%>&nbsp;<%=round(pRateReview,2)%></a>&nbsp;
           <img src="images/<%=Cint(pRateReview)%>stars.gif">
    
        <%else%>                   
           <br><a href="comersus_optReviewAddForm.asp?idProduct=<%=pIdProduct%>&description=<%= Server.UrlEncode(pDescription)%>"><%=getMsg(24,"Not rated")%></a><br><br>
        <%end if%>  
         
      <%end if ' reviews%>
                            
 </td>
</tr>
        
<tr> 
  <td colspan=4>            
                                                                        
    <%if pDiscountPerQuantity=-1 then%>
        <br><%=getMsg(27,"Discounts qty")%><br>
    <%end if%>
                
    <%if pEmailToFriend="-1" then%>
      <br><a href="comersus_optEmailToFriendForm.asp?idProduct=<%=pIdProduct%>&description=<%=Server.UrlEncode(pDescription) & " " &pCurrenCySign& money(pPrice)%>"><%=getMsg(30,"Email 2 friend")%></a>
    <%end if%>            

    <%if pWishList="-1" then%>
      <br><a href="comersus_customerWishListAdd.asp?idProduct=<%=pIdProduct%>&redirect=wish"><%=getMsg(31,"Add 2 WL")%></a>              
    <%end if%>                                                
    
    <%if pProductCustomField1="YouTubeV" and pUser1<>"" then%>     
     <br><object><param name="movie" value="http://www.youtube.com/v/<%=pUser1%>"></param><param name="wmode" value="transparent"></param><embed src="http://www.youtube.com/v/<%=pUser1%>" type="application/x-shockwave-flash" wmode="transparent"></embed></object>
    <%end if%>
    
  </td>          
</tr>

<tr><td colspan=4>&nbsp;</td></tr>        

<tr>
 <td><b><%=getMsg(665,"img")%></b></td>
 <td><b><%=getMsg(666,"desc")%></b></td>
 
 <%if pShowStockView="-1" then%>
  <td><b><%=getMsg(667,"stock")%></b></td>
 <%end if%>
 
 <td><b><%=getMsg(668,"price")%></b></td>
 <td colspan=2><b><%=getMsg(669,"purch")%></b></td>
</tr>
<form method="post"action="comersus_addItemBundle.asp" name="additem">                          
<%
' get inner products

mySQL="SELECT products.idProduct, smallImageUrl, description, formQuantity, map FROM bundles, products WHERE mainIdProduct=" &pIdProduct& " AND bundles.idProduct=products.idProduct"

call getFromDatabase (mySql, rsTemp, "ViewItem")                             
pCounter=1
      
do while not rsTemp.eof
 
 pPrice		  = getPrice(rstemp("idProduct"), pIdCustomerType, pIdCustomer)             
 pFormQuantity    = rsTemp("formQuantity")       
 pMapPrice    = rsTemp("map")       
 if pFormQuantity=0 then 
  pFormQuantity=1
 end if

%>  

 <input type="hidden" name="idProduct<%=pCounter%>" value="<%=rsTemp("idProduct")%>">
<tr>
 <td>
 <%if rstemp("smallImageUrl")<>"" then%>
  <img src='catalog/<%=rstemp("smallImageUrl")%>'>
 <%else%>
  <img src="catalog/imageNa_sm.gif">
 <%end if%>
 </td>
 <td><%=rstemp("description")%></td>
 
 <%if pShowStockView="-1" then%>
  <td><%=getStock(rstemp("idProduct"))%></td>
 <%end if%>
 
 <td><% If pMapPrice = "-1" Then
 		response.write "(Add to the cart to find out)"
 	else%>
 <%=pCurrencySign&money(pPrice)%>
 	<%end if%>
 </td>
 
 <td>
 <select name="quantity<%=pCounter%>">
  <option value="0" selected>0</option>
  <option value="1">1</option>
      <%if pFormQuantity>1 then
        for f=2 to pFormQuantity%>              
         <option value="<%=f%>"><%=f%></option>              
       <%next
      end if%>
</select>            
 </td>             
 
 </tr>
<%
		pCounter= pCounter+1
	rstemp.movenext
loop
%> 
<tr> <td colspan=5 align=right><input type=submit value="<%=getMsg(664,"add")%>" name=add></td>
</tr>
</form>
<tr><td colspan=4>&nbsp;</td></tr>        
</table>      

<!--#include file="footer.asp"--> 
<%call closeDb()%>
