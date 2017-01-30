<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: view item details for donation products
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

' if is not a donation, redirect to regular viewItem
if not isDonation(pIdProduct) then 
 response.redirect "comersus_viewItem.asp?idProduct="&pIdProduct
end if

' gets item details from db

mySQL="SELECT description, details, sku, searchKeywords, imageUrl FROM products WHERE idProduct=" &pIdProduct& " AND active=-1 AND idStore=" &pIdStore  
 
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

' for meta tags
pTitle			= pDescription 

%>
<!--#include file="header.asp"--> 
<br>  

<br><br>    
<table width="500" border="0" align="left">

<tr><td colspan=4>
<%if pImageUrl<>"" then%>
<img src='catalog/<%=pImageUrl%>'>              
<%else%>
<img src='catalog/imageNA.gif'>              
<%end if%>
</td>
</tr>

<tr><td colspan=4><b><%=pSku& " " &pDescription%></b></td></tr>
<tr> <td colspan=4><%=pDetails%></td></tr>                             
        
<tr> 
  <td colspan=4>            
                                                                                        
    <%if pEmailToFriend="-1" then%>
      <br><a href="comersus_optEmailToFriendForm.asp?idProduct=<%=pIdProduct%>&description=<%=Server.UrlEncode(pDescription) & " " &pCurrenCySign& money(pPrice)%>"><%=getMsg(30,"Email 2 friend")%></a>
    <%end if%>            
    
  </td>          
</tr>

<tr> 
  <td colspan=4>            

<form method="post"action="comersus_addItemDonation.asp" name="additem">                          
 <input type="hidden" name="idProduct" value="<%=pIdProduct%>"> 
 <%=pCurrencySign%>&nbsp;<input type=text name=price value=0 size=4> 
 <input type=submit value="<%=getMsg(670,"donate")%>" name=add> 
</form>
</td>
</tr>

<tr><td colspan=4>&nbsp;</td></tr>        
</table>      

<!--#include file="footer.asp"--> 
<%call closeDb()%>
