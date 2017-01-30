<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: view item details, increase visits
%>

<!--#include file="../includes/screenMessages.asp"-->
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/stringFunctions.asp"--> 
<!--#include file="../includes/cartFunctions.asp"--> 
<!--#include file="../includes/customerTracking.asp"-->  
<!--#include file="comersus_optCompareWithAm.asp"--> 
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
pShowSearchBox 		= getSettingKey("pShowSearchBox")
pUnderStockBehavior	= lCase(getSettingKey("pUnderStockBehavior"))

pAuctions		= getSettingKey("pAuctions")
pSuppliersList		= getSettingKey("pSuppliersList")
pNewsLetter 		= getSettingKey("pNewsLetter")
pAffiliatesStoreFront   = getSettingKey("pAffiliatesStoreFront")
pAllowNewCustomer	= getSettingKey("pAllowNewCustomer")
pRssFeedServer		= getSettingKey("pRssFeedServer")
pRealTimeShipping	= lCase(getSettingKey("pRealTimeShipping"))
pCompareWithAmazon	= getSettingKey("pCompareWithAmazon")
pAdSenseClient 		= getSettingKey("pAdSenseClient")
pShowNews		= getSettingKey("pShowNews")

pEmailToFriend		= getSettingKey("pEmailToFriend")
pWishList		= getSettingKey("pWishList")
pRelatedProducts 	= getSettingKey("pRelatedProducts")
pGetRelatedProductsLimit= getSettingKey("pGetRelatedProductsLimit")

pProductCustomField1 	= getSettingKey("pProductCustomField1")
pProductCustomField2 	= getSettingKey("pProductCustomField2")
pProductCustomField3 	= getSettingKey("pProductCustomField3")
pZoomItemImage		= getSettingKey("pZoomItemImage")

pShowSuggestionBox	= getSettingKey("pShowSuggestionBox")
pVerificationCodeEnabled= getSettingKey("pVerificationCodeEnabled")

' session
pIdCustomer     	= getSessionVariable("idCustomer",0)
pIdCustomerType 	= getSessionVariable("idCustomerType",1)
pIdDbSessionCart	= getSessionVariable("idDbSessionCart",0)

pIdProduct 		= getUserInput(request.queryString("idProduct"),10)
pImage			= getUserInput(request.queryString("image"),2)
pIdAffiliate		= getUserInput(request.querystring("idAffiliate"),4)

if pImage="" then 
 pImage=1
end if

' set affiliate
if isNumeric(pIdAffiliate)then
   session("idAffiliate")= pIdAffiliate
end if

if trim(pIdProduct)="" or IsNumeric(pIdProduct)=false then
   response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(17,"Invalid")) 
end if

call customerTracking("comersus_viewItem.asp", request.querystring)

' increase visits to the product
mySQL="UPDATE products SET visits=visits+1 WHERE idProduct="& pIdProduct

call updateDatabase(mySQL, rsTemp, "ViewItem")

' redirect to rental screen
if isRental(pIdProduct) then 
 response.redirect "comersus_optRentalEnterInterval.asp?idProduct="&pIdProduct
end if

' redirect to bundle screen
if isBundleMain(pIdProduct) then 
 response.redirect "comersus_viewItemBundle.asp?idProduct="&pIdProduct
end if

' redirect to donation screen
if isDonation(pIdProduct) then
 response.redirect "comersus_viewItemDonation.asp?idProduct="&pIdProduct
end if

' check for discount per quantity
pDiscountPerQuantity=itHasDiscountPerQty(pIdProduct)

' check for auctions
pIdAuction=getIdAuction(pIdProduct)

' get item price
pPrice		 = getPrice(pIdProduct, pIdCustomerType, pIdCustomer)

' gets item details from db

mySQL="SELECT description, details, listPrice, map, imageUrl, imageUrl2, imageUrl3, imageUrl4, sku, formQuantity, hasPersonalization, freeShipping, isDonation, searchKeywords, user1, user2, user3, emailText FROM products WHERE idProduct=" &pIdProduct& " AND active=-1 AND idStore=" &pIdStore  
 
call getFromDatabase (mySql, rsTemp, "ViewItem")

if  rsTemp.eof then 
  response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(580,"Cannot get product details. Please contact us to request more information about item "&pIdProduct))  
end if

pDescription		= rsTemp("description")
pDetails		= rsTemp("details")
pListPrice		= rsTemp("listPrice")
pMapPrice		= rsTemp("map")
pImageUrl		= rsTemp("imageUrl")
pImageUrl2		= rsTemp("imageUrl2")
pImageUrl3		= rsTemp("imageUrl3")
pImageUrl4		= rsTemp("imageUrl4")
pStock			= getStock(pIdProduct)
pSku			= rsTemp("sku")
pFormQuantity		= rsTemp("formQuantity")
pHasPersonalization 	= rsTemp("hasPersonalization")
pFreeShipping 		= rsTemp("freeShipping")
pIsDonation 		= rsTemp("isDonation")
pSearchKeywords   	= rsTemp("searchKeywords")
pUser1			= rsTemp("user1")
pUser2			= rsTemp("user2")
pUser3			= rsTemp("user3")
pEmailText		= rsTemp("emailText")

' for meta tags
pTitle			= pDescription & " " &pCurrencySign & money(pPrice)
pMetaDescription	= pDescription & " " &pCurrencySign & money(pPrice)

%>
<!--#include file="header.asp"--> 
<script language="JavaScript1.2">
function findDOM(objectId) {
if (document.getElementById) {
return (document.getElementById(objectId));}
if (document.all) {
return (document.all[objectId]);}
}
function zoom(type,imgx,sz) {
imgd = findDOM(imgx);
if (type=="+" && imgd.width < 600) {
imgd.width += 2;imgd.height += (2*sz);}
if (type=="-" && imgd.width > 20) {
imgd.width -= 2;imgd.height -= (2*sz);}
} 
</script>
<br>  

<br><br>    

<form method="post"action="comersus_addItem.asp" name="additem">
 <input type="hidden" name="idProduct" value="<%=pIdProduct%>">
 
 <!-- google_ad_section_start -->
 <div class="description"><%=pSku& " " &pDescription%></div>        
 <br><br>
 
 <%call getImage(pImageUrl, pImageUrl2, pImageUrl3, pImageUrl4, pImage)%>
 
 <br><br><%=pDetails%>
 
 <%if instr(pEmailText,".zip")<>0 then%>
  <img src="images/download.gif" alt="<%=getMsg(753,"download")%>"> &nbsp;<%=getMsg(753,"download")%>
 <%end if%>
 
 <br><br><div class=price><%=getMsg(15,"Price")%> &nbsp;
 
 <%If pMapPrice = "-1" Then%>
 	[<%=getMsg(745,"add to cart to find out")%>] 	
 <%else%>
 	<%=pCurrencySign & money(pPrice)%></div>
                                          
      <%if pCurrencyConversion="static" then%>
       <a href="comersus_optCurrencyConversion.asp?amount=<%=money(pPrice)%>&returnUrl=comersus_viewItem.asp?idProduct=<%=pIdProduct%>"><%=getMsg(18,"Other cur")%></a>
      <%end if%>
      <%if pCurrencyConversion="dynamic" then%>
       <a href="comersus_optCurrencyConversionOnLine.asp?amount=<%=money(pPrice)%>" target="_blank"><%=getMsg(18,"Other cur")%></a>              
      <%end if%>                      

     <%if trim(pShowBtoBPrice)="-1" and pIdCustomerType=1 and getPrice(pIdProduct, 2, pIdCustomer)>0 then%> 
        <br><%=getMsg(151,"Wholesale")%> <%=getMsg(15,"Price")%>:&nbsp <%=pCurrencySign & money(getPrice(pIdProduct, 2, pIdCustomer))%>        
     <%end if%>    
      
      <%if (pListPrice-pPrice)>0 then%> 
	<br><%=getMsg(19,"List price")%>&nbsp;<%=pCurrencySign & money(pListPrice)%> 
	<%=getMsg(16,"Saving")%>&nbsp;<%=pCurrencySign & money(pListPrice-pPrice)%> 
      <%end if%> 
  <%end if%> 
    
      <%if pShowStockView="-1" then%>
        <br><%=getMsg(20,"Stock")%>&nbsp;: <%=pStock%>                   
      <%end if%>  
      
    <%if pProductCustomField1="YouTubeV" and pUser1<>"" then%>     
     <br><object><param name="movie" value="http://www.youtube.com/v/<%=pUser1%>"></param><param name="wmode" value="transparent"></param><embed src="http://www.youtube.com/v/<%=pUser1%>" type="application/x-shockwave-flash" wmode="transparent"></embed></object>
    <%end if%>
    
    <%if lcase(pProductCustomField1)="mp3sample" and pUser1<>"" then%>     
     
     <script type="text/javascript" src="mp3player/swfobject.js"></script>
	<div id="flashcontent">
		To listen the mp3 sample preview you will need to have Javascript turned on and have <a href="http://www.adobe.com/go/getflashplayer/" target="_blank">Flash Player 9</a> or better installed.
	</div>

	<script type="text/javascript">
		// <![CDATA[

		var so = new SWFObject("mp3player/ep_player.swf", "ep_player", "301", "16", "9", "#FFFFFF");
		so.addVariable("skin", "mp3player/skins/micro_player/skin.xml");
		so.addVariable("file", "<location>mp3player/<%=pUser1%></location><creator></creator><title><%=pDescription%></title>");
		so.addVariable("autoplay", "false");
		so.addVariable("shuffle", "false");
		so.addVariable("repeat", "false");
		so.addVariable("buffertime", "1");
		so.write("flashcontent");

		// ]]>
	</script>
	
    <%end if%>
    
    <%if lcase(pProductCustomField2)="mp3sample" and pUser2<>"" then%>     
     
     <script type="text/javascript" src="mp3player/swfobject.js"></script>
	<div id="flashcontent">
		To listen the mp3 sample preview you will need to have Javascript turned on and have <a href="http://www.adobe.com/go/getflashplayer/" target="_blank">Flash Player 9</a> or better installed.
	</div>

	<script type="text/javascript">
		// <![CDATA[

		var so = new SWFObject("mp3player/ep_player.swf", "ep_player", "301", "16", "9", "#FFFFFF");
		so.addVariable("skin", "mp3player/skins/micro_player/skin.xml");
		so.addVariable("file", "<location>mp3player/<%=pUser2%></location><creator></creator><title><%=pDescription%></title>");
		so.addVariable("autoplay", "false");
		so.addVariable("shuffle", "false");
		so.addVariable("repeat", "false");
		so.addVariable("buffertime", "1");
		so.write("flashcontent");

		// ]]>
	</script>
	
    <%end if%>
    
    <%if lcase(pProductCustomField3)="mp3sample" and pUser3<>"" then%>     
     
     <script type="text/javascript" src="mp3player/swfobject.js"></script>
	<div id="flashcontent">
		To listen the mp3 sample preview you will need to have Javascript turned on and have <a href="http://www.adobe.com/go/getflashplayer/" target="_blank">Flash Player 9</a> or better installed.
	</div>

	<script type="text/javascript">
		// <![CDATA[

		var so = new SWFObject("mp3player/ep_player.swf", "ep_player", "301", "16", "9", "#FFFFFF");
		so.addVariable("skin", "mp3player/skins/micro_player/skin.xml");
		so.addVariable("file", "<location>mp3player/<%=pUser3%></location><creator></creator><title><%=pDescription%></title>");
		so.addVariable("autoplay", "false");
		so.addVariable("shuffle", "false");
		so.addVariable("repeat", "false");
		so.addVariable("buffertime", "1");
		so.write("flashcontent");

		// ]]>
	</script>
	
    <%end if%>    
    
      <%               
      if pCompareWithAmazon="-1" then
       pAmazonPrice=getAmazonPrice(pSku)
       if pAmazonPrice>pPrice then%>
        <br><%=getMsg(21,"Amazon price")%>&nbsp; $ <%=(pAmazonPrice)%> 
        <br><a href="http://www.amazon.com/exec/obidos/ASIN/<%=pSku%>" target="_blank"><%=getMsg(22,"See yself")%></a><br>
       <%end if ' price is better
       
     end if%>            
     
     <%if pProductCustomField1<>"" and pUser1<>"" and pProductCustomField1<>"YouTubeV" and lcase(pProductCustomField1)<>"mp3sample" then%>  
      <br><%=pProductCustomField1%>: <a href="comersus_listItems.asp?user1=<%=pUser1%>"><%=pUser1%></a>             
     <%end if%>
     
     <%if pProductCustomField2<>"" and pUser2<>"" and lcase(pProductCustomField2)<>"mp3sample" then%>  
      <br><%=pProductCustomField2%>: <a href="comersus_listItems.asp?user2=<%=pUser2%>"><%=pUser2%></a>             
     <%end if%>
     
     <%if pProductCustomField3<>"" and pUser3<>"" and lcase(pProductCustomField3)<>"mp3sample" then%>  
      <br><%=pProductCustomField3%>: <a href="comersus_listItems.asp?user3=<%=pUser3%>"><%=pUser3%></a>             
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
      
      <%if pHasPersonalization=-1 then%>
        <br><%=getMsg(25,"Person")%> <input type="text" name="personalizationDesc" size="20" maxlength="150">
      <%end if%>

      <!-- google_ad_section_end -->
      
<br><br>
		<%=getOptionsGroups(pIdProduct)%>                                                             
		
		<%if pUnderStockBehavior="backorder" and Cdbl(pStock)<=0 then%>                        
		
		 <%if pIdCustomer>0 then%>
		  <a href="comersus_optAddItemBackorder.asp?idProduct=<%=pIdProduct%>"><img src="images/buttonBackOrder.gif" border=0 alt="<%=getMsg(746,"2 backorder")%> - <%=getMsg(747,"we will notify...")%>"></a>
		 <%else%>
		  <br><a href="comersus_customerAuthenticateForm.asp"><%=getMsg(748,"login 2 backorder")%></a>.
		 <%end if%>
		<%else%>       
		
		 <br><select name="quantity">
		  <option value="1" selected>1</option>
		  <%if pFormQuantity>1 then
		  for optF=2 to pFormQuantity%>              
		   <option value="<%=optF%>"><%=optF%></option>              
		  <%next
		  end if%>
		 </select>   
		                 
		 <input type="submit" name="add" value="<%=getMsg(664,"Add 2 cart")%>">
		<%end if%>      
		 
		<br>                                    
		
		<%if pDiscountPerQuantity=-1 then%>
		 <br><%=getMsg(27,"Discounts qty")%><br>
		<%end if%>
		
		<%if pIdAuction>0 then%>
		 <br><%=getMsg(28,"Get 4 less")%>
		 <a href="comersus_optAuctionOfferForm.asp?idAuction=<%=pIdAuction%>&redirectUrl=comersus_optAuctionOfferForm.asp?idAuction=<%=pIdAuction%>"><%=getMsg(29,"Auction")%></a><br>
		<%end if%>                                                  
		
		<%if pEmailToFriend="-1" then%>
		 <br><a href="comersus_optEmailToFriendForm.asp?idProduct=<%=pIdProduct%>&description=<%=Server.UrlEncode(pDescription) & " " &pCurrenCySign& money(pPrice)%>">
		 <img src="images/buttonEmailToFriend.gif" border=0 alt="<%=getMsg(30,"Email 2 friend")%>"></a>
		<%end if%>            
		
		<%if pWishList="-1" then%>
		 <a href="comersus_customerWishListAdd.asp?idProduct=<%=pIdProduct%>&redirect=wish">
		 <img src="images/buttonWishList.gif" border=0 alt=" <%=getMsg(31,"Add 2 WL")%>"></a>              
		<%end if%>            
		
		<%if pFreeShipping=-1 and (pRealTimeShipping="none" or pRealTimeShipping="ups") then%>
		 <br><br><img src="images/freeShippingTruck.gif"> <%=getMsg(32,"Free Shipping")%>         
		<%end if%>
            
<!--#include file="comersus_optGetRelatedProducts.asp"-->                          
            
</form>                         
<%If pShowSuggestionBox="-1" Then%>
<form action="comersus_suggestionBoxExec.asp" method="post">
 <br><br><b><%=getMsg(727,"Sugg Box")%></b><br>
 <%=getMsg(728,"Your opin")%> &nbsp;<%=pCompany%>. <%=getMsg(729,"If you have...")%>
 <br><textarea name="suggestion" cols="40" rows="4"></textarea><br>
 <input type="hidden" name="idProduct" value="<%=pIdProduct%>">
 <input type="hidden" name="sku" value="<%=pSku%>">
 <input type="hidden" name="description" value="<%=pDescription%>"> 
 
 <%if pVerificationCodeEnabled="-1" then%>
  <br>Verification code
  <!--#include file="comersus_insertImageCode.asp"-->  
  <input type="text" name="verificationCode" size=6>
 <%end if%>
 
 <br><input type="submit" name="submit" value="Submit">
</form>
<%end if%>

<!--#include file="footer.asp"--> 
<%
call closeDb()
set pIdproduct		= Nothing
set pDescription	= Nothing
set pDetails		= Nothing
set pListPrice		= Nothing
set pImageUrl		= Nothing
%>
