<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: show order details for one customer
%>

<!--#include file="comersus_customerLoggedVerify.asp"-->
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/screenMessages.asp"-->
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/stringFunctions.asp"--> 
<!--#include file="../includes/cartFunctions.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->  
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/miscFunctions.asp"--> 
<!--#include file="../includes/customerTracking.asp"-->  
<!--#include file="../includes/adSenseFunctions.asp"-->
<% 

on error resume next 

dim mySQL, connTemp, rsTemp, rsTemp2, rsTemp3

redim fileNameList(19)
redim saveAsList(19)

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanySlogan	 	= getSettingKey("pCompanySlogan")
pAboutUsLink 		= getSettingKey("pAboutUsLink")
pMoneyDontRound	 	= getSettingKey("pMoneyDontRound")
pHeaderKeywords 	= getSettingKey("pHeaderKeywords")

pAffiliatesStoreFront   = getSettingKey("pAffiliatesStoreFront")
pAuctions		= getSettingKey("pAuctions")
pListBestSellers 	= getSettingKey("pListBestSellers")
pNewsLetter 		= getSettingKey("pNewsLetter")
pPriceList 		= getSettingKey("pPriceList")
pStoreNews 		= getSettingKey("pStoreNews")
pOneStepCheckout	= getSettingKey("pOneStepCheckout")
pDownloadDigitalGoods   = getSettingKey("pDownloadDigitalGoods")
pDateSwitch		= getSettingKey("pDateSwitch")
pOrderPrefix		= getSettingKey("pOrderPrefix")
pIdRmaPrefix		= getSettingKey("pIdRmaPrefix")
pAllowDelayPayment	= getSettingKey("pAllowDelayPayment")
pByPassShipping		= getSettingKey("pByPassShipping")
pShowSearchBox		= getSettingKey("pShowSearchBox")
pAdSenseClient 		= getSettingKey("pAdSenseClient")

pDownloadGoodsType      = UCase(getSettingKey("pDownloadGoodsType"))
pDaysToDownloadGoods	= Cdbl(getSettingKey("pDaysToDownloadGoods"))
pDigitalGoodsWatchVideo	= getSettingKey("pDigitalGoodsWatchVideo")

' session
pIdCustomer 		= getSessionVariable("idCustomer",0)
pIdDbSessionCart	= getSessionVariable("idDbSessionCart",0)

' form
pIdOrder		= getUserInput(request("idOrder"),12)

' extract real idorder (without prefix)
pIdOrder 		= removePrefix(pIdOrder,pOrderPrefix)

if trim(pIdOrder)="" then
  response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(517,"Missing order number"))
end if

call customerTracking("comersus_customerShowOrderDetails.asp", request.querystring)

' get order data and check if the order belongs to that customer

mySQL="SELECT orders.idOrder, name, lastName, orderstatus, phone, email, viewed, orders.address AS address, orders.state AS state, orders.stateCode as stateCode, orders.zip AS zip, orders.city AS city, orders.countryCode AS countryCode, taxAmount, orderDate, shipmentDetails, paymentDetails, discountDetails, obs, total, details, orders.shippingAddress, orders.shippingCity, orders.shippingState, orders.shippingStateCode, orders.shippingZip, orders.shippingCountryCode, orders.idCustomerType, shipmentTracking, transactionResults FROM orders, customers WHERE orders.idcustomer=customers.idcustomer AND orders.idOrder=" &pIdOrder&" AND customers.idCustomer="&pIdCustomer

call getFromDatabase(mySQL, rstemp, "customerShowOrderDetails") 

if  rstemp.eof then 
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(518,"Cannot locate the order inside your history"))
end if 

' Today
pToday=Date
  
' contact

pName			= rstemp("name")
pLastName		= rstemp("lastName")
pEmail			= rstemp("email")
pPhone			= rstemp("phone")   

' order
pOrderDate		= rstemp("orderDate")   
pDetails		= rstemp("details")
pTotal			= rstemp("total")
pPaymentDetails		= rstemp("paymentDetails")
pShipmentDetails	= rstemp("shipmentDetails")
pDiscountDetails	= rstemp("discountDetails")
pTaxAmount		= rstemp("taxAmount")
pAddress		= rstemp("address")
pZip			= rstemp("zip")
pState			= rstemp("state")
pStateCode		= rstemp("stateCode")
pCity			= rstemp("city")
pCountryCode		= rstemp("countryCode")
pShippingaddress	= rstemp("shippingAddress")
pShippingzip		= rstemp("shippingzip")
pShippingstate		= rstemp("shippingstate")
pShippingstateCode	= rstemp("shippingstateCode")
pShippingcity		= rstemp("shippingcity")
pShippingcountryCode	= rstemp("shippingCountryCode")
pOrderStatus		= rstemp("orderStatus")
pIdCustomerType		= rstemp("idCustomerType")
pShipmentTracking	= rstemp("shipmentTracking")
pTransactionResults	= rstemp("transactionResults")
   
' init download array
for f=0 to 19
 fileNameList(f) = ""
 saveAsList(f)   = ""
next   

' remove 0 amount for payment method
pPaymentDetails=replace(pPaymentDetails,"$0.00","")

%> 

<!--#include file="header.asp"--> 
<div class="description"><%=getMsg(519,"details")%></div><br><br>
<br>
<table width="460" border="0">
  <tr bgcolor="#eeeecc"> 
    <td colspan="3">
      <b><%=getMsg(520,"order")%>&nbsp; <%=pOrderPrefix&pIdOrder%>, <%=getMsg(521,"date")%>: <%=formatDate(pOrderDate)%></b> </td>
  </tr>
  <tr> 
    <td colspan="2"><%=getMsg(522,"name")%></td>
    <td><%=pName&" "&pLastName%></td>
  </tr>
  <tr> 
    <td colspan="2"><%=getMsg(523,"email")%></td>
    <td><%=pEmail%></td>
  </tr>
  <tr> 
    <td colspan="2"><%=getMsg(524,"phone")%></td>
    <td><%=pPhone%></td>
  </tr>
  <tr> 
    <td colspan="2"><%=getMsg(525,"bill add")%></td>
    <td><%=pAddress& ", " &pCity& " " &pState&pStateCode& " " &pZip& " " &getCountryName(pCountryCode)%> 
    </td>
  </tr>  
  <tr> 
    <td colspan="2"><%=getMsg(526,"ship add")%></td>
    <td> 
    <%if pShippingAddress<>"" then%>
     <%=pShippingaddress &", " &pShippingCity& " " &pShippingState&pShippingStateCode& " " &pShippingZip& " " & getCountryName(pShippingcountryCode)%> 
    <%else%>
     <%=getMsg(527,"same")%>
    <%end if%>
    </td>
  </tr>     
  <tr> 
    <td colspan="2"><%=getMsg(528,"payment")%></td>
    <td><%=pPaymentDetails%>
    <%if pTransactionResults<>"" then%>
    (<%=pTransactionResults%>)
    <%end if%>
    </td>
  </tr>
  
  <%if pShipmentTracking<>"" then%>
  <tr> 
    <td colspan="2"><%=getMsg(529,"ship track")%></td>
    <td><%=pShipmentTracking%>
   
    <%if instr(lcase(pShipmentDetails),"fedex")<>0 then%>
     <a href="http://www.fedex.com/Tracking?sum=n&ascend_header=1&clienttype=dotcom&spnlk=spnl0&initial=n&cntry_code=us&tracknumber_list=<%=pShipmentTracking%>&language=english&track_number_0=<%=pShipmentTracking%>&track_number_replace_0=<%=pShipmentTracking%>&resubmit_all=Resubmit" target="_blank">Track</a>
    <%end if%>     
    
    <%if instr(lcase(pShipmentDetails),"ups")<>0 then%>
     <a href="http://wwwapps.ups.com/WebTracking/processInputRequest?HTMLVersion=5.0&loc=en_US&Requester=UPSHome&tracknum=<%=pShipmentTracking%>&AgreeToTermsAndConditions=yes&ignore=&track.x=22&track.y=10" target="_blank">Track</a>
    <%end if%>       
    
    </td>
  </tr>
  <%end if%>
  
  <%if pDiscountDetails<>"" then%>
  <tr> 
    <td colspan="2"><%=getMsg(530,"discounts")%></td>
    <td><%=pDiscountDetails%></td>
  </tr>
  <%end if%>
  
  <tr>
    <td colspan="2"><%=getMsg(531,"status")%></td>
    <td>   
    <%select case pOrderStatus
	case 1
   		response.write getMsg(549,"pending")
   		
   		if pAllowDelayPayment="-1" then 
   		
   		  ' shows link to pay if CC has not been entered for off line payment 
   		  mySQL="SELECT * FROM creditCards WHERE idOrder=" &pIdOrder   		
		  call getFromDatabase(mySQL, rstemp2, "customerShowOrderDetails") 
		  
		  if rstemp2.eof then
   		    %>&nbsp;<a href='comersus_optDelayPaymentForm.asp?idOrder=<%=pIdOrder%>'><%=getMsg(532,"pay")%></a><%   		  
   		  end if
   		     		  
   		end if
   		
   		%>&nbsp;<a href="comersus_customerCancelOrderExec.asp?idOrder=<%=pOrderPrefix&pIdOrder%>"><%=getMsg(533,"cancel")%></a><%
   		
	case 2
   		response.write getMsg(534,"delivered")
	case 3
   		response.write getMsg(535,"cancelled")
   	case 4
   		response.write getMsg(536,"paid")
   	case 5
   		response.write getMsg(537,"cback")
   	case 6
   		response.write getMsg(538,"rfund")
	end select%>	                                
    </td>
  </tr>
  <tr> 
    <td colspan="2">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3">
 <table width="460" border="0" cellspacing="0" cellpadding="0">    
  <tr> 
    <td><b><%=getMsg(539,"sku")%></b></td>
    <td><b><%=getMsg(540,"desc")%></b></td>
    <td><b><%=getMsg(541,"qty")%></b></td>
    <td><b><%=getMsg(542,"variation")%></b></td>
    <td><b><%=getMsg(543,"price")%></b></td>
  </tr>
    <%
' get cartRow contents
    
mySQL="SELECT cartRows.idCartRow, products.idProduct, products.sku, products.description, cartRows.quantity, cartRows.unitPrice, emailText, personalizationDesc FROM cartRows, dbSessionCart, products WHERE dbSessionCart.idDbSessionCart=cartRows.idDbSessionCart AND cartRows.idProduct=products.idProduct AND dbSessionCart.idOrder=" &pIdOrder

call getFromDatabase(mySQL, rstemp2, "customerShowOrderDetails") 

pSubTotal = 0
i	  = 0

do while not rstemp2.eof

	pIdCartRow		= rstemp2("idCartRow")
	pDescription		= rstemp2("description")		
	pDigitalGoodsFile	= rstemp2("emailText")
	pPersonalizationDesc	= rstemp2("personalizationDesc")	        
    %>
    
  
  <tr> 
    <td><a href="comersus_viewItem.asp?idProduct=<%=rstemp2("idProduct")%>"><%=rstemp2("sku")%></a> </td>
    <td><%=left(pDescription,30)%>
    <%if pPersonalizationDesc<>"" then response.write "&nbsp;(" &pPersonalizationDesc& ")"%>
    <%
    ' if the order is paid/delivered, setting is enabled and the field has a zip file     
    
    if (pOrderStatus=2 or pOrderStatus=4) and pDownloadDigitalGoods="-1" and isDownloadFile(pDigitalGoodsFile) and pDownloadGoodsType<>"ZIP" then
    
     ' set the file origin and destination name
     i=i+1
     fileNameList(i)  ="../digitalGoods/"& trim(pDigitalGoodsFile)
     
     ' compile file name as order name + file name
     saveAsList(i)	  =pOrderPrefix&pIdOrder&"_"&trim(pDigitalGoodsFile)
     
     ' verify expiration
     if datediff("d",pOrderDate,pToday)>pDaysToDownloadGoods then
       %>&nbsp; <%=getMsg(544,"expired")%><%
     else
       %>&nbsp;<a href="comersus_optDownloadDigitalGoods.asp?fileIndex=<%=i%>&component=<%=pDownloadGoodsType%>"><img src="images/download.gif" border=0></a>
       <%
     end if ' digital goods
    end if%>
    <%if (pOrderStatus=2 or pOrderStatus=4) and pDownloadDigitalGoods="-1" and trim(pDigitalGoodsFile)<>"" and pDownloadGoodsType="ZIP" then
     ' set the file origin and destination name
     i=i+1
     fileNameList(i)  ="../digitalGoods/"& trim(pDigitalGoodsFile)
     
     ' compile file name as order name + file name
     saveAsList(i)	  =pOrderPrefix&pIdOrder&"_"&trim(pDigitalGoodsFile)
     
     if datediff("d",pOrderDate,pToday)>pDaysToDownloadGoods then
      %>&nbsp; <%=getMsg(544,"expired")%><%
     else
      %><a href="comersus_optDownloadDigitalGoodsZip.asp?fileIndex=<%=i%>&idOrder=<%=pOrderPrefix&pIdOrder%>" target="_blank"><%=getMsg(545,"dload")%></a><%            
      end if 
      
    end if%>   

   <%if pDigitalGoodsWatchVideo="-1" and isVideoFile(pDigitalGoodsFile) and (pOrderStatus=2 or pOrderStatus=4) then%> 
   <%
   ' set the file origin and destination name
     i=i+1
     fileNameList(i)  =trim(pDigitalGoodsFile)     
    ' verify expiration
     if datediff("d",pOrderDate,pToday)>pDaysToDownloadGoods then
       %>&nbsp; <%=getMsg(544,"expired")%><%
     else%>   
      <a href="comersus_optWatchVideo.asp?fileIndex=<%=i%>" target="_blank"><img src="images/icon-video.jpg" border=0 alt="Watch Video now!"></a>     
     <%end if%>
   <%end if%>
         
    </td>
    <td><%=rstemp2("quantity")%></td>
    <td><%=getCartRowOptionals(pIdCartRow)%></td>
    <td><%=pCurrencySign &  money(getCartRowPrice(pIdCartRow))%></td>
  </tr>     
     
<%rstemp2.movenext
loop

session("fileName")	=filenameList
session("saveAs")	=saveAsList
%>          
                            
</table>         
</td> </tr>
 
 <tr> 
    <td colspan="2">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>

<%if pByPassShipping="0" then%>  
  <tr> 
    <td colspan="2"><%=getMsg(546,"ship")%></td>
    <td><%=pShipmentDetails%></td>
  </tr>    
<%end if%>  

  <tr> 
    <td colspan="2"><%=getMsg(547,"tax")%></td>
    <td ><%=pCurrencySign & money(pTaxAmount)%></td>
  </tr>
  
  <tr> 
    <td colspan="2"><%=getMsg(548,"total")%></td>
    <td><%=pCurrencySign & money(pTotal)%></td>
  </tr>    

  <tr> 
    <td colspan="3">&nbsp;</td>
  </tr>    
   
<%
 ' check RMA	
 
  mySQL="SELECT idRma, rmaDate, rmaStatus, customerReasons, adminReasons FROM RMA WHERE idOrder=" &pIdOrder
  call getFromDatabase (mySql, rsTemp2,"customerShowOrderDetails")
  
  if not rstemp2.eof then

%>
 
  <tr bgcolor="#eeeecc"> 
    <td colspan="3"><b><%=getMsg(738,"rma")%></b>
    </td>
  </tr>  
   
  <tr> 
    <td colspan="2">Date</td>
    <td><%=formatDate(rstemp2("rmaDate"))%></td>
  </tr>
 
 <tr> 
    <td colspan="2">Status</td>
    <td><%
     select case rstemp2("rmaStatus")
     
	case 1
   		response.write getMsg(739,"pending")
	case 2
   		response.write getMsg(740,"approved") &" (" &pIdRmaPrefix&rstemp2("idRma")&")"
	case 3
	 	response.write getMsg(741,"rejected")	
    end select
    %></td>
  </tr>
  
  <tr> 
    <td colspan="2">Customer reasons</td>
    <td><%=rstemp2("customerReasons")%></td>
  </tr>
  
  <tr> 
    <td colspan="2">Admin reasons</td>
    <td><%=rstemp2("adminReasons")%></td>
  </tr> 

 <%if rstemp2("rmaStatus")=2 then%>
   <tr> 
     <td colspan="2">Shipping Label</td>
     <td><a target="_blank" href="comersus_rmaGenerateLabelExec.asp?idRma=<%=pIdRmaPrefix&rstemp2("idRma")%>">Generate</a></td>
   </tr> 
 <%end if%>  	 	     

<%end if%>  	 	     
</table>
<br> 
<!--#include file="footer.asp"-->
<%call closeDb()%>

