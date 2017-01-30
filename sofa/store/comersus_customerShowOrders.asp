<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com   
' Details: show a list of orders
%>
<!--#include file="comersus_customerLoggedVerify.asp"-->
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/screenMessages.asp"-->
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/stringFunctions.asp"--> 
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->  
<!--#include file="../includes/customerTracking.asp"-->  
<!--#include file="../includes/cartFunctions.asp"-->
<!--#include file="../includes/adSenseFunctions.asp"-->
<% 

on error resume next 

dim connTemp, rsTemp, mySql, counter, pTotalPages, count


' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanySlogan	 	= getSettingKey("pCompanySlogan")
pAboutUsLink 		= getSettingKey("pAboutUsLink")
pMoneyDontRound	 	= getSettingKey("pMoneyDontRound")

pAffiliatesStoreFront   = getSettingKey("pAffiliatesStoreFront")
pAuctions		= getSettingKey("pAuctions")
pListBestSellers 	= getSettingKey("pListBestSellers")
pNewsLetter 		= getSettingKey("pNewsLetter")
pPriceList 		= getSettingKey("pPriceList")
pStoreNews 		= getSettingKey("pStoreNews")
pOneStepCheckout	= getSettingKey("pOneStepCheckout")
pAllowDelayPayment	= getSettingKey("pAllowDelayPayment")
pOrderPrefix		= getSettingKey("pOrderPrefix")
pDateSwitch		= getSettingKey("pDateSwitch")
pShowSearchBox		= getSettingKey("pShowSearchBox")

' session
pIdCustomer     	= getSessionVariable("idCustomer","0")
pIdCustomerType     	= getSessionVariable("idCustomerType","1")
pIdDbSessionCart	= getSessionVariable("idDbSessionCart",0)

if request.querystring("pCurPage") = "" then
    pCurPage = 1 
else
    pCurPage = getUserInput(request.QueryString("pCurPage"),4)
end if

Const pNumPerPage = 10

call customerTracking("comersus_customerShowOrders.asp", request.querystring)

' get orders, for demo store get only 10 orders
if pStoreFrontDemoMode="-1" and lCase(pDataBase)<>"mysql" then
 mySQL="SELECT TOP 10 orders.idOrder, orderDate, total, orderStatus FROM orders, customers WHERE orders.idCustomer=customers.idCustomer AND customers.idCustomer=" &pIdCustomer& " ORDER by orders.idOrder DESC"
else
 mySQL="SELECT orders.idOrder, orderDate, total, orderStatus FROM orders, customers WHERE orders.idCustomer=customers.idCustomer AND customers.idCustomer=" &pIdCustomer& " ORDER by orders.idOrder DESC"
end if

call getFromDatabasePerPage (mySql, rsTemp,"customerShowOrders")

if rstemp.eof then 
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(550,"Empty order list"))      
end if           

rstemp.MoveFirst
rstemp.PageSize = pNumPerPage

pTotalPages = rstemp.PageCount
rstemp.AbsolutePage = pCurPage

Count = 0

%> 

<!--#include file="header.asp"--> 
<b><%=getMsg(551,"history")%></b><br><br>
<table align=left border=0 cellspacing=1>  
  <tr> 
    <td bgcolor=#eeeecc><b><%=getMsg(552,"date")%></b></td>
    <td bgcolor=#eeeecc><b><%=getMsg(553,"order")%></b></td>
    <td bgcolor=#eeeecc><b><%=getMsg(554,"total")%></b></td>
    <td bgcolor=#eeeecc><b><%=getMsg(556,"items")%></b></td>
    <td bgcolor=#eeeecc><b><%=getMsg(555,"status")%></b></td>
    <td bgcolor=#eeeecc><b><%=getMsg(738,"rma")%></b></td>
  </tr>
  <%do while not rstemp.eof And Count < rstemp.PageSize

     
     pIdOrder		= rstemp("idOrder")
     pOrderDate		= rstemp("orderDate")
     pTotal		= rstemp("total")
     pOrderStatus	= rstemp("orderStatus")
     
     mySQL="SELECT SUM(quantity) AS sumQTY FROM cartRows, dbSessionCart WHERE dbSessionCart.idDbSessionCart=cartRows.idDbSessionCart AND dbSessionCart.idOrder=" &pIdOrder
     call getFromDatabase (mySql, rsTemp2,"customerShowOrders")
     
     if not isNull(rstemp2("sumQty")) then
      pSumQty		= rstemp2("sumQty")
     else
      pSumQty		= 0
     end if
   %>
  <tr> 
    <td><%=formatDate(pOrderDate)%></td>
    <td><a href="comersus_customerShowOrderDetails.asp?idOrder=<%=pOrderPrefix&pIdorder%>"><%=pOrderPrefix&pIdorder%></a></td>
    <td><%=pCurrencySign & money(ptotal)%></td>
    <td><%=pSumQty%></td>
    <td><%select case pOrderStatus
	case 1
   		response.write getMsg(549,"pending")
	case 2
   		response.write getMsg(534,"delivered")
	case 3
   		response.write getMsg(535,"cancel")
   	case 4
   		response.write getMsg(536,"paid")
	case 5
   		response.write getMsg(537,"cback")
   	case 6
   		response.write getMsg(538,"rfunded")
	end select
	
	%></td>
	
	<td>
	
	<%
  ' check RMA	
  mySQL="SELECT idRma, rmaStatus FROM RMA WHERE idOrder=" &pIdOrder
  call getFromDatabase (mySql, rsTemp2,"customerShowOrders")
  
  if not rstemp2.eof then
  
     select case rstemp2("rmaStatus")
	case 1
   		response.write getMsg(739,"pending")
	case 2
   		response.write getMsg(740,"approved")
	case 3
	 	response.write getMsg(741,"rejected")
    end select
  else
   	response.write "-"
  end if
  
 %>
	</td>
	    
  </tr>
  <%rstemp.movenext
    Count		= Count +1
  loop%>  
<tr><td colspan=6>&nbsp;</td></tr>
<tr><td colspan=6>
  <form method="post" action="" name="">

<%response.Write("Page "& pCurPage & " of "& pTotalPages)%>
<%

'Display Next / Prev buttons
if pCurPage > 1 then
 Response.Write("<INPUT TYPE=BUTTON VALUE='<' ONCLICK=""document.location.href='comersus_customerShowOrders.asp?pCurPage="& pCurPage - 1 & "';"">")
end If

if CInt(pCurPage) <> CInt(pTotalPages) then
 Response.Write("<INPUT TYPE=BUTTON VALUE='>' ONCLICK=""document.location.href='comersus_customerShowOrders.asp?pCurPage="& pCurPage + 1 & "';"">")
end If
%>
</form>
</td></tr>
</table>
<!--#include file="footer.asp"-->
<%call closeDb()%>