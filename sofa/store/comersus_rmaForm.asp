<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: Add a Rma issue.
%>
<!--#include file="comersus_customerLoggedVerify.asp"-->
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/screenMessages.asp"--> 
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/cartFunctions.asp"--> 
<!--#include file="../includes/adSenseFunctions.asp"-->

<%
on error resume next

dim mySQL, connTemp, rsTemp

' get settings 
pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pStoreLocation	 	= getSettingKey("pStoreLocation")
pHeaderKeywords 	= getSettingKey("pHeaderKeywords")

pAuctions		= getSettingKey("pAuctions")
pAllowNewCustomer	= getSettingKey("pAllowNewCustomer")
pNewsLetter 		= getSettingKey("pNewsLetter")
pStoreNews 		= getSettingKey("pStoreNews")
pSuppliersList		= getSettingKey("pSuppliersList")

pOrderPrefix		= getSettingKey("pOrderPrefix")
pIdCustomer     	= getSessionVariable("idCustomer",0)
pIdDbSessionCart	= getSessionVariable("idDbSessionCart",0)

mySQL="SELECT orders.idOrder, orderDate FROM orders LEFT JOIN RMA ON orders.idOrder = RMA.idOrder WHERE RMA.idOrder IS NULL AND orders.idCustomer=" & pIdCustomer & " AND orderStatus=2 ORDER by orders.idOrder DESC"
call getFromDatabase (mySql, rsTemp,"customerShowOrders")

if rsTemp.eof then 
	response.redirect "comersus_message.asp?message="&Server.Urlencode("Delivered orders not found.")      
end if           
	
%>

<!--#include file="header.asp"-->
<br><br><b>Send a RMA issue.</b>
<form method="post" name="EmF" action="comersus_rmaExec.asp">
    <table width="400" border="0" align="left">
      <tr> 
        <td width="40%"><%=getMsg(283,"Order #")%></td>
        <td width="60%">                     
	  <select name="idOrder">
	  	<%Do While Not rsTemp.eof 
	  		pIdOrder= rsTemp("idOrder")
	  		pDate 	= rsTemp("orderDate")
	  	%>
	          <option value="<%=pIdOrder%>"><%=pOrderPrefix&pIdOrder&" Posted: "&pDate%></option>
		<%
			rsTemp.moveNext
		  loop
		%>  	
          </select>
        </td>
      </tr>
      
      <tr> 
        <td width="40%"><%=getMsg(742,"Reason")%></td>
        <td width="60%">           
          <textarea name=reasons cols="35" rows="10" maxlength="500"></textarea>
        </td>
      </tr>
      
      <tr> 
        <td width="40%"> 
          <input type="submit" name="Submit" value="<%=getMsg(736,"request")%>">
        </td>
        <td width="60%">&nbsp;</td>
      </tr>            
      
    </table>
</form>  
<br><br>
<!--#include file="footer.asp"--> 
<%call closeDb()%>