<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: off line payment form
%>
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/stringFunctions.asp"--> 
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/screenMessages.asp"--> 
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/adSenseFunctions.asp"-->
<%
on error resume next

dim connTemp, rsTemp

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanyLogo	 	= getSettingKey("pCompanyLogo")
pCompanySlogan	 	= getSettingKey("pCompanySlogan")
pHeaderKeywords 	= getSettingKey("pHeaderKeywords")

pAuctions		= getSettingKey("pAuctions")
pListBestSellers 	= getSettingKey("pListBestSellers")
pNewsLetter 		= getSettingKey("pNewsLetter")
pPriceList 		= getSettingKey("pPriceList")
pStoreNews 		= getSettingKey("pStoreNews")

pOffLineCardsAccepted	= getSettingKey("pOffLineCardsAccepted")

pIdOrder 		= getUserInput(request("idOrder"),20)
pName 			= getUserInput(request("name"),50) &" " &getUserInput(request("lastName"),50)
pAddress		= getUserInput(request("address"),100) &" " &getUserInput(request("ciy"),100)

%>
<!--#include file="header.asp"--> 
 <br><b><%=getMsg(500,"CC payment")%></b>
 
 <%if pStoreFrontDemoMode="-1" then%>
    <br><br><%=getMsg(501,"Comp with")%>
    <br><%=getMsg(502,"gateways list")%>
    <br><br><i><%=getMsg(503,"SSL")%></i>    
 <%end if%> 
 
 <br><br> 
 <form action = "comersus_offLinePaymentExec.asp" method="post">
 <table width="500" border="0">
  <input type = "hidden" name = "orderTotal" value ="<%=request.querystring("ordertotal")%>"> 
  <input type = "hidden" name = "idOrder"    value ="<%=pIdOrder%>">
  <tr> 
      <td><%=getMsg(512,"type")%></td>
      <td>   
        <select name="cardType">
        <%if instr(pOffLineCardsAccepted,"VI")<>0 then%>
          <option value="V">Visa 
        <%end if%>  
        <%if instr(pOffLineCardsAccepted,"MA")<>0 then%>
          <option value="M">MasterCard      
        <%end if%>                 
        <%if instr(pOffLineCardsAccepted,"AM")<>0 then%>
          <option value="A">American Express           
        <%end if%>                 
        <%if instr(pOffLineCardsAccepted,"DIN")<>0 then%>  
          <option value="D">Diners Club
        <%end if%>                 
        <%if instr(pOffLineCardsAccepted,"DIS")<>0 then%>
          <option value="Discover">Discover
        <%end if%>                 
        <%if instr(pOffLineCardsAccepted,"JC")<>0 then%>
          <option value="JCB">JCB
        <%end if%>  
         </select>
         </td>
 </tr>
 
 <tr> 
      <td><%=getMsg(504,"number")%></td>
      <td colspan="3">   
        <%if pStoreFrontDemoMode ="-1" then%>
        <input type="text" name="cardNumber" value="4300000000000"> 
      <%else%>
        <input type="text" name="cardNumber"> 
      <%end if%>
         </td>
 </tr>
 <tr> 
      <td width="205" height="32"><%=getMsg(505,"expi")%></td>
      <td colspan="3" height="32"><%=getMsg(506,"month")%> 
        <select name="expMonth">
          <option value="1">1 
          <option value="2">2 
          <option value="3">3 
          <option value="4">4 
          <option value="5">5 
          <option value="6">6 
          <option value="7">7 
          <option value="8">8 
          <option value="9">9 
          <option value="10">10 
          <option value="11">11 
          <option value="12" selected>12 
        </select>
        <%=getMsg(507,"year")%> 
        <select name="ExpYear">                                        
          <option value="2007" selected>2007           
          <option value="2008">2008           
          <option value="2009">2009           
          <option value="2010">2010         
          <option value="2011">2011         
          <option value="2012">2012           
          <option value="2013">2013           
          <option value="2014">2014           
          <option value="2015">2015           
          <option value="2016">2016                     
        </select><br>
         </td>
 </tr>
 
 <tr> 
      <td><%=getMsg(508,"CVV2")%></td>
      <td colspan="3">   
        <input type="text" name="seqCode" size="6">
         </td>
 </tr>
 
 <tr> 
      <td><%=getMsg(509,"name")%></td>
      <td colspan="3">  
        <input type="text" name="nameOnCard" value="<%=pName%>">
         </td>
 </tr>    
<tr> 
      <td><%=getMsg(510,"statement")%></td>
      <td colspan="3">   
        <textarea name="statementAddress"><%=pAddress%></textarea>
         </td>
 </tr>
        
  <tr>
   <td>       
   <input type="submit" value="<%=getMsg(511,"post")%>">   
   </td>
  </tr>
  
 </table>
</form>
<!--#include file="footer.asp"-->
<%call closeDb()%> 