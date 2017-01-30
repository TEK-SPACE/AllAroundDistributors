<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: show general conditions
%>

<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/screenMessages.asp"--> 
<!--#include file="../includes/currencyFormat.asp"-->
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/cartFunctions.asp"--> 
<!--#include file="../includes/adSenseFunctions.asp"-->

<%
on error resume next

dim connTemp, rsTemp

' get settings 
pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanySlogan 		= getSettingKey("pCompanySlogan")
pAboutUsLink 		= getSettingKey("pAboutUsLink")
pStoreLocation		= getSettingKey("pStoreLocation")
pHeaderKeywords 	= getSettingKey("pHeaderKeywords")
pShowSearchBox 		= getSettingKey("pShowSearchBox")
pAdSenseClient 		= getSettingKey("pAdSenseClient")
pAuctions		= getSettingKey("pAuctions")
pAllowNewCustomer	= getSettingKey("pAllowNewCustomer")
pNewsLetter 		= getSettingKey("pNewsLetter")
pStoreNews 		= getSettingKey("pStoreNews")
pSuppliersList		= getSettingKey("pSuppliersList")
pRssFeedServer		= getSettingKey("pRssFeedServer")
pShowNews		= getSettingKey("pShowNews")
pVerificationCodeEnabled= getSettingKey("pVerificationCodeEnabled")

pIdCustomer	 	= getSessionVariable("idCustomer",0)
pIdCustomerType  	= getSessionVariable("idCustomerType",1)
pIdDbSessionCart	= getSessionVariable("idDbSessionCart",0)
%>

<!--#include file="header.asp"--> 
<b><%=getMsg(337,"Survey")%></b>
<br><br><%=getMsg(338,"Take the time...")%>
<br>
<form action="comersus_storeSurveyExec.asp" method="post" name="survey">
<table width="100%" border="0">
  <tr> 
    <td>1. <%=getMsg(339,"Visited")%></td>
    <td> 
      <input type="text" name="1" value=1 size=2>
    </td>
  </tr>
  <tr> 
    <td>2. <%=getMsg(340,"You are")%></td>
    <td> 
      <select name="2">
        <option selected><%=getMsg(341,"Prof")%></option>
        <option><%=getMsg(342,"Technician")%></option>
        <option><%=getMsg(343,"Businessman")%></option>
        <option><%=getMsg(344,"Civil")%></option>
        <option><%=getMsg(345,"Law")%></option>
        <option><%=getMsg(346,"Student")%></option>
        <option><%=getMsg(347,"Pensioner")%></option>
        <option><%=getMsg(348,"Teacher")%></option>
        <option><%=getMsg(349,"Housewife")%></option>
        <option><%=getMsg(350,"Other")%></option>
      </select>
    </td>
  </tr>
  <tr> 
    <td>3. <%=getMsg(351,"Learn about us")%></td>
    <td> 
      <select name="3">
        <option selected><%=getMsg(352,"Search eng")%></option>
        <option><%=getMsg(353,"Link")%></option>
        <option><%=getMsg(354,"Ad")%></option>
        <option><%=getMsg(355,"Other")%></option>        
      </select>
    </td>
  </tr>
  
  <tr> 
    <td>4. <%=getMsg(356,"Found")%></td>
    <td> 
      <select name="4">
        <option selected><%=getMsg(357,"Yes")%></option>
        <option><%=getMsg(358,"No")%></option>        
      </select>
    </td>
  </tr>
  
    <tr> 
    <td>5. <%=getMsg(359,"Navigate")%></td>
    <td> 
      <select name="5">
        <option selected><%=getMsg(357,"Yes")%></option>
        <option><%=getMsg(358,"No")%></option>        
      </select>
    </td>
  </tr>
  
   <tr> 
    <td>6. <%=getMsg(360,"Features")%></td>
    <td> 
      <select name="6">
        <option selected><%=getMsg(361,"New prods")%></option>
        <option><%=getMsg(362,"Aucti")%></option>        
        <option><%=getMsg(363,"Sales")%></option>        
        <option><%=getMsg(364,"Best sellers")%></option>        
        <option><%=getMsg(365,"Last chance")%></option>        
        <option><%=getMsg(366,"Wap")%></option>                
      </select>
    </td>
  </tr>
  
    <tr> 
    <td>7. <%=getMsg(367,"Payment")%></td>
    <td> 
      <select name="7">
        <option selected><%=getMsg(357,"Yes")%></option>
        <option><%=getMsg(358,"No")%></option>        
      </select>
    </td>
  </tr>
  
  <tr> 
    <td>8. <%=getMsg(368,"Shipment")%></td>
    <td> 
      <select name="8">
 	<option selected><%=getMsg(357,"Yes")%></option>
        <option><%=getMsg(358,"No")%></option>   
      </select>
    </td>
  </tr>
  
  <tr> 
    <td>9. <%=getMsg(369,"Prices")%></td>
    <td> 
      <select name="9">        
        <option selected><%=getMsg(370,"ok")%></option>        
        <option><%=getMsg(371,"high")%></option>
        <option><%=getMsg(372,"low")%></option>        
      </select>
    </td>
  </tr>
  
  <tr> 
    <td>10. <%=getMsg(373,"privacy")%></td>
    <td> 
      <select name="10">
 	<option selected><%=getMsg(357,"Yes")%></option>
        <option><%=getMsg(358,"No")%></option>                
      </select>
    </td>
  </tr>
  
  <tr> 
    <td>11. <%=getMsg(374,"phone")%></td>
    <td> 
      <select name="11">
 	<option selected><%=getMsg(357,"Yes")%></option>
        <option><%=getMsg(358,"No")%></option>                   
      </select>
    </td>
  </tr>
  
  <tr> 
    <td>12. <%=getMsg(375,"discount")%></td>
    <td> 
      <select name="12">
 	<option selected><%=getMsg(357,"Yes")%></option>
        <option><%=getMsg(358,"No")%></option>                    
      </select>
    </td>
  </tr>
  
  <tr> 
    <td>13. <%=getMsg(376,"affiliate")%></td>
    <td> 
      <select name="13">
 	<option selected><%=getMsg(357,"Yes")%></option>
        <option><%=getMsg(358,"No")%></option>                
      </select>
    </td>
  </tr>
  
  <tr> 
    <td>14. <%=getMsg(377,"suggestions")%></td>
    <td> 
      <textarea name="14" cols="30"></textarea>
    </td>
  </tr>  

 <%if pVerificationCodeEnabled="-1" then%>
  <tr> 
    <td><!--#include file="comersus_insertImageCode.asp"--> Verification code</td>
    <td> 
        <input type="text" name="verificationCode" size=6> 
    </td>
  </tr>
 <%end if%>
   <tr> 
    <td colspan=2>&nbsp;</td>
  </tr>
 </table>
<input type="submit" name="smit1" value="<%=getMsg(378,"send")%>">
</form>

<!--#include file="footer.asp"--> 
<%call closeDb()%>
