<%
' Comersus 4.2x Sophisticated Cart
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com
' Details: verifies if cart is empty in order to allow userauthenticateform and display first userauthenticateform screen
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/getSettingKey.asp"-->
<!--#include file="../includes/screenMessages.asp"-->
<!--#include file="../includes/cartFunctions.asp"-->
<!--#include file="../includes/currencyFormat.asp"-->
<!--#include file="../includes/databaseFunctions.asp"-->
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/sessionFunctions.asp"-->
<!--#include file="../includes/adSenseFunctions.asp"-->
<%
dim mySQL, connTemp, rsTemp

on error resume next

pIdCustomer		= getSessionVariable("idCustomer",0)
pIdCustomerType		= getSessionVariable("idCustomerType",1)

if pIdCustomer<>0 then
 response.redirect "comersus_customerUtilitiesMenu.asp"
end if

' get settings

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanySlogan 		= getSettingKey("pCompanySlogan")
pHeaderKeywords 	= getSettingKey("pHeaderKeywords")
pAdSenseClient 		= getSettingKey("pAdSenseClient")
pAuctions		= getSettingKey("pAuctions")
pListBestSellers 	= getSettingKey("pListBestSellers")
pNewsLetter 		= getSettingKey("pNewsLetter")
pPriceList 		= getSettingKey("pPriceList")
pStoreNews 		= getSettingKey("pStoreNews")
pOneStepCheckout	= getSettingKey("pOneStepCheckout")
pForgotPassword		= getSettingKey("pForgotPassword")
pAllowNewCustomer	= getSettingKey("pAllowNewCustomer")

pRedirect		= getUserInputL(request.querystring("redirect"),20)
pIdProduct		= getUserInputL(request.querystring("idProduct"),20)

pIdDbSessionCart	= getSessionVariable("idDbSessionCart",0)
%>

<!--#include file="header.asp"-->
<br><br><b><%=getMsg(136,"Login")%></b>
<form method="post" name="auth" action="comersus_customerAuthenticateExec.asp">
<input type="hidden" name="redirect" value="<%=pRedirect%>">
<input type="hidden" name="idProduct" value="<%=pIdProduct%>">
    <table width="400" border="0" align="left">
      <tr>
        <td width="24%"><%=getMsg(137,"Em")%></td>
        <td width="76%">
        <%if pStoreFrontDemoMode="-1" then%>
          <input type=text name=email value="test@comersus.com" size=30>                    
        <%else%>
          <input type=text name=email value="" size=30>          
        <%end if%>
        </td>
      </tr>
      <tr>
        <td width="24%"><%=getMsg(138,"Pwd")%></td>
        <td width="76%">
        <%if pStoreFrontDemoMode="-1" then%>
          <input type=password name=password value="123456" size=20>          
        <%else%>
          <input type=password name=password value="" size=20>          
        <%end if%>
        </td>
      </tr>      
     
      <td width="24%">&nbsp;</td>
        <td width="76%">&nbsp;
        </td>
      </tr>
      <tr>
        <td width="24%">
          <input type="submit" name="Submit" value="<%=getMsg(139,"Login")%>">
        </td>
      <td width="76%">
      
        <%if pAllowNewCustomer="-1" then%>
         <i><%=getMsg(140,"Not a member")%>&nbsp; <a href="comersus_customerRegistrationForm.asp"><%=getMsg(141,"Here")%></a></i><br>
        <%end if%>
        <%if pForgotPassword="-1" then%>
         <i><%=getMsg(142,"Forgot")%>&nbsp;<a href="comersus_optForgotPasswordForm.asp"><%=getMsg(143,"Here")%></a></i></td>
        <%end if%>

      </tr>
    </table>
</form>
<br><br>
<!--#include file="footer.asp"-->
<%call closeDb()%>
