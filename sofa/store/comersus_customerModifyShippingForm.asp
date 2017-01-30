<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: modify shipping at checkout
%>

<!--#include file="comersus_customerLoggedVerify.asp"-->   
<!--#include file="../includes/screenMessages.asp" --> 
<!--#include file="../includes/databaseFunctions.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/encryption.asp" --> 
<!--#include file="../includes/sessionFunctions.asp"-->  
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/cartFunctions.asp"-->
<!--#include file="../includes/adSenseFunctions.asp"-->

<%
on error resume next

dim connTemp, mySql, rsTemp, rsTemp2

pIdCustomer     	= getSessionVariable("idCustomer",0)
pIdDbSessionCart	= getSessionVariable("idDbSessionCart",0)

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pDisableState		= getSettingKey("pDisableState")
pHeaderKeywords 	= getSettingKey("pHeaderKeywords")

pAuctions		= getSettingKey("pAuctions")
pAllowNewCustomer	= getSettingKey("pAllowNewCustomer")
pNewsLetter 		= getSettingKey("pNewsLetter")
pStoreNews 		= getSettingKey("pStoreNews")
pSuppliersList		= getSettingKey("pSuppliersList")

pEncryptionPassword 	= getSettingKey("pEncryptionPassword")
pEncryptionMethod 	= getSettingKey("pEncryptionMethod")

pCustomerPrefix		= getSettingKey("pCustomerPrefix")

' custom field names
pCustomerFieldName1	= getSettingKey("customerFieldName1")
pCustomerFieldName2	= getSettingKey("customerFieldName2")
pCustomerFieldName3	= getSettingKey("customerFieldName3")

pRedirect		= getUserInput(request.querystring("redirect"),20)

' get customer data

mySql="SELECT shippingName, shippingLastName, shippingAddress, shippingZip, shippingStateCode, shippingState, shippingCity, shippingCountryCode FROM customers WHERE idCustomer=" &pIdCustomer

call getFromDatabase(mySQL, rstemp, "customerModifyForm")

if rstemp.eof then
 response.redirect "comersus_supportError.asp?error="&Server.Urlencode("Cannot locate customer information in customerModifyForm")
end if

pStateCode	= rstemp("shippingStateCode")
pCountryCode	= rstemp("shippingCountryCode")

%> 

<!--#include file="header.asp"-->

<br><br><b><%=getMsg(420,"shipping")%></b><br><br>
<form method="post" name="modCust" action="comersus_customerModifyShippingExec.asp">
  <input type=hidden name=redirect value="<%=pRedirect%>">
  
  <table width="421" border="0">  
    <tr> 
      <td width="168"><%=getMsg(173,"Name")%></td>
      <td width="220">      
        <input type=text name=shippingName value="<%=rstemp("shippingName")%>">
      </td>
    </tr>    
    <tr> 
      <td width="168"><%=getMsg(174,"L Name")%></td>
      <td width="220">      
        <input type=text name=shippingLastName value="<%=rstemp("shippingLastName")%>">
      </td>
    </tr>
    
    <tr> 
      <td width="168"><%=getMsg(179,"Add")%></td>
      <td width="220">         
        <input type=text name=shippingAddress value="<%=rstemp("shippingAddress")%>">
      </td>
    </tr>
    <tr> 
      <td width="168"><%=getMsg(180,"Zip")%></td>
      <td width="220">         
        <input type=text name=shippingZip value="<%=rstemp("shippingZip")%>">
      </td>
    </tr>

<%if pDisableState="0" then%>    
    <tr> 
      <td width="168"><%=getMsg(181,"st code")%></td>
      <td width="220">         
        <%
      ' get stateCodes
      mySQL="SELECT stateCode, stateName FROM stateCodes ORDER by stateName"

      call getFromDatabase(mySQL, rstemp2, "customerModifyForm")
      %>
      <SELECT name=shippingStateCode size=1>
      <OPTION value=""><%=getMsg(186,"Select state")%>
      <%do while not rstemp2.eof
      pStateCode2=rstemp2("stateCode")%>
        <option value="<%=pStateCode2%>"<%
        if pStateCode2=pStateCode then
           response.write "selected"
        end if%>><%=rstemp2("stateName")%>
      <%rstemp2.movenext
      loop%>
      </OPTION>
      </SELECT>
      </td>
    </tr>
    <tr> 
      <td width="168"><%=getMsg(182,"An state")%></td>
      <td width="220">         
       <input type=text name=shippingState value="<%=rstemp("shippingState")%>">
      </td>
    </tr>
<%end if%>

    <tr> 
      <td width="168"><%=getMsg(183,"City")%></td>
      <td width="220">         
        <input type=text name=shippingCity value="<%=rstemp("shippingCity")%>">
      </td>
    </tr>    
    <tr> 
      <td width="168"><%=getMsg(184,"Ctry")%></td>
      <td width="220">                 
        <%
      ' get CountryCodes
      mySQL="SELECT countryCode, countryName FROM countryCodes ORDER by countryName"

      call getFromDatabase(mySQL, rstemp2, "customerModifyForm")
      
      %>
      <SELECT name=shippingCountryCode>
      <OPTION value=""><%=getMsg(187,"Select")%>
      <%do while not rstemp2.eof
      pCountryCode2=rstemp2("countryCode")%>
        <option value="<%=pCountryCode2%>"<%
        if pCountryCode=pCountryCode2 then
            response.write "selected"           
        end if
        %>><%=rstemp2("countryName")%>
      <%rstemp2.movenext
      loop%>
      </OPTION>
      </SELECT>       
      </td>
    </tr>
    
    <tr> 
      <td width="168">&nbsp;</td>
      <td width="220">&nbsp;</td>
    </tr>
    <tr> 
      <td colspan="2">        
          <input type="submit" name="Modify" value="<%=getMsg(185,"Mdf")%>">                            
      </td>
    </tr>
    </table>
   </form>            
<!--#include file="footer.asp"-->
<%call closeDb()%>