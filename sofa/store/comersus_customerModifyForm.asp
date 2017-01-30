<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: modify customer data
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
pIdCustomerType 	= getSessionVariable("idCustomerType",1)
pIdDbSessionCart	= getSessionVariable("idDbSessionCart",0)

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanySlogan	 	= getSettingKey("pCompanySlogan")
pAboutUsLink 		= getSettingKey("pAboutUsLink")
pDisableState		= getSettingKey("pDisableState")

pAuctions		= getSettingKey("pAuctions")
pAllowNewCustomer	= getSettingKey("pAllowNewCustomer")
pNewsLetter 		= getSettingKey("pNewsLetter")
pStoreNews 		= getSettingKey("pStoreNews")
pSuppliersList		= getSettingKey("pSuppliersList")
pShowNews		= getSettingKey("pShowNews")
pHeaderKeywords 	= getSettingKey("pHeaderKeywords")

pEncryptionPassword 	= getSettingKey("pEncryptionPassword")
pEncryptionMethod 	= getSettingKey("pEncryptionMethod")

pCustomerPrefix		= getSettingKey("pCustomerPrefix")

' custom field names
pCustomerFieldName1	= getSettingKey("customerFieldName1")
pCustomerFieldName2	= getSettingKey("customerFieldName2")
pCustomerFieldName3	= getSettingKey("customerFieldName3")

pRedirect		= getUserInput(request.querystring("redirect"),20)

' get customer data

mySql="SELECT name, lastName, customerCompany, phone, email, password, address, zip, state, city, stateCode, countryCode, user1, user2, user3 FROM customers WHERE idCustomer=" &pIdCustomer

call getFromDatabase(mySQL, rstemp, "customerModifyForm")

if rstemp.eof then
 response.redirect "comersus_supportError.asp?error="&Server.Urlencode("Cannot locate customer information in customerModifyForm")
end if

pStateCode	= rstemp("stateCode")
pCountryCode	= rstemp("countryCode")

%> 

<!--#include file="header.asp"-->

<br><br><b><%=getMsg(171,"Pers Inf")%></b><br><br>
<form method="post" name="modCust" action="comersus_customerModifyExec.asp">
  <input type=hidden name=redirect value="<%=pRedirect%>">
  
  <table width="421" border="0">  
      <tr> 
      <td width="168"><%=getMsg(172,"Cust")%></td>
      <td width="220">      
         <%=pCustomerPrefix&pIdCustomer%> 
      </td>
    </tr>
    <tr> 
      <td width="168"><%=getMsg(173,"Name")%></td>
      <td width="220">      
        <input type=text name=customerName value="<%=rstemp("name")%>">
      </td>
    </tr>    
    <tr> 
      <td width="168"><%=getMsg(174,"L Name")%></td>
      <td width="220">      
        <input type=text name=lastName value="<%=rstemp("lastName")%>">
      </td>
    </tr>
    <tr> 
      <td width="168"><%=getMsg(175,"Company")%></td>
      <td width="220">      
        <input type=text name=customerCompany value="<%=rstemp("customerCompany")%>">
      </td>
    </tr>
    <tr> 
      <td width="168"><%=getMsg(176,"Phone")%></td>
      <td width="220">        
       <input type=text name=phone value="<%=rstemp("phone")%>">
      </td>
    </tr>
    <tr> 
      <td width="168"><%=getMsg(177,"Email")%></td>
      <td width="220">   
      <input type=text name=email value="<%=rstemp("email")%>">          
      </td>
    </tr>
    <tr> 
      <td width="168"><%=getMsg(178,"Pwd")%></td>
      <td width="220">         
        <input type=password name=password value="<%=DeCrypt(rstemp("password"), pEncryptionPassword)%>">
      </td>
    </tr>
    <tr> 
      <td width="168"><%=getMsg(179,"Add")%></td>
      <td width="220">         
        <input type=text name=address value="<%=rstemp("address")%>">
      </td>
    </tr>
    <tr> 
      <td width="168"><%=getMsg(180,"Zip")%></td>
      <td width="220">         
        <input type=text name=zip value="<%=rstemp("zip")%>">
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
      <SELECT name=stateCode size=1>
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
       <input type=text name=state value="<%=rstemp("state")%>">
      </td>
    </tr>
<%end if%>

    <tr> 
      <td width="168"><%=getMsg(183,"City")%></td>
      <td width="220">         
        <input type=text name=city value="<%=rstemp("city")%>">
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
      <SELECT name=countryCode>
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
    
    <%if pCustomerFieldName1<>"" then%>
    <tr> 
      <td width="168"> <%=pCustomerFieldName1%> </td>
      <td width="220">         
       <input type=text name=user1 value="<%=rstemp("user1")%>">
      </td>
    </tr>
    <%end if%>
    
    <%if pCustomerFieldName2<>"" then%>
    <tr> 
      <td width="168"> <%=pCustomerFieldName2%> </td>
      <td width="220">         
       <input type=text name=user2 value="<%=rstemp("user2")%>">
      </td>
    </tr>
    <%end if%>
    
    <%if pCustomerFieldName3<>"" then%>
    <tr> 
      <td width="168"> <%=pCustomerFieldName3%> </td>
      <td width="220">         
       <input type=text name=user3 value="<%=rstemp("user3")%>">
      </td>
    </tr>
    <%end if%>
    
    <tr> 
      <td width="168">&nbsp;</td>
      <td width="220">&nbsp;</td>
    </tr>
    <tr> 
      <td colspan="2">      
      <input type=image src="images/buttonModify.gif" type=image alt="<%=getMsg(185,"Mdf")%>">                     
      </td>
      
    </tr>
    </table>
   </form>            
<!--#include file="footer.asp"-->
<%call closeDb()%>