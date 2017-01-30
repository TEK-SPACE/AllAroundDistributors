<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: register a new customer
%>
<!--#include file="../includes/screenMessages.asp" --> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"-->
<!--#include file="../includes/miscFunctions.asp"-->
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/currencyFormat.asp"-->
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/cartFunctions.asp"-->
<!--#include file="../includes/adSenseFunctions.asp"-->

<% 
on error resume next

dim connTemp, mySql, rsTemp, rsTemp2

' get settings 
pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanySlogan 		= getSettingKey("pCompanySlogan")
pAboutUsLink 		= getSettingKey("pAboutUsLink")
pHeaderKeywords 	= getSettingKey("pHeaderKeywords")
pShowSearchBox 		= getSettingKey("pShowSearchBox")
pAdSenseClient 		= getSettingKey("pAdSenseClient")
pAuctions		= getSettingKey("pAuctions")
pAllowNewCustomer	= getSettingKey("pAllowNewCustomer")
pNewsLetter 		= getSettingKey("pNewsLetter")
pStoreNews 		= getSettingKey("pStoreNews")
pSuppliersList		= getSettingKey("pSuppliersList")
pRssFeedServer		= getSettingKey("pRssFeedServer")

pRandomPassword		= getSettingKey("pRandomPassword")
pDisableState		= getSettingKey("pDisableState")

' custom field names
pCustomerFieldName1	= getSettingKey("customerFieldName1")
pCustomerFieldName2	= getSettingKey("customerFieldName2")
pCustomerFieldName3	= getSettingKey("customerFieldName3")

pIdCustomer     	= getSessionVariable("idCustomer",0)
pIdCustomerType 	= getSessionVariable("idCustomerType",1)
pIdDbSessionCart	= getSessionVariable("idDbSessionCart",0)

%> 

<!--#include file="header.asp"-->

<br><br><b><%=getMsg(235,"Cust reg")%></b><br><br>
 <form method="post" name="cutR" action="comersus_customerRegistrationExec.asp">
  <table width="460" border="0">
    <tr> 
      <td width="40%"><%=getMsg(236,"name")%></td>
      <td width="60%">         
        <input type=text name=customerName>
      </td>
    </tr>    
    <tr> 
      <td width="40%"><%=getMsg(237,"last")%></td>
      <td width="60%">         
        <input type=text name=lastName>
      </td>
    </tr>
    <tr> 
      <td width="40%"><%=getMsg(238,"comp")%></td>
      <td width="60%">         
      <input type=text name=customerCompany>
      </td>
    </tr>
    <tr> 
      <td width="40%"><%=getMsg(239,"phone")%></td>
      <td width="60%">         
        <input type=text name=phone>
      </td>
    </tr>
    <tr> 
      <td width="40%"><%=getMsg(240,"email")%></td>
      <td width="60%">         
        <input type=text name=email>
      </td>
    </tr>
    <%if pRandomPassword="0" then%>
    <tr> 
      <td width="40%"><%=getMsg(241,"pwd")%></td>
      <td width="60%">         
        <input type=password name=password maxlength=15> (A-Z 0-9)
      </td>
    </tr>
    <%else%>
     <input type="hidden" name="password" value="<%=randomNumber(999999)%>">
    <%end if%>
    <tr> 
      <td width="40%"><%=getMsg(242,"addr")%></td>
      <td width="60%">         
        <input type=text name=address>
      </td>
    </tr>
    <tr> 
      <td width="40%"><%=getMsg(243,"zip")%></td>
      <td width="60%">         
        <input type=text name=zip>
      </td>
    </tr>

<%if pDisableState<>"-1" then%>        
    <tr> 
      <td width="40%"><%=getMsg(244,"state")%></td>
      <td width="60%">         
        <%      
      
      ' get stateCodes
      mySQL="SELECT stateCode, stateName FROM stateCodes ORDER BY stateName"
      
      call getFromDatabase(mySql, rstemp2,"customerRegistrationForm")      

      %>
      <SELECT name=stateCode size=1>
      <OPTION value=""><%=getMsg(245,"select")%>
      <%do while not rstemp2.eof
      pStateCode2=rstemp2("stateCode")%>
        <option value="<%=pStateCode2%>" <%if request.form("stateCode")=pStateCode2 then response.write " selected" %>><%=rstemp2("stateName")%>
      <%rstemp2.movenext
      loop%>
      </OPTION>
      </SELECT>
      </td>
    </tr>
    <tr> 
      <td width="40%"><%=getMsg(250,"non listed")%></td>
      <td width="60%"> 
      <input type=text name=state>
      </td>
    </tr>
<%end if%>    

    <tr> 
      <td width="40%"><%=getMsg(246,"city")%></td>
      <td width="60%">         
        <input type=text name=city>
      </td>
    </tr>
     <tr> 
      <td width="40%"><%=getMsg(247,"ctry")%></td>
      <td width="60%">         
        <%
      ' get CountryCodes
      mySQL="SELECT countryCode, countryName FROM countryCodes ORDER BY countryName"
    
      call getFromDatabase(mySql, rstemp2,"customerRegistrationForm")            
      
      %>      
      <SELECT name=CountryCode>
      <OPTION value=""><%=getMsg(248,"select")%>
      <%do while not rstemp2.eof
      pCountryCode2=rstemp2("countryCode")%>
        <option value="<%=pCountryCode2%>" <%if trim(request.form("countryCode"))=pCountryCode2 then response.write " selected" %><%
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
      <td width="40%"> <%=pCustomerFieldName1%> </td>
      <td width="60%">         
        <input type=text name=user1>
      </td>
    </tr>
    <%end if%>
    
    <%if pCustomerFieldName2<>"" then%>
    <tr> 
      <td width="40%"> <%=pCustomerFieldName2%> </td>
      <td width="60%">         
        <input type=text name=user2>
      </td>
    </tr>
    <%end if%>
    
    <%if pCustomerFieldName3<>"" then%>
    <tr> 
      <td width="40%"> <%=pCustomerFieldName3%> </td>
      <td width="60%">         
        <input type=text name=user3>
      </td>
    </tr>
    <%end if%>                  
    <tr> 
      <td width="40%">&nbsp;</td>
      <td width="60%">&nbsp;</td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <input type="submit" name="save" value="<%=getMsg(249,"register")%>">
        <br>
        <br>
      </td>
    </tr>
  </table>
 </form> 
<!--#include file="footer.asp"-->
<%call closeDb()%>