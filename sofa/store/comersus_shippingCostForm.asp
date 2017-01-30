<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: get off line shipping quotes from cart view
%> 
<!--#include file="../includes/screenMessages.asp"--> 
<!--#include file="../includes/cartFunctions.asp"--> 
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/adSenseFunctions.asp"-->
<%

on error resume next

dim mySQL, connTemp, rstemp5

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanyLogo	 	= getSettingKey("pCompanyLogo")
pHeaderKeywords 	= getSettingKey("pHeaderKeywords")

pAuctions		= getSettingKey("pAuctions")
pAllowNewCustomer	= getSettingKey("pAllowNewCustomer")
pNewsLetter 		= getSettingKey("pNewsLetter")
pStoreNews 		= getSettingKey("pStoreNews")
pSuppliersList		= getSettingKey("pSuppliersList")
pDisableState  		= getSettingKey("pDisableState")

pIdDbSessionCart     	= getSessionVariable("idDbSessionCart",0)
pIdCustomer     	= getSessionVariable("idCustomer",0)

if pIdDbSessionCart=0 then 
  response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(424,"It seems that your session is lost. please try again from store home."))
end if

' if the customer is logged-in retrieve shipping address 
 
 if pIdCustomer<>0 then 
 
 mySQL="SELECT state, stateCode, zip, countryCode, shippingState, shippingStateCode, shippingZip, shippingCountryCode FROM customers WHERE idCustomer=" &pIdCustomer
 call getFromDatabase(mySQL, rsTemp5, "checkShippingAddress") 
 
 if not rsTemp5.eof then
    if rsTemp5("shippingZip")<>"" then
          ' load shipping
	  pZip		= rsTemp5("shippingZip")
	  pState	= rsTemp5("shippingState")
	  pStateCode	= rsTemp5("shippingStateCode") 
	  pCountryCode	= rsTemp5("shippingCountryCode")
    else
          ' load billing
	  pZip		= rsTemp5("zip")
	  pState	= rsTemp5("state")
	  pStateCode	= rsTemp5("stateCode") 
	  pCountryCode	= rsTemp5("countryCode")  	
    end if
 end if 
  
end if ' customer logged in

%>
          
<!--#include file="header.asp"--> 
<br>
<form METHOD="POST" name="shippingPriceForm" action="comersus_shippingCostExec.asp">
<b><%=getMsg(425,"s cost")%></b><br><br>
<TABLE BORDER="0" CELLPADDING="0">
    
    <%if pDisableState<>"-1" then%>
    <TR> 
      <TD WIDTH="202"><%=getMsg(414,"state")%></TD>
      <TD WIDTH="393">                 
        <%
      ' get stateCodes 
      mySQL="SELECT stateCode, stateName FROM stateCodes ORDER BY stateName"
      
      call getFromDatabase(mySQL, rstemp, "orderForm")      
      
      %>
      <SELECT name=stateCode size=1>
      <OPTION value=""><%=getMsg(415,"select state")%>
      <%do while not rstemp.eof
      pStateCode2=rstemp("stateCode")%>
        <option value="<%=pStateCode2%>"<%
        if (request("StateCode")="" and pStateCode2=pStateCode) Or (request("StateCode")=pStateCode2) then
           response.write "selected"
        end if%>><%=rstemp("stateName")%>
      <%rstemp.movenext
      loop%>
      </OPTION>
      </SELECT>
      </TD>
    </TR>
    
    <TR>  
      <TD><%=getMsg(416,"non listed")%></TD>
      <TD>
      <input type=text name=state value="<%=pState%>">
      </TD>
      </TD>
     </TR>
   <%end if%>
   
        
    <TR> 
      <TD WIDTH="202"><%=getMsg(417,"zip")%></TD>
      <TD WIDTH="393">           
       <input type=text name=zip value="<%=pZip%>">
         </TD>
    </TR>
    <TR> 
      <TD WIDTH="202"><%=getMsg(418,"ctry")%></TD>
      <TD WIDTH="393">        
      <%
      ' get CountryCodes
      mySQL="SELECT countryCode, countryName FROM countryCodes ORDER BY countryName"

      call getFromDatabase(mySQL, rstemp, "orderForm")
      
      %>
      <SELECT name="countryCode">
      <OPTION value=""><%=getMsg(419,"select ctry")%>
      <%do while not rstemp.eof
      pCountryCode2=rstemp("countryCode")%>
        <option value="<%=pCountryCode2%>"<%
        if (request("countryCode")="" and pCountryCode=pCountryCode2) Or (request("countryCode")=pCountryCode2) then
            response.write "selected"           
        end if
        %>><%=rstemp("countryName")%>
      <%rstemp.movenext
      loop%>
      </OPTION>
      </SELECT>
         </TD>
    </TR>    
    <TR> 
      <TD width="202">&nbsp;</TD>
      <TD width="393">&nbsp;</TD>
    </TR>
    <TR> 
      <TD width="202">&nbsp;</TD>
      <TD width="393">&nbsp;</TD>
    </TR>
    
    <TR> 
      <TD width="595" colspan="2"> 
        <input type="SUBMIT" name="Submit" value="<%=getMsg(426,"get")%>">        
      </TD>
    </TR>    
    <TR> 
      <TD width="202">&nbsp;</TD>
      <TD width="393">&nbsp;</TD>
    </TR>

  </table>
</form><br><br>
<!--#include file="footer.asp"-->
<%call closeDb()%>

