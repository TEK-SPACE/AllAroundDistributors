<%
' Comersus 6.0x Sophisticated Cart 
' Developed by Rodrigo S. Alhadeff for Comersus Open Technologies
' United States 
' Open Source License can be found at License.txt
' http://www.comersus.com  
' Details: if cart is not empty get orderform screen for first buy
%>

<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/screenMessages.asp"-->
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/cartFunctions.asp"--> 
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/encryption.asp" --> 
<!--#include file="../includes/sessionFunctions.asp"-->
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/adSenseFunctions.asp"-->


<%
on error resume next

dim mySQL, connTemp, rsTemp

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanyLogo	 	= getSettingKey("pCompanyLogo")

pAuctions		= getSettingKey("pAuctions")
pListBestSellers 	= getSettingKey("pListBestSellers")
pNewsLetter 		= getSettingKey("pNewsLetter")
pPriceList 		= getSettingKey("pPriceList")
pStoreNews 		= getSettingKey("pStoreNews")
pOneStepCheckout	= getSettingKey("pOneStepCheckout")
pUseShippingAddress	= getSettingKey("pUseShippingAddress")
pRandomPassword		= getSettingKey("pRandomPassword")
pUseVatNumber		= getSettingKey("pUseVatNumber")
pEncryptionMethod 	= getSettingKey("pEncryptionMethod")
pBonusPointsPerPrice	= getSettingKey("pBonusPointsPerPrice")
pDisableState		= getSettingKey("pDisableState")
pByPassShipping		= getSettingKey("pByPassShipping")
pUseShippingAddress	= getSettingKey("pUseShippingAddress")
pTeleSignCustomerId	= getSettingKey("pTeleSignCustomerId")
pAllowNewCustomer	= getSettingKey("pAllowNewCustomer")
pHeaderKeywords 	= getSettingKey("pHeaderKeywords")

' custom field names
pCustomerFieldName1	= getSettingKey("customerFieldName1")
pCustomerFieldName2	= getSettingKey("customerFieldName2")
pCustomerFieldName3	= getSettingKey("customerFieldName3")

' session
pWishListIdCustomer    	= getSessionVariable("wishListIdCustomer",0)
pIdDbSessionCart     	= getSessionVariable("idDbSessionCart",0)

pIdDbSession	 	= checkSessionData()
pIdDbSessionCart 	= checkDbSessionCartOpen()

session("idCustomer") 	= Cint(0)

if countCartRows(pIdDbSessionCart)=0 then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(616,"It seems that your session was lost. Please try again from store home and make sure that you have Cookies enabled."))
end if

if pIdDbSessionCart=0 then 
 response.redirect "comersus_supportError.asp?error="&Server.Urlencode(getMsg(617,"Database session cart lost, please try again from store home."))
end if

if pAllowNewCustomer="0" or (pTeleSignCustomerId<>"0" and pTeleSignCustomerId<>"") then
 response.redirect "comersus_message.asp?message="&Server.Urlencode("Feature disabled in this store")
end if

%>

<!--#include file="header.asp"--> 
<br><b><%=getMsg(618,"personal info")%></b><br>
<form method="POST" name="orderform" action="comersus_checkOutCustomerExec.asp">
  <TABLE BORDER="0" CELLPADDING="0" WIDTH="100%" align="left">
    <TR> 
      <TD COLSPAN="2"><b><%=getMsg(619,"customer")%></b></TD>
    </TR>    
    <TR> 
      <TD><%=getMsg(173,"name")%></TD>
      <TD>          
        <input type=text name=name>        
         </TD>
    </TR>
     <TR> 
      <TD><%=getMsg(174,"lName")%></TD>
      <TD>  
        <input type=text name=lastName>        
         </TD>
    </TR>   
     <TR> 
      <TD><%=getMsg(175,"company")%></TD>
      <TD>  
       <input type=text name=customerCompany>        
         </TD>
    </TR>   

    <TR> 
      <TD><%=getMsg(176,"")%></TD>
      <TD>  
        <input type=text name=phone>        
         </TD>
    </TR>
    <TR> 
      <TD><%=getMsg(177,"email")%></TD>
      <TD>        
        <input type=text name=email>        
         </TD>
    </TR>    

<%if pRandomPassword="0" then%>    
    <TR> 
      <TD><%=getMsg(178,"")%></TD>      
       <TD><input type=password name=password maxlength=15> (A-Z 0-9)</TD>      
    </TR>
    <TR> 
      <TD><%=getMsg(621,"retype")%></TD>      
       <TD><input type=password name=password2 maxlength=15></TD>      
    </TR>        
<%end if%>    

    <%if pCustomerFieldName1<>"" then%>
    <tr> 
      <td width="168"> <%=pCustomerFieldName1%> </td>
      <td width="220">         
       <input type=text name=user1>              
      </td>
    </tr>
    <%end if%>
    
    <%if pCustomerFieldName2<>"" then%>
    <tr> 
      <td width="168"> <%=pCustomerFieldName2%> </td>
      <td width="220">         
       <input type=text name=user2>        
      </td>
    </tr>
    <%end if%>
    
    <%if pCustomerFieldName3<>"" then%>
    <tr> 
      <td width="168"> <%=pCustomerFieldName3%> </td>
      <td width="220">         
       <input type=text name=user3>        
      </td>
    </tr>
    <%end if%>
               
    <TR> 
      <TD WIDTH="595" colspan="2">
      <br><b><%=getMsg(622,"billing")%></b></TD>
    </TR>
    <TD><%=getMsg(179,"addr")%></TD>
    <TD>   
      <input type=text name=address>        
       </TD>
    </TR>
    <TR> 
      <TD WIDTH="202"HEIGHT="34"><%=getMsg(183,"city")%></TD>
      <TD WIDTH="393"HEIGHT="34">           
        <input type=text name=city>        
         </TD>
    </TR>
    
<%if pDisableState<>"-1" then%>    
    <TR> 
      <TD><%=getMsg(181,"state")%></TD>
      <TD>                 
        <%
      ' get stateCodes
      mySQL="SELECT stateCode, stateName FROM stateCodes ORDER BY stateName"
      
      call getFromDatabase(mySQL, rstemp, "orderForm")      
      
      %>
      <SELECT name=stateCode size=1>
      <OPTION value=""><%=getMsg(186,"select")%>
      <%do while not rstemp.eof%>
        <option value="<%=rstemp("stateCode")%>"><%=rstemp("stateName")%>
      <%rstemp.movenext
      loop%>
      </OPTION>
      </SELECT>
      </TD>
    </TR>
    <TR>  
    <TD><%=getMsg(182,"non listed")%></TD>
      <TD>
      <input type=text name=state>                         
      </TD>
      </TD>
     </TR>
<%end if%>    

    <TR> 
      <TD><%=getMsg(180,"zip")%></TD>
      <TD>           
        <input type=text name=zip>        
         </TD>
    </TR>
    <TR> 
      <TD><%=getMsg(184,"ctry")%></TD>
      <TD>        
      <%
      ' get CountryCodes
      mySQL="SELECT countryCode, countryName FROM countryCodes ORDER BY countryName"

      call getFromDatabase(mySQL, rstemp, "orderForm")
      
      %>
      <SELECT name="countryCode">
      <OPTION value=""><%=getMsg(187,"select")%>
      <%do while not rstemp.eof%>
        <option value="<%=rstemp("countryCode")%>"><%=rstemp("countryName")%>
      <%rstemp.movenext
      loop%>
      </OPTION>
      </SELECT>
         </TD>
    </TR>

<%if pByPassShipping="0" and pUseShippingAddress="-1" then%>        
    <%
    
    if pWishListIdCustomer<>0 then
    
      ' get address of the owner of the wish list
      mySQL="SELECT address, zip, state, stateCode, city, countryCode, shippingAddress, shippingZip, shippingStateCode, shippingState, shippingCity, shippingCountryCode FROM customers WHERE idCustomer="&pWishListIdCustomer
      call getFromDatabase(mySQL, rstemp, "checkOutCustomerForm")
      
      if not rstemp.eof then
        pAddress		= rstemp("address")
	pZip			= rstemp("zip")
	pStateCode		= rstemp("stateCode")
	pState			= rstemp("state")
	pCity			= rstemp("city")
	pCountryCode		= rstemp("countryCode")
	
	pShippingAddress	= rstemp("shippingAddress")
	pShippingZip		= rstemp("shippingZip")
	pShippingStateCode	= rstemp("shippingStateCode")
	pShippingState		= rstemp("shippingState")
	pShippingCity		= rstemp("shippingCity")
	pShippingCountryCode	= rstemp("shippingCountryCode")
	
	' customer doesnt have shipping address, use billing
	if pShippingAddress="" then
	 pShippingAddress	= pAddress
	 pShippingZip		= pZip
	 pShippingStateCode	= pStateCode
	 pShippingState		= pShippingState
	 pShippingCity		= pShippingCity
	 pShippingCountryCode	= pShippingCountryCode
	end if
	
      end if
      
    end if
    
    %>    
    <tr> 
      <td colspan="2"><br><b><%=getMsg(471,"ship addr")%></b> </td>
    </tr>
    
    <tr> 
      <td><%=getMsg(179,"addr")%></td>
      <td>     
      <%if pWishListIdCustomer=0 then%>   
        <input type=text name=shippingAddress>        
      <%else%>
        <input type="hidden" name="shippingAddress" value="<%=pShippingAddress%>">                
        <%=pShippingAddress%>
      <%end if%>
         </td>
    </tr>
    <tr> 
      <td><%=getMsg(183,"city")%></td>
      <td> 
      <%if pWishListIdCustomer=0 then%>   
        <input type=text name=shippingCity>        
         <%else%>
        <input type="hidden" name="shippingCity" value="<%=pShippingCity%>">                
        <%=pShippingCity%>
      <%end if%>
         </td>
    </tr>

<%if pDisableState<>"-1" then%>        
    <tr> 
      <td><%=getMsg(181,"state")%></td>
      <td>
      <%if pWishListIdCustomer=0 then%>   
          <%
      ' get stateCodes
      mySQL="SELECT stateCode, stateName FROM stateCodes ORDER BY stateName"
      
      call getFromDatabase(mySQL, rstemp, "checkShippingAddress") 
            
      %>
      <SELECT name="shippingStateCode" size=1>
      <OPTION value=""><%=getMsg(186,"select")%>
      <%do while not rstemp.eof
      pStateCode2=rstemp("stateCode")
      %>
        <option value="<%=pStateCode2%>"
        <%if pStateCode2=pShippingStateCode then
           response.write "selected"
        end if%>
        ><%=rstemp("stateName")%>
      <%rstemp.movenext
      loop%>
      </OPTION>
      </SELECT>
      <%else%>            
        <input type="hidden" name="shippingStateCode" value="<%=pShippingStateCode%>">                
        <input type="hidden" name="shippingState" value="<%=pShippingState%>">                
        <%=pShippingStateCode%><%=pShippingState%>
      <%end if%>
      
        </td>
    </tr>
    
    <%if pWishListIdCustomer=0 then%>   
    <tr> 
      <td><%=getMsg(182,"non listed")%></td>
      <td><input type=text name=shippingState></td>
    </tr>
    <%end if%>
<%end if%>    

    <tr> 
      <td><%=getMsg(180,"zip")%></td>
      <td>
      <%if pWishListIdCustomer=0 then%>   
        <input type=text name=shippingZip>                            
      <%else%>
        <input type="hidden" name="shippingZip" value="<%=pShippingZip%>">                
        <%=pShippingZip%>
      <%end if%>
        </td>
    </tr>
    <tr> 
      <td><%=getMsg(184,"country")%></td>
      <td>                 
      <%if pWishListIdCustomer=0 then%>   
      <%
      ' get CountryCodes
      mySQL="SELECT countryCode, countryName FROM countryCodes ORDER BY countryName"
      
      call getFromDatabase(mySQL, rstemp, "checkShippingAddress") 
            
      %>
      <SELECT name="shippingCountryCode" size=1>
      <OPTION value=""><%=getMsg(187,"select")%>
      <%do while not rstemp.eof
      pShippingCountryCode2=rstemp("countryCode")%>
        <option value="<%=pShippingCountryCode2%>"
        <%if pShippingCountryCode2=pShippingCountryCode then
        response.write "selected"
        end if%>><%=rstemp("countryName")%>
      <%rstemp.movenext
      loop%>
      </OPTION>
      </SELECT>
      <%else%>
      <input type="hidden" name="shippingCountryCode" value="<%=pShippingCountryCode%>">                
        <%=pShippingCountryCode%>
      <%end if%>
         </td>
    </tr>        
<%end if ' use shipping address%>    
    <TR> 
      <TD>&nbsp;</TD>
      <TD>&nbsp;</TD>
    </TR>
    <TR> 
      <TD width="595" colspan="2"> 
        <input type="SUBMIT" name="Submit" value="<%=getMsg(421,"continue")%>">        
      </TD>
    </TR>
    <TR> 
      <TD>&nbsp;</TD>
      <TD>&nbsp;</TD>
    </TR>
  </table>
</form>
 <br>
<!--#include file="footer.asp"-->
<%call closeDb()%>
