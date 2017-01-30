<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: autenticate customer using email+password and get user data 
' If the customer is already logged in, just get customer data from database
%>

<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/screenMessages.asp"--> 
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/currencyFormat.asp" --> 
<!--#include file="../includes/cartFunctions.asp"--> 
<!--#include file="../includes/encryption.asp" -->
<!--#include file="../includes/sessionFunctions.asp"-->
<!--#include file="../includes/itemFunctions.asp"-->

<%
on error resume next

dim rsTemp, connTemp, mySql, pEmail, pPassword

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanyLogo	 	= getSettingKey("pCompanyLogo")
pCustomerPrefix		= getSettingKey("pCustomerPrefix")
pUseShippingAddress	= getSettingKey("pUseShippingAddress")
pRandomPassword		= getSettingKey("pRandomPassword")
pUseVatNumber		= getSettingKey("pUseVatNumber")
pEncryptionPassword 	= getSettingKey("pEncryptionPassword")
pEncryptionMethod 	= getSettingKey("pEncryptionMethod")
pBonusPointsPerPrice	= getSettingKey("pBonusPointsPerPrice")
pDisableState		= getSettingKey("pDisableState")

' custom field names
pCustomerFieldName1	= getSettingKey("customerFieldName1")
pCustomerFieldName2	= getSettingKey("customerFieldName2")
pCustomerFieldName3	= getSettingKey("customerFieldName3")

pAuctions		= getSettingKey("pAuctions")
pListBestSellers 	= getSettingKey("pListBestSellers")
pNewsLetter 		= getSettingKey("pNewsLetter")
pPriceList 		= getSettingKey("pPriceList")
pStoreNews 		= getSettingKey("pStoreNews")
pUseShippingAddress	= getSettingKey("pUseShippingAddress")
pByPassShipping		= getSettingKey("pByPassShipping")

pIdDbSession	 	= checkSessionData()
pIdDbSessionCart 	= checkDbSessionCartOpen()

pEmail 			= getUserInput(request.form("email"),50)
pPassword 		= getUserInput(request.form("password"),50)

pIdCustomer     	= getSessionVariable("idCustomer",0)
pWishListIdCustomer	= getSessionVariable("wishListIdCustomer",0)

if pIdCustomer<>0 then

 ' customer already logged in

 mySQL="SELECT name, lastName, customerCompany, phone, email, address, zip, state, stateCode, city, countryCode, shippingAddress, shippingZip, shippingStateCode, shippingState, shippingCity, shippingCountryCode, user1, user2, user3, bonusPoints FROM customers WHERE idCustomer=" &pIdCustomer
 call getFromDatabase(mySQL, rstemp, "login")		

 if rstemp.eof then 
  response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(402,"Cannot get customer details.")) 
 end if
	
else  
 
 ' verify password for that email	

 mySQL="SELECT idCustomer, name, lastName, customerCompany, phone, email, address, zip, state, stateCode, city, countryCode, shippingAddress, shippingZip, shippingStateCode, shippingState, shippingCity, shippingCountryCode, user1, user2, user3 FROM customers WHERE email='" &pEmail& "' AND password='" &EnCrypt(pPassword, pEncryptionPassword)& "' AND active=-1"		
 call getFromDatabase(mySQL, rstemp, "login")				

 if rstemp.eof then 
  response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(403,"Invalid login information.")) 
 end if
 
 ' save logged customer in session
 session("idCustomer") = rstemp("idCustomer")  
end if

pName			= rstemp("name")
pLastName		= rstemp("lastName")
pCustomerCompany	= rstemp("customerCompany")
pPhone			= rstemp("phone")
pEmail			= rstemp("email")

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

pUser1			= rstemp("user1")
pUser2			= rstemp("user2")
pUser3			= rstemp("user3")

pBonusPoints		= rstemp("bonusPoints")

session("customerName")	= pName

pIdDbSession	 	= checkSessionData()
pIdDbSessionCart 	= checkDbSessionCartOpen()
%>

<!--#include file="header.asp"--> 
<br><b><%=getMsg(404,"checkout")%></b><br>
<form METHOD="POST" name="orderform" action="comersus_checkOutModifyCustomerExec.asp">
  <TABLE BORDER="0" CELLPADDING="0" WIDTH="500" align="left">
    <TR> 
      <TD COLSPAN="2"> <b><%=getMsg(405,"cust infot")%></b> </TD>
    </TR>    
    <TR> 
      <TD><%=getMsg(406,"name")%></TD>
      <TD>          
        <input type=text name="name" value="<%=pName%>">     
         </TD>
    </TR>
     <TR> 
      <TD><%=getMsg(407,"l name")%></TD>
      <TD>  
        <input type=text name="lastName" value="<%=pLastName%>">     
         </TD>
    </TR>   
     <TR> 
      <TD><%=getMsg(408,"company")%></TD>
      <TD>  
       <input type=text name="customerCompany" value="<%=pCustomerCompany%>">     
         </TD>
    </TR>   

    <TR> 
      <TD><%=getMsg(409,"company")%></TD>
      <TD>  
        <input type=text name="phone" value="<%=pPhone%>">     
         </TD>
    </TR>
    <TR> 
      <TD><%=getMsg(410,"email")%></TD>
      <TD> <%=pEmail%>
         </TD>
    </TR>                    
    
    <%if pCustomerFieldName1<>"" then%>
    <tr> 
      <td width="168"> <%=pCustomerFieldName1%> </td>
      <td width="220">         
       <input type=text name="user1" value="<%=pUser1%>">     
      </td>
    </tr>
    <%end if%>
    
    <%if pCustomerFieldName2<>"" then%>
    <tr> 
      <td width="168"> <%=pCustomerFieldName2%> </td>
      <td width="220">         
       <input type=text name="user2" value="<%=pUser2%>">     
      </td>
    </tr>
    <%end if%>
    
    <%if pCustomerFieldName3<>"" then%>
    <tr> 
      <td width="168"> <%=pCustomerFieldName3%> </td>
      <td width="220">         
       <input type=text name="user3" value="<%=pUser3%>">     
      </td>
    </tr>
    <%end if%>
               
    <TR> 
      <TD WIDTH="595" colspan="2">
      <br><b><%=getMsg(411,"billing info")%></b></TD>
    </TR>
    <TD><%=getMsg(412,"addre")%></TD>
    <TD>   
      <input type=text name="address" value="<%=pAddress%>">     
       </TD>
    </TR>
    <TR> 
      <TD><%=getMsg(413,"city")%></TD>
      <TD>           
        <input type=text name="city" value="<%=pCity%>">     
         </TD>
    </TR>

<%if pDisableState<>"-1" then%>        
    <TR> 
      <TD><%=getMsg(414,"state")%></TD>
      <TD>                 
        <%
      ' get stateCodes
      mySQL="SELECT stateCode, stateName FROM stateCodes ORDER BY stateName"
      
      call getFromDatabase(mySQL, rstemp, "orderForm")      
      
      %>
      <SELECT name=stateCode size=1>
      <OPTION value=""><%=getMsg(415,"select")%>
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
      <input type=text name="state" value="<%=pState%>">     
      </TD>
      </TD>
     </TR>
<%end if%>    

    <TR> 
      <TD><%=getMsg(417,"zip")%></TD>
      <TD>           
        <input type=text name="zip" value="<%=pZip%>">   
         </TD>
    </TR>
    <TR> 
      <TD><%=getMsg(418,"ctry")%></TD>
      <TD>           
      <%
      ' get Country
      mySQL="SELECT countryCode, countryName FROM countryCodes ORDER BY countryName"

      call getFromDatabase(mySQL, rstemp, "orderForm")
      
      %>
      <SELECT name="countryCode">
      <OPTION value=""><%=getMsg(419,"select")%>
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
    
<%if pByPassShipping="0" and pUseShippingAddress="-1" then

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
      <td colspan="2"><br><b><%=getMsg(420,"ship info")%></b> </td>
    </tr>
    
    <tr> 
      <td><%=getMsg(412,"addr")%> </td>
      <td>     
      <%if pWishListIdCustomer=0 then%>   
        <input type=text name="shippingAddress" value="<%=pShippingAddress%>">   
      <%else%>
        <input type="hidden" name="shippingAddress" value="<%=pShippingAddress%>">                
        <%=pShippingAddress%>
      <%end if%>
         </td>
    </tr>
    <tr> 
      <td> <%=getMsg(413,"city")%></td>
      <td> 
      <%if pWishListIdCustomer=0 then%>   
        <input type=text name="shippingCity" value="<%=pShippingCity%>">   
         <%else%>
        <input type="hidden" name="shippingCity" value="<%=pShippingCity%>">                
        <%=pShippingCity%>
      <%end if%>
         </td>
    </tr>

<%if pDisableState<>"-1" then%>        
    <tr> 
      <td><%=getMsg(414,"state")%></td>
      <td>
      <%if pWishListIdCustomer=0 then%>   
          <%
      ' get stateCodes
      mySQL="SELECT stateCode, stateName FROM stateCodes ORDER BY stateName"
      
      call getFromDatabase(mySQL, rstemp, "checkShippingAddress") 
            
      %>
      <SELECT name="shippingStateCode" size=1>
      <OPTION value=""><%=getMsg(415,"select")%>
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
      <td><%=getMsg(416,"non listed")%></td>
      <td><input type=text name="shippingState" value="<%=pShippingState%>">   </td>
    </tr>
    <%end if%>
<%end if%>    
    <tr> 
      <td><%=getMsg(417,"zip")%></td>
      <td>
      <%if pWishListIdCustomer=0 then%>   
        <input type=text name="shippingZip" value="<%=pShippingZip%>">   
      <%else%>
        <input type="hidden" name="shippingZip" value="<%=pShippingZip%>">                
        <%=pShippingZip%>
      <%end if%>
        </td>
    </tr>
    <tr> 
      <td><%=getMsg(418,"ctry")%></td>
      <td>                 
      <%if pWishListIdCustomer=0 then%>   
      <%
      ' get country
      mySQL="SELECT countryCode, countryName FROM countryCodes ORDER BY countryName"
      
      call getFromDatabase(mySQL, rstemp, "checkShippingAddress") 
            
      %>
      <SELECT name="shippingCountryCode" size=1>
      <OPTION value=""><%=getMsg(419,"select")%>
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
<%end if%>    
    <TR> 
      <TD>&nbsp;</TD>
      <TD>&nbsp;</TD>
    </TR>
    <TR> 
      <TD colspan="2"> 
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