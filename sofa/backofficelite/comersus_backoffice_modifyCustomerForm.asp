<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"-->   
<!--#include file="../includes/settings.asp" --> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/encryption.asp"--> 
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../store/comersus_optDes.asp"--> 

<%
on error resume next

dim conntemp, mysql, rstemp, rstemp2

' get settings 
pEncryptionPassword 	= getSettingKey("pEncryptionPassword")
pEncryptionMethod 	= getSettingKey("pEncryptionMethod")

pIdCustomer  		= getUserInput(request("idCustomer"),20)

mySql="SELECT password, active, idCustomerType, wapAccess, name, lastName, customerCompany, phone, email, address, zip, stateCode, state, city, countryCode, bonusPoints FROM customers WHERE idCustomer="&pIdCustomer

call getFromDatabase(mySQL, rstemp, "comersus_backoffice_modifyCustomerForm.asp") 

pPassword = DeCrypt(rstemp("password"), pEncryptionPassword) 
%>

<!--#include file="header.asp"-->
<!--#include file="./includes/settings.asp"-->

<br><b>Modify customer record</b><br>
<form method="post" name="modCust" action="comersus_backoffice_modifyCustomerExec.asp">
<table width="600" border="0">
  <tr> 
    <td width="100">Name</td>
    <td width="200"> 
      <input type="hidden" name="idCustomer" value="<%=pIdcustomer%>">        
      <input type=text name="name" value="<%=rstemp("name")%>"> 
    </td>
  </tr>    
  <tr> 
    <td width="100">Last name</td>
    <td width="200">         
       <input type=text name="lastName" value="<%=rstemp("lastName")%>">       
    </td>
  </tr>
  <tr> 
    <td width="100">Company</td>
    <td width="200">         
       <input type=text name="customerCompany" value="<%=rstemp("customerCompany")%>">       
    </td>
  </tr>
  <tr> 
    <td width="100">Phone</td>
    <td width="200">        
      <input type=text name="phone" value="<%=rstemp("phone")%>">       
    </td>
  </tr>
  <tr> 
    <td width="100">Email</td>
    <td width="200">    
      <input type=text name="email" value="<%=rstemp("email")%>">           
    </td>
  </tr>
  <tr> 
    <td width="100">Password</td>
    <td width="200">    
      <input type=password name="password" value="<%=pPassword%>">            
    </td>
  </tr>
  <tr> 
    <td width="100">Address</td>
    <td width="200">      
       <input type=text name="address" value="<%=rstemp("address")%>">         
    </td>
  </tr>
  <tr> 
    <td width="100">Zip</td>
    <td width="200">    
      <input type=text name="zip" value="<%=rstemp("zip")%>">          
    </td>
  </tr>
  <tr> 
    <td width="100">State</td>
    <td width="200">                 
      <%
    ' get stateCodes
    mySQL="SELECT * FROM stateCodes"
    call getFromDatabase(mySQL, rsTemp2, "comersus_backoffice_modifyCustomerForm.asp") 
    %>
    <select name=stateCode size=1>
    <option value="">Select a State
    <%do while not rstemp2.eof
    pStateCode  = rstemp("stateCode")
    pStateCode2 = rstemp2("stateCode")%>
      <option value="<%=pStateCode2%>"<%
      if pStateCode2=pStateCode then
         response.write "selected"
      end if%>><%=rstemp2("stateName")%>
    <%rstemp2.movenext
    loop%>
    </option>
    </select>
    </td>
  </tr>
  <tr> 
    <td width="100">City</td>
    <td width="200">         
       <input type=text name="city" value="<%=rstemp("city")%>">       
    </td>
  </tr>
  <tr> 
    <td width="100">Non listed state</td>
    <td width="200">         
      <input type=text name="state" value="<%=rstemp("state")%>">       
    </td>
  </tr>
  <tr> 
    <td width="100">Country</td>
    <td width="200">  
           
    <%
    ' get CountryCodes
    mySQL="SELECT * FROM countryCodes"
    call getFromDatabase(mySQL, rsTemp2, "comersus_backoffice_modifyCustomerForm.asp") 
     %>
    <select name=countryCode>
    <option value="">Country
    <%do while not rstemp2.eof
    pCountryCode  = rstemp("countryCode")
    pCountryCode2 = rstemp2("countryCode")%>
      <option value="<%=pCountryCode2%>"<%
      if pCountryCode=pCountryCode2 then
          response.write "selected"           
      end if
      %>><%=rstemp2("countryName")%>
    <%rstemp2.movenext
    loop%>
    </option>
    </select>
    </td>
  </tr>
  <tr> 
    <td width="100">Bonus Points</td>
    <td width="200">
        <%pBonusPoints=money(rstemp("bonusPoints"))%>
           <input type=text name="bonusPoints" value="<%=pBonusPoints%>">       
    </td>
  </tr>
</table>
<table width="500" border="0">
  <tr> 
    <td width="100">Active</td>
    <td width="200">         
       <select name="active">          
        <option value='0' 
        <%if rstemp("active")=0 then 
         response.write "selected"
         end if%>
        >No</option>
        <option value='-1'
        <%if rstemp("active")=-1 then 
         response.write "selected"
        end if%>
        >Yes</option>
       </select>         
    </td>
    <td width="100">Type</td>
    <td width="200">         
       <select name="idCustomerType">          
        <option value='1' 
        <%if rstemp("idCustomerType")=1 then 
         response.write "selected"
        end if%>
        >Retail</option>
        <option value='2'
        <%if rstemp("idCustomerType")=2 then 
         response.write "selected"
        end if%>
        >Wholesale</option>
       </select>
    </td>
    <td width="180">WAP Access</td>
    <td width="150">         
       <select name="wapAccess">          
        <option value='0' 
        <%if rstemp("wapAccess")=0 then 
         response.write "selected"
        end if%>
        >No</option>
        <option value='-1'
        <%if rstemp("wapAccess")=-1 then 
         response.write "selected"
        end if%>
        >Yes</option>
       </select>
    </td>
  </tr>
  <tr> 
    <td width="100">&nbsp;</td>
    <td width="200">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="6">       
        <input type="submit" name="Modify" value="Modify">
        <input type="submit" name="Delete" value="Delete">        
    </td>
  </tr>
</table> 
<%call closeDb()%>      
<!--#include file="footer.asp"-->
</form>  