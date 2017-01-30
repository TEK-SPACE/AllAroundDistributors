<%
' Comersus BackOffice 
' e-commerce ASP Open Source
' CopyRight Comersus Open Technologies LC
' United States - 2005
' http://www.comersus.com 
%>
<!--#include file="comersus_backoffice_adminVerify.asp"-->  
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/sessionFunctions.asp"-->    
<!--#include file="../includes/getSettingKey.asp"-->  
<!--#include file="../includes/currencyFormat.asp"-->    
<!--#include file="../includes/settings.asp" --> 
<!--#include file="includes/validateForm.asp" --> 


<%
on error resume next

dim  mySQL, conntemp, rstemp

' get settings 
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")


pIdStoreBackOffice	= getSessionVariable("idStoreBackOffice", 1)

validateForm "comersus_backoffice_addTaxPerPlaceExec.asp" 
%> 

<!--#include file="header.asp"-->
<!--#include file="includes/settings.asp"-->
<b><div class=title>Add Tax Rule</div></b>
<br><br>
<form method="post" name="addSh">
<% validateError %>
  <table width="533" border="0">
    
     <tr> 
      <td width="168">Tax</td>
      <td colspan="2"> %         
        <%validate "taxPerPlace", "number"%> <%textbox "taxPerPlace", "0.00", 6, "textbox"%> (use - to discount)       
      </td>
    </tr>        
    <tr> 
      <td width="168">Zip Code
         </td>
      <td colspan="2">
        <%textbox "zip", "", 12, "textbox"%> 
      </td>
    </tr>
    <tr> 
      <td width="220">Zip Equal</td>
      <td width="684" colspan="2">Yes  
        <input type="checkbox" name="zipEq" value="-1">
      </td>
    </tr>
    <tr> 
    <td width="168">State Code</td>
      <td width="220">         
        <%
    
      ' get stateCodes
      mySQL="SELECT * FROM stateCodes"
      call getFromDatabase(mySQL, rstemp2, "comersus_backoffice_addTaxPerPlaceForm.asp") 


      %>
      <SELECT name=stateCode size=1>
      <OPTION value="">State Code
      <%do while not rstemp2.eof
      pStateCode2=rstemp2("stateCode")%>
        <option value="<%response.write pStateCode2%>"><%response.write rstemp2("stateName")%>
      <%rstemp2.movenext
      loop%>
      </OPTION>
      </SELECT>
      </td>
    </tr>
    <tr> 
      <td width="220">State Code Equal</td>
      <td width="684" colspan="2">Yes  
        <input type="checkbox" name="stateCodeEq" value="-1">
      </td>
    </tr>
 
     <tr> 
      <td width="168">Country</td>
      <td width="220">                 
        <%
      ' get CountryCodes
      mySQL="SELECT * FROM countryCodes ORDER BY countryName"
      call getFromDatabase(mySQL, rstemp2, "comersus_backoffice_addTaxPerPlaceForm.asp") 
      
      %>
      <SELECT name=CountryCode>
      <OPTION value="">Country
      <%do while not rstemp2.eof
      pCountryCode2=rstemp2("countryCode")%>
        <option value="<%=pCountryCode2%>"<%
        if pCountryCode=pCountryCode2 then
            response.write "selected"           
        end if
        %>><%response.write rstemp2("countryName")%>
      <%rstemp2.movenext
      loop%>
      </OPTION>
      </SELECT>       
      </td>
    </tr>
    <tr> 
      <td width="220">Country Code Equal</td>
      <td width="684" colspan="2">Yes  
        <input type="checkbox" name="countryCodeEq" value="-1" checked>
      </td>
    </tr> 
    
     <tr> 
      <td width="168">Customer Type</td>
      <td width="220">                 
        <%      
      mySQL="SELECT * FROM customerTypes"
      call getFromDatabase(mySQL, rstemp2, "comersus_backoffice_addTaxPerPlaceForm.asp")       
      %>
      <SELECT name=idCustomerType>
      <%do while not rstemp2.eof%>      
        <option value="<%=rsTemp2("idCustomerType")%>"><%=rstemp2("customerTypeDesc")%>
      <%rstemp2.movenext
      loop%>
      </OPTION>
      </SELECT>       
      </td>
    </tr>
    <tr> 
      <td colspan="3"> 
        &nbsp;
      </td>
    </tr>
    <tr> 
      <td colspan="3"> 
        <input type="submit" name="Submit" value="Save">
      </td>
    </tr>
  </table>
</form>

<%
' get taxes per place

mySQL="SELECT * FROM taxPerPlace WHERE idStore="&pIdStoreBackOffice
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_listTaxPerPlace.asp") 

if  rstemp.eof then 
 response.redirect "comersus_backoffice_message.asp?message="&Server.Urlencode("No Taxes Per Place available.")
end if
%>

<b><font size=2>Edit Tax Rule</font></b>
<br><br>
<table>
<tr bgcolor="#CCCCCC">
<td>Tax</td>
<td>Zip</td>
<td>State</td>
<td>Country</td>
<td>Customer Type</td>
<td>Actions</td>
</tr>
<%

if rstemp.eof then
%><tr><td colspan=6>No rules defined yet</td></tr><%
end if

 do while not rstemp.eof 

   pIdTaxPerPlace	= rstemp("idTaxPerPlace")
   pTaxPerPlace		= rstemp("taxPerPlace")
   pCountryCode		= rstemp("countryCode")
   pCountryCodeEq	= rstemp("countryCodeEq")
   pStateCode		= rstemp("stateCode")
   pStateCodeEq		= rstemp("stateCodeEq")
   pZip			= rstemp("zip")      
   pZipEq		= rstemp("zipEq")      
   pIdCustomerType	= rstemp("idCustomerType")      
   
   if isNull(pIdCustomerType) or pIdCustomerType=0 then
    pCustomerTypeDesc="All"    
   else
    mySQL="SELECT customerTypeDesc FROM customerTypes WHERE idCustomerType="&pIdCustomerType
    call getFromDatabase(mySQL, rstemp2, "comersus_backoffice_listTaxPerPlace.asp") 

    pCustomerTypeDesc	= rstemp2("customerTypeDesc")      
   end if
   %>
   <tr>
<td><%=money(pTaxPerPlace)%></td>
<td><%if pZipEq=0 and pZip<>"" then response.write "<>"%><%if pZip="" or isNull(pZip) then response.write "All"%><%=pZip%></td>
<td><%if pStateCodeEq=0 and pStateCode<>"" then response.write "<>"%><%if pStateCode="" or isNull(pStateCode) then response.write "All"%><%=pStateCode%></td>
<td><%if pCountryCodeEq=0 and pCountryCode<>"" then response.write "<>"%><%if pCountryCode="" or isNull(pCountryCode) then response.write "All"%><%=pCountryCode%></td>
<td><%=pCustomerTypeDesc%></td>
<td> <a href="comersus_backoffice_taxPerPlaceDeleteExec.asp?idTaxPerPlace=<%=pIdTaxPerPlace%>">Delete</a> </td>
</tr>         
   <%     
   rstemp.MoveNext
 loop
 
 %>
</table>
<!--#include file="footer.asp"-->

<% call closeDb() %>