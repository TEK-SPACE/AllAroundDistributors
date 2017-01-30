<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/stringFunctions.asp"-->    
<!--#include file="../includes/itemFunctions.asp"-->    
<!--#include file="../includes/currencyFormat.asp"-->    
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/settings.asp"-->    

<% 
on error resume next 

dim mySQL, conntemp, rstemp, rsTemp1

pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pMoneyDontRound	 	= getSettingKey("pMoneyDontRound")

' get data of the tax to modify
mySQL="SELECT * FROM taxPerPlace WHERE countryCodeEq=-1 OR stateCodeEq=-1 AND idStore=" & pIdStore
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_taxesModifyForm.asp") 

If rsTemp.eof then
	pTaxPerPlaceLocal=0
	pStateCode=""
	pCountryCode=""
else
	pTaxPerPlaceLocal=rsTemp("taxPerPlace")
	pStateCode=rsTemp("stateCode")
	pCountryCode=rsTemp("countryCode")
end if	


' get data of the tax2 to modify
mySQL="SELECT * FROM taxPerPlace WHERE countryCodeEq=0 AND stateCodeEq=0 AND idStore=" & pIdStore
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_taxesModifyForm.asp") 

If rsTemp.eof then
	pTaxPerPlaceNotLocal=0
else
	pTaxPerPlaceNotLocal=rsTemp("taxPerPlace")
end if	

%>

<!--#include file="header.asp"-->
<!--#include file="./includes/settings.asp"-->

<br><b>Edit taxes</b><br>
<form method="post" name="modify" action="comersus_backoffice_taxesModifyExec.asp">
  <table width="100%" border="0">
    <tr> 
      <td width="269">Enter the tax percentage for customers located in your State and/or Country: (example 35.5%)</td>
      <td width="139">                
				<input type="text" name="taxValueSameCountry" value="<%=money(pTaxPerPlaceLocal)%>" size=5>% 
      </td>
    </tr>
    <tr> 
      <td width="269">Enter the tax percentage for customers located outside your State and/or Country:</td>
      <td width="139">                
        <input type="text" name="taxValueDifferentCountry" value="<%=money(pTaxPerPlaceNotLocal)%>" size=5>%
      </td>
    </tr>    
    <tr> 
      <td width="269">Select your country</td>
      <td width="139">                
	      <% ' get CountryCodes
	      mySQL="SELECT countryCode, countryName FROM countryCodes ORDER BY countryName"	
	      call getFromDatabase(mySQL, rstemp, "taxes")	      
	      %>
	      <SELECT name="taxCountryCode">
	      <OPTION value="">Select
	      <%do while not rstemp.eof%>      
	        <option value="<%=rstemp("countryCode")%>" <%If rstemp("countryCode")=pCountryCode Then response.write "selected"%>><%=rstemp("countryName")%>
	      <%rstemp.movenext
	      loop%>
	      </OPTION>
      	      </SELECT>
      </td>
    </tr>        
    <tr> 
      <td width="269">Select your state</td>
      <td width="139">                
	      <%' get stateCodes
	      mySQL="SELECT stateCode, stateName FROM stateCodes ORDER BY stateName"	      
	      call getFromDatabase(mySQL, rstemp, "orderForm")   %>
	      <SELECT name=taxStateCode size=1>
	      <OPTION value="">Select
	      <%do while not rstemp.eof%>	      
	        <option value="<%=rstemp("stateCode")%>" <%If rstemp("stateCode")=pStateCode Then response.write "selected"%>><%=rstemp("stateName")%>
	      <%rstemp.movenext
	      loop%>
	      </OPTION>
      	      </SELECT>
      </td>
    </tr>
    <tr><td><input type="submit" name="modify" value="Modify"></td>
    <td></td></tr>
    <tr><td colspan=2> </td></tr>
  </table>
<!--#include file="footer.asp"-->
</form>