<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="../includes/settings.asp"-->    
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 

<% 
on error resume next

dim mySQL, conntemp, rstemp

pRunInstallationWizard 	= getSettingKey("pRunInstallationWizard")

if pRunInstallationWizard<>"-1" then 
 response.redirect "comersus_backoffice_message.asp?message="&Server.UrlEncode("Sorry, the installation wizard is disabled")
end if

%>
<!--#include file="header.asp"--> 
<!--#include file="./includes/settings.asp"-->

<br><font size="3"><b>Installation Wizard</b></font><br>
<br><br>Step 1 - Company Information<br><br>
<form method="post" name="install1" action="comersus_backoffice_install2.asp">
      <table width="418" border="0">
        <tr> 
          <td width="189">Company</td>
          <td width="219"> 
            <input type="text" name="pCompany" value="">
          </td>
        </tr>
        <tr> 
          <td width="189">Address</td>
          <td width="219"> 
            <input type="text" name="pCompanyAddress" value="">
          </td>
        </tr>
        <tr> 
          <td width="189">City</td>
          <td width="219"> 
            <input type="text" name="pCompanyCity" value="">
          </td>
        </tr>
        
         <tr> 
          <td width="189">Zip</td>
          <td width="219"> 
            <input type="text" name="pCompanyZip" value="">
          </td>
        </tr>
        
        <tr> 
          <td width="189">State Code</td>
          <td width="219"> 
            
            <%
      ' get state codes
      mySQL="SELECT stateCode, stateName FROM stateCodes ORDER BY stateName"

      call getFromDatabase(mySQL, rstemp, "orderForm")
      
      %> 
            <SELECT name="pCompanyStateCode">            
              <%do while not rstemp.eof
      pStateCode2=rstemp("stateCode")%> 
              <option value="<%=pStateCode2%>"><%=rstemp("stateName")%> </OPTION>
      <%rstemp.movenext
      loop%> 
      <option value="">State Code not listed here</OPTION>
            </SELECT>
            
            
          </td>
        </tr>
                      
        <tr> 
          <td width="189">State Code not listed</td>
          <td width="219"> 
            <input type="text" name="pCompanyStateCode" value="">
          </td>
        </tr>                 
        
        <tr> 
          <td width="189">Country</td>
          <td width="219"> <%
      ' get CountryCodes
      mySQL="SELECT countryCode, countryName FROM countryCodes ORDER BY countryName"

      call getFromDatabase(mySQL, rstemp, "orderForm")
      
      %> 
            <SELECT name="pCompanyCountryCode">
              <%do while not rstemp.eof
      pCountryCode2=rstemp("countryCode")%> 
              <option value="<%=pCountryCode2%>"<%
        if pCountryCode2="US" then
            response.write "selected"           
        end if
        %>><%=rstemp("countryName")%> <%rstemp.movenext
      loop%> </OPTION>
            </SELECT>
          </td>
        </tr>
        <tr> 
          <td width="189">Phone</td>
          <td width="219"> 
            <input type="text" name="pCompanyPhone" value="">
          </td>
        </tr>
        <tr> 
          <td width="189">Fax</td>
          <td width="219"> 
            <input type="text" name="pCompanyFax" value="">
          </td>
        </tr>       
        <tr> 
          <td width="189">Company Slogan</td>
          <td width="219"> 
            <input type="text" name="pCompanySlogan" value="">
          </td>
        </tr>             
        <tr> 
          <td colspan="2"> <br>
            <input type="submit" name="Submit" value="Continue">
          </td>
        </tr>
      </table>    
</form>
<%call closeDb()%>
<!--#include file="footer.asp"-->