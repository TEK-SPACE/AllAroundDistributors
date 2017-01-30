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
 response.redirect "comersus_backoffice_message.asp?message="&Server.UrlEncode("Sorry, the Installation Wizard is disabled")
end if

' save settings into db

for each field in request.form 
 fieldValue = request.form(field)
   
 if fieldValue<>"" and field<>"Submit" and field<>"pStateCodes" then
  mySQL1="UPDATE settings SET settingValue='" &fieldValue& "' WHERE settingKey='" &field& "' AND idStore=" &pIdStore
  call updateDatabase(mySQL1, rstemp, "comersus_backoffice_install.asp") 
 end if
 	if field="pStateCodes" Then
 		' load state codes
 		if fieldValue<>"" Then
 				arrFields = split(fieldValue, chr(10))
 				
				     for i=0 to ubound(arrFields)
				      
				      if len(arrFields(i)) > 1 then
				       
				       pPos = instr(arrFields(i)," ")
				       pStateCode = left(arrFields(i), pPos-1)
				       pStateName = mid(arrFields(i), pPos+1)       
				       if pStateCode <> "" and pStateName <> "" then
				        mySQL1="INSERT INTO stateCodes(stateCode, stateName) VALUES ('"&pStateCode&"', '"&pStateName&"')"
				        call updateDatabase(mySQL1, rstemp, "comersus_backoffice_install.asp")
				       end if
				      
				      end if
				      
				     next
 				
 		end if
 		
	end if
next 

%>
<!--#include file="header.asp"--> 
<!--#include file="./includes/settings.asp"-->

<br><font size="3"><b>Installation Wizard</b></font><br>
<br>Step 3 - Cart Behaviour<br><br>
<form method="post" name="install3" action="comersus_backoffice_install4.asp" >
<table width="418" border="0">
    <tr> 
      <td width="189">Control product stock before adding to the cart</td>
      <td width="219">        
        <select name="pUnderStockBehavior">
        <option value='dontadd'>Yes</option>
        <option value='none' selected>No</option>
        </select>
      </td>
    </tr>    
    <tr> 
      <td width="189">Minimum purchase amount</td>
      <td width="219">        
      <select name="pMinimumPurchase">        
        <option value='0'>No minimum</option>
        <option value='5'>$5</option>
        <option value='10'>$10</option>
        <option value='50'>$50</option>
        <option value='100'>$100</option>
        <option value='500'>$500</option>
        <option value='1000'>$1000</option>
       </select>
        
      </td>
    </tr>
    <tr> 
      <td width="189">Allow new registration of customers</td>
      <td width="219">        
        <select name="pAllowNewCustomer">
        <option value='-1'>Yes</option>
        <option value='0'>No</option>
        </select>
      </td>      
      <tr> 
      <td width="189">Auto generate passwords for customers</td>
      <td width="219">        
        <select name="pRandomPassword">
        <option value='-1'>Yes</option>
        <option value='0' selected>No</option>
        </select>
      </td>
    </tr>    
    <tr> 
      <td width="189">Show inventory to the customer</td>
      <td width="219">        
      <select name="pShowStockView">
        <option value='-1'>Yes</option>
        <option value='0' selected>No</option>
      </select>      
      </td>
    </tr>
    
    <tr> 
      <td width="189">Items to display in store home</td>
      <td width="219">        
      <select name="pItemsShown">
        <option value='2'>2</option>
        <option value='4' selected>4</option>
        <option value='6'>6</option>
        <option value='8'>8</option>
      </select>              
      </td>
    </tr>
    <tr> 
      <td width="189">Custom field number 1 for each order posted (1)</td>
      <td width="219">        
        <input type="text" name="orderFieldName1" value="">
      </td>
    </tr>    
    <tr> 
      <td width="189">Custom field number 2 for each order posted (1)</td>
      <td width="219">        
        <input type="text" name="orderFieldName2" value="">
      </td>
    </tr>
    <tr> 
      <td width="189">Custom field number 3 for each order posted (1)</td>
      <td width="219">        
        <input type="text" name="orderFieldName3" value="">
      </td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <br>
        <input type="submit" name="Submit" value="Continue">
      </td>
    </tr>
</table>    
</form>

<br>Additional information 
<br>(1) Enter here the caption for custom fields to be included in order checkout. 
<br>Examples: date to receive the order, where did you hear about us, etc 
<br>Leave the fields blank if you don't want to use custom fields at order checkout 
<%call closeDb()%>
<!--#include file="footer.asp"-->