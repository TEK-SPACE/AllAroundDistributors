<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: send an email to a admin of the store
%>
<!--#include file="comersus_customerLoggedVerify.asp"-->
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/screenMessages.asp"--> 
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/cartFunctions.asp"-->
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
pCompanySlogan	 	= getSettingKey("pCompanySlogan")
pStoreLocation	 	= getSettingKey("pStoreLocation")
pHeaderKeywords 	= getSettingKey("pHeaderKeywords")

pAuctions		= getSettingKey("pAuctions")
pAllowNewCustomer	= getSettingKey("pAllowNewCustomer")
pNewsLetter 		= getSettingKey("pNewsLetter")
pStoreNews 		= getSettingKey("pStoreNews")
pSuppliersList		= getSettingKey("pSuppliersList")

pOrderPrefix		= getSettingKey("pOrderPrefix")

pIdCustomer     	= getSessionVariable("idCustomer",0)
pIdDbSessionCart	= getSessionVariable("idDbSessionCart",0)

' get personal info

mySQL="SELECT name, lastName, email, phone FROM customers WHERE idCustomer=" &pIdCustomer
	
call getFromDatabase(mySQL, rsTemp, "representativeStore")    			
 
if not rstemp.eof then 
   pName	=rstemp("name") & " " & rstemp("lastName")  
   pPhone	=rstemp("phone")
   pEmail	=rstemp("email")
end if

mySQL="SELECT MAX(idOrder) AS maxIdOrder FROM orders WHERE idCustomer=" &pIdCustomer
call getFromDatabase(mySQL, rsTemp, "representativeStore")    			

pIdOrder=rstemp("maxIdOrder") 

%>

<!--#include file="header.asp"-->
<br><br><b><%=getMsg(279,"Contact repr")%></b>
<form method="post" name="EmF" action="comersus_customerContactAdminExec.asp">
    <table width="400" border="0" align="left">
      <tr> 
        <td width="40%"><%=getMsg(280,"name")%></td>
        <td width="60%">  
          <input type=text name=name size=30 value="<%=pName%>">
        </td>
      </tr>
      <tr> 
        <td width="40%"><%=getMsg(281,"email")%></td>
        <td width="60%">  
          <input type=text name=email size=30 value="<%=pEmail%>">
        </td>
      </tr>      
      <tr> 
        <td width="40%"><%=getMsg(282,"phone")%></td>
        <td width="60%">  
          <input type=text name=phone size=30 value="<%=pPhone%>">
        </td>
      </tr>      
      <tr> 
        <td width="40%"><%=getMsg(283,"Order #")%></td>
        <td width="60%">  
          <input type=text name=idOrder size=10 value="<%=pOrderPrefix&pIdOrder%>">
        </td>
      </tr>            
      <tr> 
        <td width="40%"><%=getMsg(284,"Categ")%></td>
        <td width="60%">                     
	  <select name="category">
	          <option value="Sales"><%=getMsg(285,"sales")%></option>
	          <option value="Support"><%=getMsg(286,"support")%></option>
	          <option value="Shipping"><%=getMsg(287,"shipping")%></option>
	          <option value="Refund"><%=getMsg(288,"refund")%></option>
	          <option value="Other"><%=getMsg(289,"other")%></option>
          </select>
        </td>
      </tr>
      
      <tr> 
        <td width="40%"><%=getMsg(290,"msg")%></td>
        <td width="60%">           
          <textarea name=body cols="35"></textarea>
        </td>
      </tr>
      
      <tr> 
        <td width="40%"> 
          <input type="submit" name="Submit" value="<%=getMsg(291,"contact")%>">
          <br><br>
        </td>
        <td width="60%">&nbsp;</td>
      </tr>            
      
    </table>
</form>  
<br><br>
<!--#include file="footer.asp"--> 
<%call closeDb()%>