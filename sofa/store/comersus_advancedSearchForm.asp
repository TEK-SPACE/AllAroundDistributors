<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com   
' Details: advanced search form page
%>

<!--#include file="../includes/screenMessages.asp"-->
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"-->
<!--#include file="../includes/currencyFormat.asp"-->
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/cartFunctions.asp"--> 
<!--#include file="../includes/adSenseFunctions.asp"-->


<%
on error resume next

dim mySQL, connTemp, rsTemp, rsTemp2

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanySlogan 		= getSettingKey("pCompanySlogan")
pAboutUsLink 		= getSettingKey("pAboutUsLink")
pStoreLocation		= getSettingKey("pStoreLocation")

pAuctions		= getSettingKey("pAuctions")
pListBestSellers 	= getSettingKey("pListBestSellers")
pNewsLetter 		= getSettingKey("pNewsLetter")
pPriceList 		= getSettingKey("pPriceList")
pStoreNews 		= getSettingKey("pStoreNews")
pOneStepCheckout	= getSettingKey("pOneStepCheckout")
pAffiliatesStoreFront   = getSettingKey("pAffiliatesStoreFront")
pAllowNewCustomer	= getSettingKey("pAllowNewCustomer")
pShowNews		= getSettingKey("pShowNews")
pRssFeedServer		= getSettingKey("pRssFeedServer")

pIdCustomer	 	= getSessionVariable("idCustomer",0)
pIdCustomerType  	= getSessionVariable("idCustomerType",1)
pIdDbSessionCart	= getSessionVariable("idDbSessionCart",0)


%>
<!--#include file="header.asp"--> 
<form name="adS" method="post" action="comersus_advancedSearchExec.asp">
  <b><%=getMsg(105,"Ad Search")%></b><br><br>
  
  <table width="500" border="0" cellspacing="0" cellpadding="0">
    <tr>     
      <td><%=getMsg(87,"Categ")%></td>
      <td>
<%

' get leaf categories from db
mySQL="SELECT idCategory, categoryDesc, idParentCategory FROM categories WHERE idCategory>1 AND active=-1"  

call getFromDatabase(mySQL, rstemp2, "advancedSearchForm")    

if rstemp2.eof then 
  response.redirect "comersus_supportError.asp?error="& Server.Urlencode("No defined categories") 
end if 
%>

<select name="idCategory">
<option value='0'><%=getMsg(94,"All")%></option>

<%
do while not rstemp2.eof

 mySQL="SELECT A.idCategory FROM categories A, categories B WHERE A.idCategory=B.idParentCategory and A.active=-1 AND A.idCategory ="& rstemp2("idCategory")
		
 call getFromDatabase(mySQL, rstemp, "advancedSearchForm")    

 if rstemp.eof then%> 
   <option value='<%=rstemp2("idCategory")%>'><%=rstemp2("categoryDesc")%></option>
 <%end if
 rstemp2.movenext
loop
%> 
 </select>
 </td>
</tr>


    <tr>     
      <td><%=getMsg(88,"Supl")%></td>
      <td>
  <%

mySQL="SELECT idSupplier, supplierName FROM suppliers"

call getFromDatabase(mySQL, rstemp, "advancedSearchForm")    

if  rstemp.eof then 
    response.redirect "comersus_admin_supporterror.asp?error="& Server.Urlencode("No suppliers defined") 
end if %> 

        <select name="idSupplier">
        <option value='0'><%=getMsg(94,"All")%></option>
        <%
 do while not rstemp.eof %>
          <option value='<%=rstemp("idSupplier")%>'><%=rstemp("supplierName")%></option>
          <%
   rstemp.movenext
loop
%></select>
      </td>
    </tr>
    <tr> 
      <td><%=getMsg(89,"Pri")%></td>
      <td>
        <input type="checkbox" name="clearance" value="-1">                
        </td>
    </tr>
    <tr> 
      <td><%=getMsg(92,"w/stock")%></td>
      <td>
        <input type="checkbox" name="withstock" value="-1">
        </td>
    </tr>
    <tr> 
      <td><%=getMsg(93,"Kword")%></td>
      <td>
        <input type=text name=keyword size=20>     
       </td>
    </tr>
    
    <tr> 
      <td colspan=2>&nbsp;
        </td>      
    </tr>
    
    
    <tr> 
      <td colspan=2>
        <input type="submit" name="Submit" value="<%=getMsg(106,"Search")%>">
        </td>      
    </tr>
    
  </table>
</form>
<!--#include file="footer.asp"--> 
<%call closeDb()%>