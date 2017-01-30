<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: enter birthday form
%>

<!--#include file="comersus_customerLoggedVerify.asp"-->
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/screenMessages.asp"--> 
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/currencyFormat.asp"-->
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/dateFunctions.asp"--> 
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/cartFunctions.asp"-->
<!--#include file="../includes/adSenseFunctions.asp"-->
<%
on error resume next

dim mySQL, connTemp, rsTemp

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pStoreLocation	 	= getSettingKey("pStoreLocation")
pHeaderKeywords 	= getSettingKey("pHeaderKeywords")

pAuctions		= getSettingKey("pAuctions")
pAllowNewCustomer	= getSettingKey("pAllowNewCustomer")
pNewsLetter 		= getSettingKey("pNewsLetter")
pStoreNews 		= getSettingKey("pStoreNews")
pSuppliersList		= getSettingKey("pSuppliersList")
pDateSwitch		= getSettingKey("pDateSwitch")
pDateFormat		= getSettingKey("pDateFormat")

pIdProduct		= getUserInput(request.QueryString("idProduct"),4)
pDescription 		= getUserInputL(request.QueryString("description"),50)

pIdCustomer     	= getSessionVariable("idCustomer",0)
pIdDbSessionCart	= getSessionVariable("idDbSessionCart",0)

pBirthDay=""

if pIdCustomer<>0 then
 ' get birthday
 mySQL="SELECT birthday FROM customers WHERE idCustomer=" &pIdCustomer
	
 call getFromDatabase(mySQL, rsTemp, "optEmailToFriendForm")    			  
 
 if not rstemp.eof then 
   pBirthDay=rstemp("birthday")
 end if
 
end if

if pDateFormat="DD/MM/YYYY" then
         
 if pDateSwitch="-1" then
  pDateFormat="MM/DD/YYYY" 
 end if
 
else

 if pDateSwitch="-1" then
  pDateFormat="DD/MM/YYYY" 
 end if
 
end if

%>

<!--#include file="header.asp"-->
<br><br><b><%=getMsg(166,"Bday")%></b>
<br><br><%=getMsg(168,"Explanation")%>
<form method="post" name="EmF" action="comersus_birthdayEditExec.asp">
    <table width="400" border="0" align="left">
      <tr> 
        <td width="40%"><%=getMsg(167,"Enter")%>&nbsp;(<%=pDateFormat%>)</td>
        <td width="60%">  
        <%if isNull(pBirthDay)=false then%>
          <input type=text name=birthday value="<%=formatDate(pBirthDay)%>">
        <%else%>
          <input type=text name=birthday value="">
        <%end if%>
        </td>
      </tr>            
      
      <tr> 
        <td colspan=2> 
          &nbsp;</td>
      </tr>            
                  
      <tr> 
        <td width="40%"> 
          <input type="submit" name="Submit" value="<%=getMsg(679,"Send")%>">
        </td>
        <td width="60%">&nbsp;</td>
      </tr>            
      
      <tr> 
        <td colspan=2> 
          &nbsp;
        </td>
      </tr>      
      
    </table>
</form>  
<br><br>
<!--#include file="footer.asp"--> 
<%call closeDb()%> 
