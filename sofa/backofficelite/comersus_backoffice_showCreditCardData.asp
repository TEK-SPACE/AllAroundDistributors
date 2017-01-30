<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"-->  
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/settings.asp"-->    
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->   
<!--#include file="../includes/stringFunctions.asp"-->   
<!--#include file="../includes/databaseFunctions.asp"-->   
<% 

on error resume next 

dim mySQL, conntemp, rstemp

pOrderPrefix		= getSettingKey("pOrderPrefix")

pIdOrder 		= getUserInput(request.querystring("idOrder"),12)

if pIdOrder="" then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Enter a valid order number")
end if

mySQL="SELECT * FROM creditCards WHERE idOrder="&pIdOrder
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_showcreditcarddata.asp") 

if rstemp.eof then 
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("There is no credit card information entered for this order")
end if       

pEncryptedCardNumber 	= rstemp("cardNumber")
pCardType		= rstemp("cardType") 
pExpiration		= rstemp("expiration")
pSeqCode		= rstemp("seqcode")
pObs			= rstemp("obs")

%> 

<!--#include file="header.asp"-->
<!--#include file="includes/settings.asp"-->   
<br><b>View credit card details</b><br>
<table width="590" border="0">
  <tr bgcolor="#CCCCCC"> 
    <td height="14" colspan="3"><img src="images/smallIcoBackOffice2.gif" width="11" height="11"> 
      Order <%=pOrderPrefix&pIdOrder%></td>
  </tr>
  <tr> 
    <td colspan="2" height="8">Card type</td>
    <td height="8" width="449">
    <img src="images/creditcards_<%=pCardType%>.gif"> 
    </td>    
  </tr>
  <tr> 
    <td height="11" colspan="2">Number (*)</td>
    <td height="11" width="449"><%=pEncryptedCardNumber%>    
    </td>    
  </tr>  
  <tr> 
    <td colspan="2" height="21">Expiration</td>
    <td height="21" width="449"><%=pExpiration%></td>
  </tr>
  <tr> 
    <td colspan="2" height="25">CVV2 code</td>
    <td height="25" width="449"> <%    
    if trim(pSeqCode)="" then
     response.write "-"
    else
     response.write pSeqCode
    end if
    %></td>
  </tr>  
  <tr> 
    <td colspan="2" height="25">Comments</td>
    <td height="25" width="449">     
    <%=pObs%>
    </td>
  </tr>
  
  <tr><td>&nbsp;</td></tr>
  <tr>
   <td colspan="3">
   <i>(*) TIP: To decrypt the credit card info copy the encrypted number and paste into Utilities/Encryption.
Power Pack Medium and Premium includes one step function to retrieve Credit Card information. </i>
   </td>
  </tr>
            
</table>
<br>
<%call closeDb()%> 
<!--#include file="footer.asp"-->
