<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: show error screen and send email to admin
%>

<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->  
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/sendMail.asp"-->
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/adSenseFunctions.asp"-->

<%

on error resume next

' replace by values of settings to avoid DB retrieve

pEmailSender		= pSupportErrorEmailFrom
pEmailAdmin		= pSupportErrorEmailFrom
pSmtpServer		= pSupportErrorSMTP
pEmailComponent		= pSupportErrorEmailComponent

pError 			= getQueryString(request.querystring("error"),500)

' check for common errors
if instr(pError,"updateable")<>0 then
 pAdditionalInformation = "Additional information: it seems that your database folder (comersus/database) or database file (comersus/database/comersus.mdb) has not defined all permissions. It's possible also that your database file has Read-Only attribute checked. Sometimes this error appears if your ODBC driver is outdated. Read installation instructions. You can also get detailed instructions for some operating systems at the forum."
end if

' check for common errors
if instr(pError,"Unspecified")<>0 then
 pAdditionalInformation = "Additional information: this error may be caused for different reasons. Most of the times this error is caused for wrong permissions for anonymus user."
end if

' check for common errors
if instr(pError,"ActiveX component")<>0 then
 pAdditionalInformation = "Additional information: this error may be caused when some DLLs are missing or have the wrong path. You can see more information at Microsoft site."
end if

' check for common errors
if instr(pError,"invalid pog")<>0 then
 pAdditionalInformation = "Additional information: it seems that you have enabled one feature that needs a component that s not installed in your server."
end if

' check for common errors
if instr(pError,"Data type mismatch")<>0 then
 pAdditionalInformation = "Additional information: this error may be caused when your server cannot use month names. Usually for foreign OS installations. Open comersus/includes/settings.asp and change DatabaseDateMonthName to 0"
end if

call sendmail (pCompany, pEmailSender, pEmailAdmin, "Temporary error in your Comersus Store", "Support, " &VBCrlf& "There are errors in your Comersus Store. "&VBCrlf& "Error: "&pError &Vbcrlf& "Time: "&Time &Vbcrlf& "Logged idCustomer:" &session("idCustomer") &Vbcrlf&"Additional Information:"&pAdditionalInformation )

' PayPal Token
session("token")=""

%>

<HTML>
<br><br><b>Temporary error</b>
<br><br>Sorry for the inconvenience. 

<%' show only if demo mode
if pSupportErrorShowDetails=-1 then%>
 <%=pError%>
 <br><%=pAdditionalInformation%> 
 <br><br>You can request free technical assistance at 
 <a href="http://www.comersus.org/forum">Comersus Forum</a>
<%else
' clear sessionData
 session("IdDbSession")		=""
 session("IdDbSessionCart")	=""
 ' clear header cart vars
 session("cartSubTotal") 	= 0
 session("cartItems")		= 0
%>
 <br>Please come back later or contact us to order by email or phone
<%end if%>
</HTML>
