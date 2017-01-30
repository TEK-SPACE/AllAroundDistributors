<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: send customer msg to representative
%>
<!--#include file="comersus_customerLoggedVerify.asp"-->
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/screenMessages.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/sendMail.asp"-->

<%
on error resume next

dim mySQL, connTemp, rsTemp

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCompany	 	= getSettingKey("pCompany")

pEmailSender		= getSettingKey("pEmailSender")
pEmailAdmin		= getSettingKey("pEmailAdmin")
pSmtpServer		= getSettingKey("pSmtpServer")
pEmailComponent		= getSettingKey("pEmailComponent")
pDebugEmail		= getSettingKey("pDebugEmail")

' user form data
pEmail			= request.form("email")
pName			= request.form("name")
pBody 			= request.form("body")
pIdOrder		= request.form("idOrder")
pPhone			= request.form("phone")
pCategory		= request.form("category")

if pBody="" then
 response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(570,"Enter all fields"))
end if

call sendmail (pName, pEmailSender, pEmailAdmin,"Contact from the Store, customer "&pName,"Order #: "&pIdOrder&VBcrlf&"Phone:  "&pPhone&VBcrlf&"Category: "&pCategory&VBcrlf&pBody&VBcrlf&VBcrlf)

response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(292,"Thanks, your message was sent."))

call closeDb()
%>
