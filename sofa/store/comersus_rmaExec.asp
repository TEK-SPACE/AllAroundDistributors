<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: save on the database the new RMA issue.
%>

<!--#include file="comersus_customerLoggedVerify.asp"-->   
<!--#include file="../includes/settings.asp"-->    
<!--#include file="../includes/screenMessages.asp"-->    
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->  
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/dateFunctions.asp"--> 
<!--#include file="../includes/encryption.asp"-->    
<!--#include file="../includes/sendMail.asp"--> 
<!--#include file="comersus_optDes.asp"-->    
<!--#include file="../includes/stringFunctions.asp"-->

<%
' on error resume next

dim connTemp, rsTemp, mySql

' get settings

pStoreFrontDemoMode 		= getSettingKey("pStoreFrontDemoMode")
pDateFormat			= getSettingKey("pDateFormat")
pEmailSender			= getSettingKey("pEmailSender")
pEmailAdmin			= getSettingKey("pEmailAdmin")
pSmtpServer			= getSettingKey("pSmtpServer")
pEmailComponent			= getSettingKey("pEmailComponent")
pDebugEmail			= getSettingKey("pDebugEmail")


' get from session
pIdCustomer     	= getSessionVariable("idCustomer","0")

pDate			= fixDate(Date())

' form data
pIdOrder		= getUserInput(request.form("idOrder"),50)
pCustomerReasons	= getUserInput(request.form("reasons"),500)

if pIdOrder="" or pCustomerReasons="" then
  response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(570,"Enter all fields"))
end if

' insert RMA record
mySQL="INSERT INTO RMA (idOrder, rmaDate, customerReasons, rmaStatus) VALUES ("& pIdOrder & ",'" & pDate  & "','" & pCustomerReasons &"',1)"
call updateDatabase(mySQL, rsTemp, "comersus_ramExec.asp")    

' send admin notification
call sendmail (pCompany, pEmailSender, pEmailAdmin, "RMA order"&pIdOrder, "A new RMA has been requested")

response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(737,"RMA sent"))

call closeDb()
%>