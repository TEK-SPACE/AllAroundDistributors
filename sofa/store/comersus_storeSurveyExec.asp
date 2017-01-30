<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: send the survey results by email 
%>

<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/screenMessages.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/sendMail.asp"-->
<!--#include file="../includes/currencyFormat.asp"-->
<!--#include file="../includes/dateFunctions.asp"--> 

<%
on error resume next

dim mySQL, connTemp, rsTemp

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCompany	 	= getSettingKey("pCompany")
pDateFormat		= getSettingKey("pDateFormat")

pEmailSender		= getSettingKey("pEmailSender")
pEmailAdmin		= getSettingKey("pEmailAdmin")
pSmtpServer		= getSettingKey("pSmtpServer")
pEmailComponent		= getSettingKey("pEmailComponent")
pDebugEmail		= getSettingKey("pDebugEmail")
pVerificationCodeEnabled= getSettingKey("pVerificationCodeEnabled")

pVerificationCode	= getUserInput(request.form("verificationCode"),8)
pRealBoxCode    	= getSessionVariable("boxCode","")

if pVerificationCodeEnabled="-1" then
 if pRealBoxCode<>pVerificationCode then
  response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(762,"Invalid"))
 end if
end if

' create date
pDate			= formatDate(Date())


pBody=getMsg(379,"Survey of ") &pDate&Vbcrlf&Vbcrlf

for each field in request.form 
  
  pValue = request.form(field) 
  
  select case field
  
   case "1"
    pQuestion=getMsg(339,"How many times have you visited our store before, approximately?")       
   case "2"
    pQuestion=getMsg(340,"You are a")    
   case "3"
    pQuestion=getMsg(351,"How did you learn about us?")
   case "4"
    pQuestion=getMsg(356,"Have you found what you were looking for?")
   case "5"
    pQuestion=getMsg(359,"Did you find the store easy to navigate?")
   case "6"
    pQuestion=getMsg(360,"Which of the following features are of interest to you?")
   case "7"
    pQuestion=getMsg(367,"Did you find the available payment methods convenient?")
   case "8"
    pQuestion=getMsg(368,"Did you find the available shipping methods convenient?")
   case "9"
    pQuestion=getMsg(369,"Regarding prices")
   case "10"
    pQuestion=getMsg(373,"Are you concerned about privacy and security when shopping online?")
   case "11"
    pQuestion=getMsg(374,"Would you rather place your order online but post payment by phone afterwards?")
   case "12"
    pQuestion=getMsg(375,"Would you consider making another purchase soon if you receive a discount code?")
   case "13"
    pQuestion=getMsg(376,"Would you be interested in becoming an affiliate and earning commissions on sales?")
   case "14"
    pQuestion=getMsg(377,"Do you have any suggestions to improve the store?")   
  end select 
  
  pBody=pBody&Vbcrlf&pQuestion&": "&pValue
  
next  
  
' send email
call sendmail (pName, pEmailSender, pEmailAdmin, getMsg(379,"Survey of ")&pDate, pBody&VBcrlf&VBcrlf)

response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(380,"Thanks for answering this survey.")) 

call closeDb()
%>
