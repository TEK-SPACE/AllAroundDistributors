<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com 
' Details: redirect to index page
%>
<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/dateFunctions.asp"--> 
<!--#include file="../includes/miscFunctions.asp"--> 
<!--#include file="../includes/stringFunctions.asp"--> 
<!--#include file="../includes/referrerParsing.asp"--> 

<%
on error resume next

call saveCookie()

dim connTemp, rsTemp, mySQL

' get settings 

pIndexVisitsCounter 	= getSettingKey("pIndexVisitsCounter")
pDateFormat 		= getSettingKey("pDateFormat")
pRunInstallationWizard  = getSettingKey("pRunInstallationWizard")

if pRunInstallationWizard="-1" then 
 response.redirect "../backofficeLite/comersus_backoffice_install0.asp"
end if

' init all variables needed

call sessionInit()
 
session("idDbSession")	= checkSessionData()

' set affiliate based on querystring value
if request.querystring("idAffiliate")<>"" and isNumeric(request.querystring("idAffiliate"))then
   session("idAffiliate")= request.querystring("idAffiliate")	
end if

 ' increase visits for BackOffice
if pIndexVisitsCounter="-1" then  
 
 ' visits counter
 pUserIp	 = getUserInput(request.ServerVariables("REMOTE_HOST"),20)
 pReferrer	 = getUserInput(request.ServerVariables("HTTP_Referer"),100)
 pVisitDate 	 = Date()
 pVisitTime	 = Time
 pSearchEngine   = ""
 pSearchKeywords = ""
 
 call referrerParsing(pReferrer, pSearchEngine, pSearchKeywords)
  
 mySQL="INSERT INTO visits (userIp, referrer, keywords, visitDate, visitTime, idStore) VALUES ('"&pUserIp& "','" &pReferrer& "','" &pSearchKeywords& "','" &pVisitDate& "','" &pVisitTime& "'," &pIdStore& ")"
 
 call updateDatabase(mySQL, rsTemp, "default")    
 
end if

call closeDb()

response.redirect "comersus_index.asp"
%>




