<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"--> 
<!--#include file="./includes/settings.asp"-->
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/stringFunctions.asp"-->    
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/settings.asp"-->    

<% 
on error resume next 

dim mySQL, conntemp, rstemp

 ' Modify taxes
pTaxValueSameCountry	   =request.form("taxValueSameCountry")
pTaxValueDifferentCountry  =request.form("taxValueDifferentCountry")
pTaxCountryCode		   =request.form("taxCountryCode")
pTaxStateCode		   =request.form("taxStateCode")

if pTaxValueSameCountry<>"" or pTaxValueDifferentCountry<>"" then

	mySQL="DELETE FROM taxPerPlace WHERE idStore=" &pIdStore
	call updateDatabase(mySQL, rstemp, "comersus_backoffice_install6.asp") 
	
	mySQL="INSERT INTO taxPerPlace (idStore, taxPerPlace,"	
	
	if pTaxCountryCode<>"" then
		mySQL= mySQL & "countryCode, countryCodeEq,"
	end if	
	if pTaxStateCode<>"" then
		mySQL= mySQL & "stateCode, stateCodeEq,"
	end if
	mySQL1= mySQL & " idCustomerType) VALUES (" &pIdStore & "," &pTaxValueDifferentCountry& ","
	mySQL= mySQL & " idCustomerType) VALUES  (" &pIdStore & ","  &pTaxValueSameCountry& ","	
	if pTaxCountryCode<>"" then
		mySQL= mySQL & "'" &pTaxCountryCode& "',-1,"
		mySQL1= mySQL1 & "'" &pTaxCountryCode& "',0,"
	end if			
	if pTaxStateCode<>"" then
		mySQL= mySQL & "'" &pTaxStateCode& "',-1,"
		mySQL1= mySQL1 & "'" &pTaxStateCode& "',0,"
	end if	
	mySQL= mySQL & "1)"
	mySQL1= mySQL1 & "1)"
end if	
if pTaxValueSameCountry<>"" then
	call updateDatabase(mySQL, rstemp, "comersus_backoffice_install6.asp") 
end if
if pTaxValueDifferentCountry<>"" and mySQL<>mySQL1 then 
	call updateDatabase(mySQL1, rstemp, "comersus_backoffice_install6.asp") 
end if

 response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Taxes modified")

%>
