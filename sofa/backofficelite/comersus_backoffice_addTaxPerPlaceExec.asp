<%
' Comersus BackOffice 
' e-commerce ASP Open Source
' CopyRight Comersus Open Technologies LC
' United States - 2005
' http://www.comersus.com 
%>
<!--#include file="comersus_backoffice_adminVerify.asp"-->  
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringFunctions.asp"-->
<!--#include file="../includes/sessionFunctions.asp"-->

<% 

on error resume next

dim mySQL, conntemp, rstemp

pIdStoreBackOffice	= getSessionVariable("idStoreBackOffice", 1)

pTaxPerPlace		= getUserInput(request.querystring("taxPerPlace"),50)

pZip			= getUserInput(request.querystring("zip"),20)
pZipEq	 		= getUserInput(request.querystring("zipEq"),10)

pStateCode		= getUserInput(request.querystring("stateCode"),20)
pStateCodeEq 		= getUserInput(request.querystring("stateCodeEq"),10)

pCountryCode		= getUserInput(request.querystring("countryCode"),20)
pCountryCodeEq		= getUserInput(request.querystring("countryCodeEq"),10)

pIdCustomerType		= getUserInput(request.querystring("idCustomerType"),10)

if pZipEq="" then
 pZipEq="0"
end if

if pStateCodeEq="" then
 pStateCodeEq="0"
end if

if pCountryCodeEq="" then
 pCountryCodeEq="0"
end if

' insert tax per place in to db
if pIdCustomerType<>"" then
  mySQL="INSERT INTO taxPerPlace (taxPerPlace, zip, zipEq, stateCode, stateCodeEq, countryCode, countryCodeEq, idCustomerType, idStore) VALUES (" &pTaxPerPlace& ",'" &pZip& "'," &pZipEq& ",'" &pStateCode& "'," &pStateCodeEq& ",'" &pCountryCode& "'," &pCountryCodeEq& "," &pIdCustomerType& "," &pIdStoreBackOffice& ")"
else
 mySQL="INSERT INTO taxPerPlace (taxPerPlace, zip, zipEq, stateCode, stateCodeEq, countryCode, countryCodeEq, idStore) VALUES (" &pTaxPerPlace& ",'" &pZip& "'," &pZipEq& ",'" &pStateCode& "'," &pStateCodeEq& ",'" &pCountryCode& "'," &pCountryCodeEq& "," &pIdStoreBackOffice& ")"
end if

call updateDatabase(mySQL, rstemp, "comersus_backoffice_addTaxPerPlaceExec.asp") 

call closeDb()

response.redirect "comersus_backoffice_addTaxPerPlaceForm.asp"
%>
