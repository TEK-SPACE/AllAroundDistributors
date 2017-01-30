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
<% 

on error resume next

dim mySQL, conntemp, rstemp


pIdTaxPerPlace	= getUserInput(request.Querystring("idTaxPerPlace"),10)

if trim(pIdTaxPerPlace)="" then
    response.redirect "comersus_backoffice_message.asp?message="&Server.Urlencode("Please specify tax per place id.")
end if


mySQL="DELETE FROM taxPerPlace WHERE idTaxPerPlace=" &pIdTaxPerPlace
call updateDatabase(mySQL, rstemp, "comersus_backoffice_taxPerPlaceDeleteExec.asp") 

call closeDb()

response.redirect "comersus_backoffice_addTaxPerPlaceForm.asp"
%>