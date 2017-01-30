<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"-->  
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/currencyFormat.asp"-->    
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/getSettingKey.asp"-->  
<!--#include file="../includes/stringFunctions.asp" --> 

<%
on error resume next

dim mySQL, conntemp, rstemp

pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")


pIdZoneContent	= getUserInput(request("idZoneContent"),10)
pIdZone			= getUserInput(request("idZone"),10)

If getUserInput(request.form("submit"),20) <> "" Then
	
	MySQL	= "DELETE FROM shippingZonesContents WHERE idShippingZonesContents=" & pIdZoneContent
	call updateDatabase(mySQL, rstemp, "comersus_backoffice_deleteZonesContents.asp") 
	response.redirect "comersus_backoffice_listZoneAndContent.asp?IdZone="&pIdZone
end if


%>

<!--#include file="header.asp"-->
<!--#include file="includes/settings.asp"--> 

<form action='<%=request.ServerVariables("SCRIPT_NAME")%>'	 method=post>
   <input type=hidden name="idZoneContent" value=<%=pIdZoneContent%>>	
   <input type=hidden name="idZone" value=<%=pIdZone%>>	

	<B>Attention</B><br>
	Delete the Zone?
	<br><br><br>
	<input type=button value=Cancel name=cancel onclick="location.href='comersus_backoffice_listZoneAndContent.asp?idZone=<%=pIdZone%>'">&nbsp;&nbsp;&nbsp;<input type=submit value=Delete name=submit>
</form>
<!--#include file="footer.asp"-->
