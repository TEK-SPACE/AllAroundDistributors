<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' USA - 2006
' http://www.comersus.com 
' Details: Prechargeroutines
%>
<%

function preChargeApproved(pCardNumber, pExpMonth, pExpYear, pResponse)

	dim rstemp10, pIdCustomer

	pIdCustomer=getSessionVariable("idCustomer",0)
	
	pXmlHTTPComponent		= getSettingKey("pXmlHTTPComponent")
	
	' get from settings
	pPreChargeMerchant	= getSettingKey("pPreChargeMerchant")
	pPreChargeS1		= getSettingKey("pPreChargeS1")
	pPreChargeS2		= getSettingKey("pPreChargeS2")
	
	mySQL="SELECT name, lastName, zip, countryCode, phone, email FROM customers WHERE idCustomer=" & pIdCustomer

	call getFromDatabase(mySQL, rstemp10, "preCharge")
	
	if not rstemp10.eof then

	 pName		=rstemp10("name")
	 pLastName	=rstemp10("lastName")
	 pZip		=rstemp10("zip")
	 pCountryCode	=rstemp10("countryCode")
	 pPhone		=rstemp10("phone")
	 pEmail		=rstemp10("email")
	 
	end if
	
	' client connection information
	pIp		=request.ServerVariables("REMOTE_HOST")		
	
	DataToSend = "Merchant_ID=" &pPreChargeMerchant& "&Security1="&pPreChargeS1&"&Security2="&pPreChargeS2&"&Ecom_BillTo_Postal_Name_First="&pName&"&Ecom_BillTo_Postal_Name_Last="&pLastName&"&Ecom_BillTo_Postal_PostalCode="&pZip&"&Ecom_BillTo_Postal_CountryCode="&pCountryCode&"&Ecom_BillTo_Telecom_Phone_Number="&pPhone&"&Ecom_BillTo_Online_Email="&pEmail&"&Ecom_BillTo_Online_IP="&pIp&"&Ecom_Payment_Card_Number="&pCardNumber&"&Ecom_Payment_Card_ExpDate_Month="&pExpMonth&"&Ecom_Payment_Card_ExpDate_Year="&pExpYear
	
	dim xmlhttp 
	
	if pXmlHTTPComponent="Microsoft.XMLHTTP" then
	 set xmlhttp = server.Createobject("Microsoft.XMLHTTP")
	end if
	
	if pXmlHTTPComponent="Msxml2.ServerXMLHTTP.5.0" then
	 set xmlhttp = server.Createobject("Msxml2.ServerXMLHTTP.5.0")
	end if
	
	if pXmlHTTPComponent="Msxml2.ServerXMLHTTP.4.0" then
	 set xmlhttp = server.Createobject("Msxml2.ServerXMLHTTP.4.0")
	end if
	
	if pXmlHTTPComponent="Msxml2.ServerXMLHTTP" then
	 set xmlhttp = server.Createobject("Msxml2.ServerXMLHTTP")
	end if
	
	if pXmlHTTPComponent="NONE" then
	 response.redirect "comersus_message.asp?message="&Server.UrlEncode("Please configure xmlhttp component or disable this feature.")	 
	end if
		
	
	xmlhttp.Open "POST","https://api.precharge.net/charge",false
	xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	xmlhttp.send DataToSend
	Response.ContentType = "text/xml"
	pReturnCsv=xmlhttp.responseText
	Set xmlhttp = nothing
	
	pArrayReturn=split(pReturnCsv,",")
	
	if pArrayReturn(0)="response=1" then
	 preChargeApproved=-1
	else	
	
	 if pArrayReturn(0)="response=3" then
	   ' transaction error, print common errors 
	   pResponse="transaction: "&pArrayReturn(1)
	   if instr(pArrayReturn(1),"error=108")<>0 then pResponse=pResponse&" - Invalid IP"
	   if instr(pArrayReturn(1),"error=103")<>0 then pResponse=pResponse&" - Invalid Security Code"
	   if instr(pArrayReturn(1),"error=104")<>0 then pResponse=pResponse&" - Merchant Status not verified"
	   if instr(pArrayReturn(1),"error=128")<>0 then pResponse=pResponse&" - Invalid card number"
	   if instr(pArrayReturn(1),"error=117")<>0 then pResponse=pResponse&" - Invalid country"
	 end if
	 
	 if pArrayReturn(0)="response=2" then
	   ' rejected
	   pResponse="Transaction rejected"
	 end if
	
	 preChargeApproved=0
	end if

end function
%>