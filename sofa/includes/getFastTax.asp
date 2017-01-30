<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' USA - 2006
' http://www.comersus.com 
' Details: Get Sales Tax in Real Time with http://www.serviceobjects.com
%>

<%

Dim SoapServer, SoapPath, SoapAction, SoapNamespace

' Service Objects SOAP settings
SoapServer		= "ws2.serviceobjects.net"
SoapPath		= "/ft/FastTax.asmx"
SoapAction		= "GetTaxInfo"
SoapNamespace		= "http://www.serviceobjects.com/"

Function getLiveSalesTax(strPostalCode, strKey)

 dim SoapBody, strTaxType

 strTaxType="SALES"

 if((strPostalCode <> "") and (strTaxType <> "") and (strKey <> "")) then 	
	SoapBody = xmlSoap(strPostalCode, strTaxType, strKey)
 end if
	
 if SoapBody <> "" then
  dim xml
 
  'on error resume next
 
  set xml = Server.CreateObject("Microsoft.XMLDOM")
  if(err.number = 0) then		
	
		xml.async = False
		xml.loadxml(SoapBody)

		dim oTopNode : set oTopNode = xml.selectSingleNode("soap:Envelope/soap:Body/" & SoapAction & "Response/" & SoapAction & "Result")

		if(TypeName(oTopNode.selectSingleNode("TotalTaxRate")) = "IXMLDOMElement") then 
			dim passTopNode
			passTopNode = oTopNode.selectSingleNode("TotalTaxRate").nodeTypedValue
		end if 
		
		set xml = nothing		
 end if 'DOM object is valid
end if 
	getLiveSalesTax = passTopNode
end Function


function xmlSoap(strPostalCode, strTaxType, strKey)

pXmlHTTPComponent		= getSettingKey("pXmlHTTPComponent")
	
Dim xmlhttp, strSoap  
 
if pXmlHTTPComponent="Microsoft.XMLHTTP" then
 set xml = server.Createobject("Microsoft.XMLHTTP")
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
  
 ' Build XML document:
 
 strSoap =	"<?xml version=""1.0"" encoding=""utf-8""?>" & _
		"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
		"<soap:Body>" & _
		"<h:" & SoapAction & " xmlns:h=""" & SoapNamespace & """>" & _
		"<h:PostalCode>" & strPostalCode & "</h:PostalCode>" & _
		"<h:TaxType>" & strTaxType & "</h:TaxType>" & _
	        "<h:LicenseKey>" & strKey & "</h:LicenseKey>" & _
		"</h:" & SoapAction & ">" & _
		"</soap:Body>" & _
		"</soap:Envelope>"

	' HTTP header:
	xmlhttp.Open "POST", "http://" & SoapServer & SoapPath, False	' False = Do not respond immediately
	xmlhttp.setRequestHeader "Man", "POST " & SoapPath & " HTTP/1.1"
	xmlhttp.setRequestHeader "Host", SoapServer
	xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	xmlhttp.setRequestHeader "SOAPAction", SoapNamespace & SoapAction

	' Send request
	xmlhttp.send(strSoap)
	
	if xmlhttp.Status = 200 then			' Response from server was success
		xmlSoap = xmlhttp.responseText
	else									' Response from server failed
		xmlSoap = ""
		' Tell administrator what went wrong - maybe not users though
		'Response.Write("Server Error...<br>")
		'Response.Write("status = " & xmlhttp.status)
		'Response.Write("<br>" & xmlhttp.statusText)
		'Response.Write("<br><pre>" & Request.ServerVariables("ALL_HTTP") & "</pre>")
	end if

	Set xmlhttp = nothing
	
end function
%>