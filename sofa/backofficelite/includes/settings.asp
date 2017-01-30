<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' CopyRight Rodrigo S. Alhadeff for Comersus
' Sep-2002
' http://www.comersus.com 
' Details: backoffice settings
%>

<%
dim pCanChangePassword, pStoreFrontSettingLines

pCanChangePassword	= -1
pLicensedTo		= "Final User, Open Source License"

' Note: you must have off line payment package to use this feature
pUseComersusOLPayment	= -1
pUseSslOLPayment	= 0
pModifyGenericSql	= -1
pUseDHTMLEditor         = 0

pBackOfficeDemoMode	= 0
pChangeDecimalPoint	= 0

pPathForUpload		= "\comersus\store\catalog"
%>