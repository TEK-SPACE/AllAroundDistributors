<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' USA - 2006
' http://www.comersus.com 
' Details: ChargebackProtection routines
%>
<%
function checkChargeProtection(pChargebackProtectionMerchant, pChargebackProtectionPassword, pName, pLastName, pCountryCode)

    DataToSend="idMerchant="&pChargebackProtectionMerchant&"&merchantPassword="&pChargebackProtectionPassword
    DataToSend=DataToSend&"&Name="&pName&"&lastName="&pLastName&"&countryCode="&pCountryCode
    
    set xmlhttp = server.Createobject("Microsoft.XMLHTTP")
    
    xmlhttp.Open "GET","http://www.chargebackprotection.org/api/default.asp?"& DataToSend,false
    xmlhttp.send 
    strRetval=xmlhttp.responseText

    set xmlhttp=nothing
    pArray = split(strRetval,",")
    if Not IsArray(pArray) then
        'without Values
        pChargeBackMSG = "Error: Service off line." 
    else
        if trim(pArray(1)) = "0" then
             'error
            pChargeBackMSG = "Error:" & pArray(3)
        else
            pChargeBackMSG = pArray(2)
        end if
    end if
    checkChargeProtection = pChargeBackMSG 
    
end function    
%>    